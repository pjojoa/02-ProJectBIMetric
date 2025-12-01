'use strict';

import powerbi from 'powerbi-visuals-api';
import { FormattingSettingsService } from 'powerbi-visuals-utils-formattingmodel';
import '../style/visual.less';

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import DataView = powerbi.DataView;

import { VisualSettingsModel } from './settings';
import { initializeViewerRuntime, loadModel, IdMapping, isolateDbIds, fitToView, showAll } from './viewer.utils';
import * as models from 'powerbi-models';

/**
 * Custom visual wrapper for the Autodesk Platform Services Viewer.
 */
export class Visual implements IVisual {
    // Visual state
    private host: IVisualHost;
    private statusDiv: HTMLDivElement;
    private container: HTMLElement;
    private formattingSettings: VisualSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private currentDataView: DataView = null;
    private selectionManager: ISelectionManager = null;
    private isProgrammaticSelection: boolean = false;

    // Visual inputs
    private accessTokenEndpoint: string = '';

    private currentUrn: string = "";
    private currentGuid: string | null = null;
    private viewer: Autodesk.Viewing.GuiViewer3D;
    private isViewerReady: boolean = false;

    // Map SelectionId Key -> { id: string, color: string | null, selectionId: ISelectionId }
    private elementDataMap: Map<string, { id: string, color: string | null, selectionId: powerbi.visuals.ISelectionId }> = new Map();

    // Store all external IDs (or DbIds) from the dataset for bulk mapping
    private externalIds: string[] = [];

    // Viewer runtime
    private model: Autodesk.Viewing.Model = null;
    private idMapping: IdMapping = null;

    // Interactivity state for bidirectional filtering
    private allDbIds: number[] | null = null;
    private hasClearedFilters: boolean = false;
    private isDbIdSelectionActive: boolean = false;

    /**
     * Initializes the viewer visual.
     * @param options Additional visual initialization options.
     */
    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();
        this.container = options.element;

        // Create Status Bar for on-screen debugging
        this.statusDiv = document.createElement('div');
        this.statusDiv.style.position = 'absolute';
        this.statusDiv.style.bottom = '5px';
        this.statusDiv.style.left = '5px';
        this.statusDiv.style.color = 'white';
        this.statusDiv.style.backgroundColor = 'rgba(0, 0, 0, 0.7)';
        this.statusDiv.style.padding = '5px';
        this.statusDiv.style.fontSize = '12px';
        this.statusDiv.style.zIndex = '999';
        this.statusDiv.style.pointerEvents = 'none';
        this.statusDiv.style.display = 'none'; // Hidden by default, shown when data present
        this.container.appendChild(this.statusDiv);

        this.getAccessToken = this.getAccessToken.bind(this);
        this.onPropertiesLoaded = this.onPropertiesLoaded.bind(this);
        this.onSelectionChanged = this.onSelectionChanged.bind(this);
    }



    /**
     * Notifies the viewer visual of an update (data, viewmode, size change).
     * @param options Additional visual update options.
     */
    private allRows: powerbi.DataViewTableRow[] = [];

    /**
     * Notifies the viewer visual of an update (data, viewmode, size change).
     * @param options Additional visual update options.
     */
    public async update(options: VisualUpdateOptions): Promise<void> {
        // this.logVisualUpdateOptions(options);

        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualSettingsModel, options.dataViews[0]);
        const { accessTokenEndpoint } = this.formattingSettings.viewerCard;
        if (accessTokenEndpoint.value !== this.accessTokenEndpoint) {
            this.accessTokenEndpoint = accessTokenEndpoint.value;
            if (!this.viewer) {
                this.initializeViewer();
            }
        }

        this.currentDataView = options.dataViews[0];

        // Handle Pagination (Fetch More Data)
        if (options.operationKind === powerbi.VisualDataChangeOperationKind.Create) {
            this.allRows = [];
            this.elementDataMap.clear();
            this.externalIds = [];
        }

        // 1. Get the Table DataView
        const dataView = options.dataViews[0];
        if (!dataView || !dataView.table || !dataView.table.rows) {
            this.currentUrn = '';
            this.currentGuid = null;
            return;
        }

        // 2. Identify Column Indices
        const columns = dataView.table.columns;
        const urnIndex = columns.findIndex(c => c.roles["urn"]);
        const dbidsIndex = columns.findIndex(c => c.roles["dbids"]);
        const guidIndex = columns.findIndex(c => c.roles["guid"]);
        const colorIndex = columns.findIndex(c => c.roles["color"]);

        if (urnIndex === -1) {
            console.warn('Visual: URN column not found');
            return;
        }

        // 3. Process Rows
        // We accumulate rows to handle pagination
        this.allRows = this.allRows.concat(dataView.table.rows);

        // Process ONLY the new rows for SelectionId generation
        dataView.table.rows.forEach((row, rowIndex) => {
            // Create SelectionId based on the dbids column (primary key)
            const selectionIdBuilder = this.host.createSelectionIdBuilder();

            if (dbidsIndex !== -1) {
                selectionIdBuilder.withTable(dataView.table, rowIndex);
            }

            const selectionId = selectionIdBuilder.createSelectionId();
            const dbidValue = dbidsIndex !== -1 ? row[dbidsIndex] : null;
            const colorValue = colorIndex !== -1 ? String(row[colorIndex]) : null;

            // Store in our map
            if (dbidValue != null) {
                const key = selectionId.getKey();
                this.elementDataMap.set(key, {
                    id: String(dbidValue),
                    color: colorValue,
                    selectionId: selectionId
                });

                // Also keep track of external IDs for mapping later
                this.externalIds.push(String(dbidValue));
            }
        });

        // 4. Handle Model Loading (URN)
        let modelUrn: string | null = null;
        if (urnIndex !== -1 && this.allRows.length > 0) {
            modelUrn = String(this.allRows[0][urnIndex]);
        }

        // 5. Handle View GUID
        let viewGuid: string | null = null;
        if (guidIndex !== -1 && this.allRows.length > 0) {
            viewGuid = String(this.allRows[0][guidIndex]);
        }

        // 6. Handle Pagination
        let isFetching = false;
        if (this.currentDataView.metadata.segment) {
            const moreData = this.host.fetchMoreData();
            if (moreData) {
                console.log('Visual: Fetching more data...');
                isFetching = true;
            }
        }

        console.log(`Visual: Total rows loaded: ${this.allRows.length}`);

        // Update Status Bar
        this.statusDiv.style.display = 'block';
        if (isFetching) {
            this.statusDiv.innerText = `Rows: ${this.allRows.length} (Loading more...)`;
            this.statusDiv.style.color = 'yellow';
        }

        // Initialize Viewer if needed
        if (modelUrn && modelUrn !== this.currentUrn) {
            this.currentUrn = modelUrn;
            this.currentGuid = viewGuid;
            this.initializeViewer();
        } else if (viewGuid && viewGuid !== this.currentGuid && this.viewer) {
            // If URN is same but GUID changed, switch view
            this.currentGuid = viewGuid;
            console.log("Visual: View GUID changed to", viewGuid);
        }

        // 6.5. Handle incoming filters from other visuals (bidirectional interactivity)
        if (this.viewer && this.isViewerReady && this.idMapping && !isFetching) {
            await this.handleIncomingFilters(dataView);
        }

        // 7. Sync Selection & Colors
        if (this.viewer && this.isViewerReady && this.idMapping) {
            await this.syncSelectionState(isFetching);
            await this.syncColors();
        }
    }

    /**
     * Returns properties pane formatting model content hierarchies, properties and latest formatting values, Then populate properties pane.
     * This method is called once every time we open properties pane or when the user edit any format property. 
     */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    /**
     * Displays a notification that will automatically disappear after some time.
     * @param content HTML content to display inside the notification.
     */
    private showNotification(content: string): void {
        let notifications = this.container.querySelector('#notifications');
        if (!notifications) {
            notifications = document.createElement('div');
            notifications.id = 'notifications';
            this.container.appendChild(notifications);
        }
        const notification = document.createElement('div');
        notification.className = 'notification';
        notification.innerText = content;
        notifications.appendChild(notification);
        setTimeout(() => notifications.removeChild(notification), 5000);
    }

    /**
     * Initializes the viewer runtime.
     */
    private async initializeViewer(): Promise<void> {
        try {
            // Shim localStorage and sessionStorage to prevent SecurityError in Power BI sandbox
            try {
                const _test = window.localStorage;
            } catch (e) {
                console.warn('Storage access denied. Shimming...');
                const storageShim = {
                    getItem: () => null,
                    setItem: () => { },
                    removeItem: () => { },
                    clear: () => { }
                };
                Object.defineProperty(window, 'localStorage', { value: storageShim, writable: true });
                Object.defineProperty(window, 'sessionStorage', { value: storageShim, writable: true });
            }

            // Smart Environment Detection via Backend
            let env = this.formattingSettings?.viewerCard?.viewerEnv?.value;
            let region = this.formattingSettings?.viewerCard?.viewerRegion?.value || 'US';

            // If no env is explicitly set (or default), try to auto-detect via backend
            if ((!env || env === 'AutodeskProduction2') && this.currentUrn) {
                try {
                    const tokenData = await this.fetchToken(this.currentUrn);
                    if (tokenData && tokenData.detected_env) {
                        env = tokenData.detected_env;
                        console.log(`Smart Init: Backend detected env: ${env}`);
                    }
                    if (tokenData && tokenData.detected_region) {
                        region = tokenData.detected_region;
                        console.log(`Smart Init: Backend detected region: ${region}`);
                    }
                } catch (e) {
                    console.warn('Smart Init failed, falling back to default env', e);
                }
            }

            env = env || 'AutodeskProduction2';

            // API depends on Env: SVF2 -> streamingV2, SVF -> derivativeV2
            const api = env.includes('Production2') ? 'streamingV2' : 'derivativeV2';

            await initializeViewerRuntime({
                env: env,
                api: api,
                region: region,
                getAccessToken: (callback) => this.getAccessToken(callback), // Wrap to maintain context
                // @ts-ignore
                disabledExtensions: { 'Autodesk.Viewing.MixpanelExtension': true }
            });
            this.container.innerText = '';
            this.viewer = new Autodesk.Viewing.GuiViewer3D(this.container);
            this.viewer.start();
            this.viewer.loadExtension('Autodesk.VisualClusters');
            this.viewer.addEventListener(Autodesk.Viewing.OBJECT_TREE_CREATED_EVENT, this.onPropertiesLoaded);
            this.viewer.addEventListener(Autodesk.Viewing.SELECTION_CHANGED_EVENT, this.onSelectionChanged);
            if (this.currentUrn) {
                this.updateModel();
            }
        } catch (err) {
            this.showNotification('Could not initialize viewer runtime. Please see console for more details.');
            console.error(err);
        }
    }

    private async fetchToken(urn?: string): Promise<any> {
        try {
            let url = this.accessTokenEndpoint;
            if (urn) {
                // Append URN to query params
                const separator = url.includes('?') ? '&' : '?';
                url += `${separator}urn=${encodeURIComponent(urn)}`;
            }
            const response = await fetch(url);
            if (!response.ok) return null;
            return await response.json();
        } catch {
            return null;
        }
    }



    /**
     * Retrieves a new access token for the viewer.
     * @param callback Callback function to call with new access token.
     */
    private async getAccessToken(callback: (accessToken: string, expiresIn: number) => void): Promise<void> {
        try {
            const response = await fetch(this.accessTokenEndpoint);
            if (!response.ok) {
                throw new Error(await response.text());
            }
            const share = await response.json();
            if (!share.access_token) {
                throw new Error('Token response is missing access_token');
            }
            callback(share.access_token, share.expires_in);
        } catch (err) {
            this.showNotification(`Token Error: ${err.message}`);
            console.error(err);
        }
    }

    /**
     * Ensures that the correct model is loaded into the viewer.
     */
    private async updateModel(): Promise<void> {
        if (!this.viewer) {
            return;
        }

        if (this.model && this.model.getData().urn !== this.currentUrn) {
            this.viewer.unloadModel(this.model);
            this.model = null;
            this.idMapping = null;
        }

        try {
            if (this.currentUrn) {
                console.log('Visual: Raw URN:', this.currentUrn);

                // Sanitize URN: remove whitespace
                let safeUrn = this.currentUrn.trim();

                // If it starts with 'urn:', strip it for processing
                if (safeUrn.toLowerCase().startsWith('urn:')) {
                    safeUrn = safeUrn.substring(4);
                }

                // Check if it needs encoding (contains non-base64 characters like ':' or '.')
                // OR if it contains standard base64 characters that need replacing (+ or /)
                if (/[^A-Za-z0-9\-_]/.test(safeUrn)) {
                    // It contains characters NOT allowed in URL-safe Base64.
                    // Check if it's plain text (has :)
                    if (safeUrn.includes(':')) {
                        console.log('Visual: Detected plain text URN, encoding...');
                        let urnToEncode = this.currentUrn.trim();
                        if (!urnToEncode.toLowerCase().startsWith('urn:')) {
                            urnToEncode = 'urn:' + urnToEncode;
                        }
                        try {
                            safeUrn = btoa(urnToEncode).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
                        } catch (e) {
                            console.error('Failed to encode URN:', e);
                        }
                    } else {
                        // It might be standard Base64 (with + or / or =), fix it to URL-safe
                        console.log('Visual: Detected standard Base64 or invalid chars, fixing...');
                        safeUrn = safeUrn.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
                    }
                }

                console.log('Visual: Final Safe URN:', safeUrn);
                this.model = await loadModel(this.viewer, safeUrn, this.currentGuid);
                this.isViewerReady = true;
            }
        } catch (err) {
            let decodedUrn = '';
            try {
                decodedUrn = atob(this.currentUrn.replace(/-/g, '+').replace(/_/g, '/'));
            } catch (e) {
                decodedUrn = 'Invalid Base64';
            }
            let msg = `Could not load model. URN: ${this.currentUrn.substring(0, 10)}... Decoded: ${decodedUrn.substring(0, 50)}... `;
            if (err && typeof err === 'object') {
                if ('code' in err) msg += ` Code: ${err.code}`;
                if ('message' in err) msg += ` Message: ${err.message}`;
            } else {
                msg += ` Error: ${String(err)}`;
            }
            this.showNotification(msg);
            console.error(err);
        }
    }

    private sampleModelId: string = '';

    private async onPropertiesLoaded() {
        this.idMapping = new IdMapping(this.model);
        this.isViewerReady = true;

        // DEBUG: Log sample valid IDs from the model to help user fix data mismatch
        try {
            // @ts-ignore
            this.model.getExternalIdMapping((mapping) => {
                const keys = Object.keys(mapping);
                if (keys.length > 0) {
                    this.sampleModelId = keys[0];
                }
                console.log(`Visual: Model contains ${keys.length} elements.`);
                console.log('Visual: Sample VALID Model External IDs (Copy these to your data):', JSON.stringify(keys.slice(0, 5)));
            });
        } catch (e) {
            console.error('Visual: Could not log model IDs', e);
        }

        await this.syncSelectionState(false);
    }


    private async syncSelectionState(isFetching: boolean) {
        if (!this.viewer || !this.model) return;

        // If we are fetching more data, show loading status
        if (isFetching) {
            this.statusDiv.innerText = `Loading... (${this.allRows.length} rows)`;
            this.statusDiv.style.color = 'yellow';
            return;
        }

        const rowCount = this.allRows.length;
        if (rowCount === 0) {
            this.statusDiv.innerText = 'No Data Rows';
            this.statusDiv.style.color = 'white';
            // Show full model when no data
            showAll(this.viewer, this.model);
            return;
        }

        // Check if we have any valid IDs in our map
        if (this.elementDataMap.size === 0) {
            this.statusDiv.innerText = 'No IDs mapped';
            this.statusDiv.style.color = 'orange';
            // Show full model when no IDs mapped
            showAll(this.viewer, this.model);
            return;
        }

        // NOTE: Isolation is now handled by handleIncomingFilters().
        // This method only updates the status bar.
        // Only show full model if we haven't applied any isolation yet.
        if (!this.allDbIds || this.allDbIds.length === 0) {
            // Initial load: show full model
            showAll(this.viewer, this.model);
        }
        
        this.statusDiv.innerText = `Rows: ${rowCount} | Model Ready`;
        this.statusDiv.style.color = 'lightgreen';

        console.log(`Visual: Model loaded with ${rowCount} rows.`);
    }

    /**
     * Finds the index of a category by its role name in the dataView.
     * @param dataView The data view to search in.
     * @param roleName The role name to find.
     * @returns The index of the category, or -1 if not found.
     */
    private findCategoryIndexByRole(dataView: DataView, roleName: string): number {
        if (!dataView?.categorical?.categories) return -1;
        
        const categories = dataView.categorical.categories;
        for (let i = 0; i < categories.length; i++) {
            const category = categories[i];
            if (category.source?.roles && category.source.roles[roleName]) {
                return i;
            }
        }
        return -1;
    }

    /**
     * Handles incoming filters from other visuals and isolates/focuses elements in the viewer.
     * @param dataView The current data view.
     */
    private async handleIncomingFilters(dataView: DataView): Promise<void> {
        if (!this.viewer || !this.model || !this.idMapping) return;

        // Skip if this filter change came from our own selection
        if (this.isDbIdSelectionActive) {
            this.isDbIdSelectionActive = false;
            return;
        }

        // Check if we have categorical data
        const cat = dataView?.categorical?.categories;
        if (!cat || cat.length === 0) {
            // Fallback to table data if categorical is not available
            if (dataView?.table?.rows) {
                const dbidsIndex = dataView.table.columns.findIndex(c => c.roles["dbids"]);
                if (dbidsIndex !== -1) {
                    const currentDbIds = dataView.table.rows
                        .map(row => {
                            const value = row[dbidsIndex];
                            const parsed = typeof value === 'string' ? parseInt(value) : Number(value);
                            return isNaN(parsed) ? null : parsed;
                        })
                        .filter((id): id is number => id !== null);

                    await this.applyIsolationToViewer(currentDbIds, dataView);
                }
            }
            return;
        }

        // Find the dbids category index
        const dbidsIndex = this.findCategoryIndexByRole(dataView, "dbids");
        if (dbidsIndex === -1) {
            return;
        }

        // Extract current visible dbIds from categorical data
        const category = cat[dbidsIndex];
        if (!category || !category.values) {
            return;
        }

        const currentDbIds = category.values
            .map(v => {
                const parsed = typeof v === 'string' ? parseInt(v) : Number(v);
                return isNaN(parsed) ? null : parsed;
            })
            .filter((id): id is number => id !== null);

        // Initialize allDbIds on first update (when no filters are applied)
        if (this.allDbIds === null) {
            this.allDbIds = currentDbIds.slice();
        }

        // Check if a filter is applied
        const filterApplied = dataView.metadata?.isDataFilterApplied === true;
        const hasSegment = !!dataView.metadata?.segment;

        // Apply isolation based on filter state
        await this.applyIsolationToViewer(currentDbIds, dataView, filterApplied, hasSegment);
    }

    /**
     * Applies isolation and fitToView to the viewer based on filtered dbIds.
     * @param currentDbIds The current visible dbIds after filtering.
     * @param dataView The data view.
     * @param filterApplied Whether a filter is currently applied.
     * @param segment Whether we're in a segment (pagination).
     */
    private async applyIsolationToViewer(
        currentDbIds: number[],
        dataView: DataView,
        filterApplied?: boolean,
        segment?: boolean
    ): Promise<void> {
        if (!this.viewer || !this.model) return;

        // If we're fetching more data (segment), don't apply isolation yet
        if (segment) {
            return;
        }

        // Check if filter is applied (default to true if currentDbIds is smaller than allDbIds)
        const isFiltered = filterApplied !== undefined 
            ? filterApplied 
            : (this.allDbIds && currentDbIds.length < this.allDbIds.length);

        if (isFiltered && currentDbIds.length > 0) {
            // Map external IDs to dbIds if needed
            let dbIdsToIsolate: number[] = [];
            
            // Check if currentDbIds are external IDs (strings) or dbIds (numbers)
            const firstId = currentDbIds[0];
            const isExternalId = typeof firstId === 'string' || 
                (this.externalIds.length > 0 && this.externalIds.includes(String(firstId)));

            if (isExternalId && this.idMapping) {
                try {
                    dbIdsToIsolate = await this.idMapping.getDbids(currentDbIds.map(String));
                    dbIdsToIsolate = dbIdsToIsolate.filter(id => id != null && !isNaN(id));
                } catch (e) {
                    console.error('Visual: Error mapping external IDs to dbIds for isolation', e);
                    // Fallback: assume they are already dbIds
                    dbIdsToIsolate = currentDbIds.filter(id => !isNaN(Number(id))).map(Number);
                }
            } else {
                dbIdsToIsolate = currentDbIds.filter(id => !isNaN(Number(id))).map(Number);
            }

            if (dbIdsToIsolate.length > 0) {
                console.log(`Visual: Isolating and selecting ${dbIdsToIsolate.length} elements due to filter`);
                
                // Set programmatic selection flag to prevent triggering onSelectionChanged
                this.isProgrammaticSelection = true;
                
                // Clear any existing selection first
                this.viewer.clearSelection();
                
                // Isolate and fit to view
                isolateDbIds(this.viewer, dbIdsToIsolate, this.model);
                fitToView(this.viewer, dbIdsToIsolate, this.model);
                
                // Wait a bit for isolation to complete, then select the elements
                // This ensures the elements are visible before selection
                setTimeout(() => {
                    try {
                        // Verify elements are isolated before selecting
                        const isolatedNodes = this.viewer.getIsolatedNodes();
                        console.log(`Visual: Isolated nodes count: ${isolatedNodes ? isolatedNodes.length : 0}`);
                        
                        // Select the elements visually in the viewer
                        this.viewer.select(dbIdsToIsolate);
                        
                        // Verify selection was applied
                        const selectedNodes = this.viewer.getSelection();
                        console.log(`Visual: Selected ${selectedNodes.length} elements in viewer (requested ${dbIdsToIsolate.length})`);
                        
                        if (selectedNodes.length !== dbIdsToIsolate.length) {
                            console.warn(`Visual: Selection mismatch - requested ${dbIdsToIsolate.length}, got ${selectedNodes.length}`);
                        }
                    } catch (e) {
                        console.error('Visual: Error selecting elements', e);
                    }
                    
                    // Reset flag after selection completes
                    setTimeout(() => {
                        this.isProgrammaticSelection = false;
                    }, 50);
                }, 200);
                
                this.hasClearedFilters = false;
            }
        } else if (!isFiltered && this.hasClearedFilters === false) {
            // Clear isolation and show all
            console.log('Visual: Clearing isolation, showing all elements');
            
            // Set programmatic selection flag
            this.isProgrammaticSelection = true;
            
            // Clear selection
            this.viewer.clearSelection();
            
            // Show all elements
            showAll(this.viewer, this.model);
            
            // Reset flag after a short delay
            setTimeout(() => {
                this.isProgrammaticSelection = false;
            }, 100);
            
            this.hasClearedFilters = true;
        }
    }

    /**
     * Handles dbId selection changes from the viewer and applies filters to other visuals.
     * @param selectedDbIds Array of selected dbIds from the viewer.
     */
    private async handleDbIds(selectedDbIds: number[]): Promise<void> {
        if (!this.host || !this.currentDataView) return;

        // If no selection, clear filters
        if (!selectedDbIds || selectedDbIds.length === 0) {
            this.host.applyJsonFilter(null, "general", "selfFilter", 0);
            this.host.applyJsonFilter(null, "general", "filter", 0);
            this.isDbIdSelectionActive = false;
            return;
        }

        // Find the dbids category index
        const dbidsIndex = this.findCategoryIndexByRole(this.currentDataView, "dbids");
        if (dbidsIndex === -1) {
            console.warn('Visual: dbids category not found in categorical data');
            return;
        }

        const category = this.currentDataView.categorical.categories[dbidsIndex];
        if (!category || !category.source) {
            console.warn('Visual: Invalid category structure');
            return;
        }

        // Extract table and column from the category source
        const queryName = category.source.queryName || '';
        const table = queryName.split('.')[0] || '';
        const column = category.source.displayName || category.source.queryName || '';

        if (!table || !column) {
            console.warn('Visual: Could not determine table or column for filter');
            return;
        }

        // Build filter target
        const target: models.IFilterColumnTarget = {
            table: table,
            column: column
        };

        // Map dbIds to external IDs if possible, otherwise use dbIds as strings
        let filterValues: string[] = [];
        
        if (this.idMapping) {
            try {
                const externalIds = await this.idMapping.getExternalIds(selectedDbIds);
                filterValues = externalIds.filter(id => id != null && id !== '');
                
                // Fallback to dbIds if external IDs are not available
                if (filterValues.length === 0) {
                    filterValues = selectedDbIds.map(String);
                }
            } catch (e) {
                console.error('Visual: Error mapping dbIds to external IDs', e);
                filterValues = selectedDbIds.map(String);
            }
        } else {
            filterValues = selectedDbIds.map(String);
        }

        if (filterValues.length > 0) {
            const filter = new models.BasicFilter(
                target,
                "In",
                filterValues
            );

            console.log('Visual: Applying JSON filter for selected dbIds:', filter);
            this.host.applyJsonFilter(filter, "general", "filter", 0);
            this.isDbIdSelectionActive = true;
        }
    }

    private async onSelectionChanged() {
        if (this.isProgrammaticSelection) return;

        const selectedDbids = this.viewer.getSelection();
        
        // Handle filter application to other visuals
        await this.handleDbIds(selectedDbids);

        if (selectedDbids.length === 0) {
            this.selectionManager.clear();
            this.statusDiv.innerText = `Rows: ${this.allRows.length} | Ready`;
            this.statusDiv.style.color = 'white';
            return;
        }

        let keysToLookup: string[] = [];
        keysToLookup.push(...selectedDbids.map(id => id.toString()));

        try {
            const externalIds = await this.idMapping.getExternalIds(selectedDbids);
            keysToLookup.push(...externalIds);

            const firstExt = externalIds.length > 0 ? externalIds[0] : 'None';
            const firstInt = selectedDbids[0];
            this.statusDiv.innerText = `Selected: ${firstInt} (Int) / ${firstExt} (Ext)`;
            this.statusDiv.style.color = 'cyan';
        } catch (e) {
            console.error('Visual: Error getting external IDs', e);
        }

        const selectionIds: powerbi.extensibility.ISelectionId[] = [];

        // Iterate map to find matching keys
        this.elementDataMap.forEach((data, key) => {
            if (keysToLookup.includes(data.id)) {
                selectionIds.push(data.selectionId);
            }
        });

        this.selectionManager.select(selectionIds);
    }

    /**
     * Applies colors from the Power BI data to the Viewer elements.
     */
    private async syncColors() {
        if (!this.viewer) return;

        this.viewer.clearThemingColors(this.viewer.model);
        const colorGroups: Map<string, number[]> = new Map();

        // Prepare a list of IDs that need mapping
        const idsToMap: string[] = [];

        for (const [key, data] of this.elementDataMap) {
            if (!data.color) continue;

            // If it's already a number, great. If not, we might need to map it.
            const parsedId = parseInt(data.id);
            if (isNaN(parsedId)) {
                idsToMap.push(data.id);
            }
        }

        // Batch map external IDs if needed
        let mappedIdsMap: Map<string, number> = new Map();
        if (idsToMap.length > 0 && this.idMapping) {
            try {
                const dbIds = await this.idMapping.getDbids(idsToMap);
                for (let i = 0; i < idsToMap.length; i++) {
                    if (dbIds[i]) {
                        mappedIdsMap.set(idsToMap[i], dbIds[i]);
                    }
                }
            } catch (e) {
                console.error("Visual: Error mapping IDs for colors", e);
            }
        }

        // Now assign colors
        for (const [key, data] of this.elementDataMap) {
            if (!data.color) continue;

            let dbId: number | null = null;
            const parsedId = parseInt(data.id);

            if (!isNaN(parsedId)) {
                dbId = parsedId;
            } else {
                // Try to get from our batch mapping
                const mapped = mappedIdsMap.get(data.id);
                if (mapped) dbId = mapped;
            }

            if (dbId !== null) {
                if (!colorGroups.has(data.color)) {
                    colorGroups.set(data.color, []);
                }
                colorGroups.get(data.color)!.push(dbId);
            }
        }

        colorGroups.forEach((dbIds, colorHex) => {
            const vector = this.hexToVector4(colorHex);
            if (vector) {
                dbIds.forEach(dbId => {
                    this.viewer.setThemingColor(dbId, vector, this.viewer.model);
                });
            }
        });
    }

    private hexToVector4(hex: string): THREE.Vector4 | null {
        if (!hex) return null;
        hex = hex.replace('#', '');
        if (hex.length === 6) {
            const r = parseInt(hex.substring(0, 2), 16) / 255;
            const g = parseInt(hex.substring(2, 4), 16) / 255;
            const b = parseInt(hex.substring(4, 6), 16) / 255;
            return new THREE.Vector4(r, g, b, 1);
        }
        return null;
    }

    private collectDesignUrns(dataView: DataView, urnIndex: number): string[] {
        let urns = new Set(dataView.table.rows.map(row => row[urnIndex].valueOf() as string));
        return Array.from(urns);
    }

    private logVisualUpdateOptions(options: VisualUpdateOptions) {
        const EditMode = {
            [powerbi.EditMode.Advanced]: 'Advanced',
            [powerbi.EditMode.Default]: 'Default',
        };
        const VisualDataChangeOperationKind = {
            [powerbi.VisualDataChangeOperationKind.Append]: 'Append',
            [powerbi.VisualDataChangeOperationKind.Create]: 'Create',
            [powerbi.VisualDataChangeOperationKind.Segment]: 'Segment',
        };
        const VisualUpdateType = {
            [powerbi.VisualUpdateType.All]: 'All',
            [powerbi.VisualUpdateType.Data]: 'Data',
            [powerbi.VisualUpdateType.Resize]: 'Resize',
            [powerbi.VisualUpdateType.ResizeEnd]: 'ResizeEnd',
            [powerbi.VisualUpdateType.Style]: 'Style',
            [powerbi.VisualUpdateType.ViewMode]: 'ViewMode',
        };
        const ViewMode = {
            [powerbi.ViewMode.Edit]: 'Edit',
            [powerbi.ViewMode.InFocusEdit]: 'InFocusEdit',
            [powerbi.ViewMode.View]: 'View',
        };
        console.debug('editMode', EditMode[options.editMode]);
        console.debug('isInFocus', options.isInFocus);
        console.debug('jsonFilters', options.jsonFilters);
        console.debug('operationKind', VisualDataChangeOperationKind[options.operationKind]);
        console.debug('type', VisualUpdateType[options.type]);
        console.debug('viewMode', ViewMode[options.viewMode]);
        console.debug('viewport', options.viewport);
        console.debug('Data views:');
        for (const dataView of options.dataViews) {
            console.debug(dataView);
        }
    }
}

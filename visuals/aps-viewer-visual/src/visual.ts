/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
'use strict';

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import { VisualSettingsModel } from "./settings";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { initializeViewerRuntime, loadModel, IdMapping, launchViewer, isolateDbIds, fitToView, showAll } from "./viewer.utils";
import * as models from "powerbi-models";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import DataView = powerbi.DataView;
import DataViewTableRow = powerbi.DataViewTableRow;

<<<<<<< HEAD
=======
import { VisualSettingsModel } from './settings';
import { initializeViewerRuntime, loadModel, IdMapping, isolateDbIds, fitToView, showAll } from './viewer.utils';
import * as models from 'powerbi-models';

/**
 * Custom visual wrapper for the Autodesk Platform Services Viewer.
 */
>>>>>>> 637fefe79f416cc605a6e8b3d4f2c4a2a103b2da
export class Visual implements IVisual {
    // Visual state
    private host: IVisualHost;
    private statusDiv: HTMLDivElement;
    private container: HTMLElement;
    private formattingSettings: VisualSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private currentDataView: DataView = null;
    private selectionManager: ISelectionManager = null;

    // Flags for interaction pattern
    private isDbIdSelectionActive: boolean = false;
    private lastFilteredIds: string[] | null = null; // Track last filtered IDs to detect state changes
    private allDbIds: string[] | null = null; // Using string because we store IDs as strings in elementDataMap
    private pendingSelection: string[] | null = null; // Store selection to apply when viewer is ready
    private isProgrammaticSelection: boolean = false; // Prevent handleDbIds from firing when selection is programmatic

    // Visual inputs
    private accessTokenEndpoint: string = '';

    private currentUrn: string = "";
    private currentGuid: string | null = null;
    private viewer: Autodesk.Viewing.Viewer3D = null;
    private model: Autodesk.Viewing.Model = null;
    private idMapping: IdMapping = null;
    private isViewerReady: boolean = false;

    // Data Storage
    private allRows: DataViewTableRow[] = [];
    private elementDataMap: Map<string, { id: string, color: string | null, selectionId: ISelectionId }> = new Map();
    private externalIds: string[] = []; // For mapping
    private dbIdToColumnValueMap: Map<number, string> = new Map(); // Map dbId -> column value (for filtering)

<<<<<<< HEAD
=======
    // Interactivity state for bidirectional filtering
    private allDbIds: number[] | null = null;
    private hasClearedFilters: boolean = false;
    private isDbIdSelectionActive: boolean = false;

    /**
     * Initializes the viewer visual.
     * @param options Additional visual initialization options.
     */
>>>>>>> 637fefe79f416cc605a6e8b3d4f2c4a2a103b2da
    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.selectionManager = this.host.createSelectionManager();

        // Register callback for when other visuals change selection
        // This is the PRIMARY way to detect cross-filtering from other visuals
        this.selectionManager.registerOnSelectCallback(() => {
            console.log('Visual: SelectionManager callback fired - another visual changed selection');
            // The update() method will be called automatically by Power BI after this
            // We'll handle the actual filtering logic there
        });

        // Create container
        this.container = document.createElement('div');
        this.container.id = 'forge-viewer';
        this.container.style.width = '100%';
        this.container.style.height = 'calc(100% - 20px)';
        this.container.style.position = 'relative';
        options.element.appendChild(this.container);

        // Status bar
        this.statusDiv = document.createElement('div');
        this.statusDiv.style.height = '20px';
        this.statusDiv.style.backgroundColor = '#333';
        this.statusDiv.style.color = 'white';
        this.statusDiv.style.fontSize = '12px';
        this.statusDiv.style.padding = '2px 5px';
        this.statusDiv.innerText = 'Initializing...';
        options.element.appendChild(this.statusDiv);
    }

    /**
     * Notifies the visual of an update (data, viewmode, size change).
     */
    public async update(options: VisualUpdateOptions): Promise<void> {
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualSettingsModel, options.dataViews[0]);
        const { accessTokenEndpoint } = this.formattingSettings.viewerCard;
        if (accessTokenEndpoint.value !== this.accessTokenEndpoint) {
            this.accessTokenEndpoint = accessTokenEndpoint.value;
            // If endpoint changes, we might need to re-init, but for now just update prop
        }

        this.currentDataView = options.dataViews[0];

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

        // 3. Handle Pagination and Data Accumulation
        const isCreate = options.operationKind === powerbi.VisualDataChangeOperationKind.Create;
        const isAppend = options.operationKind === powerbi.VisualDataChangeOperationKind.Append;
        const hasSegment = this.currentDataView?.metadata?.segment != null;

        // Check if data is filtered BEFORE processing
        const isDataFilterApplied = this.currentDataView.metadata && this.currentDataView.metadata.isDataFilterApplied === true;

        // CRITICAL: Only reset on Create if there's NO filter applied
        // When Power BI applies a filter, it sends Create operation, but we need to preserve the full dataset
        if (isCreate && !isDataFilterApplied) {
            // Fresh start - no filter, so reset everything
            console.log('Visual: Create operation without filter - resetting all data');
            this.allRows = [];
            this.elementDataMap.clear();
            this.externalIds = [];
            this.allDbIds = null;
            this.dbIdToColumnValueMap.clear();
        } else if (isCreate && isDataFilterApplied) {
            // Create with filter - preserve existing full dataset
            console.log('Visual: Create operation WITH filter - preserving existing full dataset');
            // Don't reset allRows, externalIds, or allDbIds
        }

        // Process current visible rows FIRST to get currentBatchIds
        // These are the filtered rows if filter is applied, or all rows if not
        const currentBatchIds: string[] = [];
        const currentRowSet = new Set<string>(); // Track unique IDs in current batch

        dataView.table.rows.forEach((row, rowIndex) => {
            const selectionIdBuilder = this.host.createSelectionIdBuilder();
            if (dbidsIndex !== -1) {
                selectionIdBuilder.withTable(dataView.table, rowIndex);
            }
            const selectionId = selectionIdBuilder.createSelectionId();
            const dbidValue = dbidsIndex !== -1 ? row[dbidsIndex] : null;
            const colorValue = colorIndex !== -1 ? String(row[colorIndex]) : null;

            if (dbidValue != null) {
                const key = selectionId.getKey();
                const idString = String(dbidValue);

                // Update or add to elementDataMap (always, for selection mapping)
                this.elementDataMap.set(key, {
                    id: idString,
                    color: colorValue,
                    selectionId: selectionId
                });

                // Add to current batch IDs (for filtering) - only if not already added
                if (!currentRowSet.has(idString)) {
                    currentBatchIds.push(idString);
                    currentRowSet.add(idString);
                }
            }
        });

        // Accumulate rows and IDs - ONLY when NOT filtered
        // This builds the full dataset during initial load and pagination
        if (!isDataFilterApplied) {
            // No filter: accumulate all rows to build full dataset (handles pagination)
            this.allRows = this.allRows.concat(dataView.table.rows);

            // Accumulate unique IDs to externalIds
            currentBatchIds.forEach(idString => {
                if (!this.externalIds.includes(idString)) {
                    this.externalIds.push(idString);
                }
            });

            // Update allDbIds with accumulated externalIds (represents full dataset)
            this.allDbIds = this.externalIds.slice(); // Create a copy of full dataset

            console.log(`Visual: Accumulated data - allRows: ${this.allRows.length}, externalIds: ${this.externalIds.length}`);
        }
        // When filtered: 
        // - allRows and externalIds remain as the last known full dataset (preserved from above)
        // - currentBatchIds contains the filtered IDs (use these for highlighting)

        // 4. Handle Model Loading (URN)
        // Use current dataView rows for URN/GUID (works even when filtered)
        let modelUrn: string | null = null;
        if (urnIndex !== -1 && dataView.table.rows.length > 0) {
            modelUrn = String(dataView.table.rows[0][urnIndex]);
        }

        // 5. Handle View GUID
        let viewGuid: string | null = null;
        if (guidIndex !== -1 && dataView.table.rows.length > 0) {
            viewGuid = String(dataView.table.rows[0][guidIndex]);
        }

        // 6. Handle Pagination Fetch
        let isFetching = false;
        if (this.currentDataView.metadata.segment) {
            const moreData = this.host.fetchMoreData();
            if (moreData) {
                isFetching = true;
                this.statusDiv.innerText = `Loading... (${this.allRows.length} rows, ${this.externalIds.length} unique IDs)`;
                this.statusDiv.style.color = 'yellow';
            } else {
                // Pagination complete
                this.statusDiv.innerText = `Rows: ${this.allRows.length} | IDs: ${this.externalIds.length} | Ready`;
                this.statusDiv.style.color = 'lightgreen';
            }
        } else {
            // Show both total rows and unique IDs
            const uniqueIdCount = this.externalIds.length;
            this.statusDiv.innerText = `Rows: ${this.allRows.length} | IDs: ${uniqueIdCount} | Ready`;
            this.statusDiv.style.color = 'white';
        }

        // 7. Initialize Viewer if needed
        if (modelUrn && modelUrn !== this.currentUrn) {
            this.currentUrn = modelUrn;
            this.currentGuid = viewGuid;
            await this.initializeViewer();
        } else if (viewGuid && viewGuid !== this.currentGuid && this.viewer) {
            this.currentGuid = viewGuid;
            console.log("Visual: View GUID changed to", viewGuid);
            // TODO: Implement view switching logic if needed
        }

<<<<<<< HEAD
        // 8. INCOMING FILTER LOGIC (Power BI -> Viewer)
        // Simplified logic with three clear states:
        // 1. External filter active: isDataFilterApplied === true AND !isDbIdSelectionActive
        // 2. Filter cleared: isDataFilterApplied === false AND lastFilteredIds !== null
        // 3. No filter (initial): isDataFilterApplied === false AND lastFilteredIds === null

        // Determine if we're still paginating (Create is NOT pagination unless it has segment, Append is pagination)
        // In Power BI, 'Create' operation usually means "new data view", which can be initial load or a filter change.
        // 'Append' means more data for the same view (pagination).
        const isPaginating = isAppend || hasSegment;

        // Detect filter by row count: if we have less rows than total known externalIds, it might be a filter
        // This helps when isDataFilterApplied is not reliable or when cross-filtering reduces row count
        const hasFilterByCount = 
            currentBatchIds.length > 0 && 
            this.externalIds.length > 0 && 
            currentBatchIds.length !== this.externalIds.length;

        // Determine if there's an external filter (from another visual)
        // External filter = we didn't initiate it AND (data is filtered metadata OR row count is reduced)
        // We exclude pagination to avoid flickering during data load
        const isExternalFilter = 
            !this.isDbIdSelectionActive && 
            !isPaginating && 
            (isDataFilterApplied || hasFilterByCount);

        // Get the IDs to isolate
        const idsToIsolate = isExternalFilter ? currentBatchIds : [];

        // Detect state transitions
        const wasFiltered = this.lastFilteredIds !== null;
        
        // Filter cleared = was filtered, and now no filter is active (metadata false AND count matches full dataset)
        const isFilterCleared = 
            wasFiltered && 
            !this.isDbIdSelectionActive && 
            !isDataFilterApplied && 
            !hasFilterByCount;

        // Enhanced logging
        console.log(`Visual: ========== FILTER DETECTION DEBUG ==========`);
        console.log(`Visual: operationKind: ${options.operationKind}, type: ${options.type}`);
        console.log(`Visual: isCreate: ${isCreate}, isAppend: ${isAppend}, hasSegment: ${hasSegment}`);
        console.log(`Visual: isDataFilterApplied (from metadata): ${isDataFilterApplied}`);
        console.log(`Visual: isDbIdSelectionActive (internal flag): ${this.isDbIdSelectionActive}`);
        console.log(`Visual: isPaginating: ${isPaginating}`);
        console.log(`Visual: hasFilterByCount: ${hasFilterByCount} (current: ${currentBatchIds.length}, total: ${this.externalIds.length})`);
        console.log(`Visual: Filter detection - isExternalFilter: ${isExternalFilter}`);
        console.log(`Visual: State transition - wasFiltered: ${wasFiltered}, isFilterCleared: ${isFilterCleared}`);
        console.log(`Visual: Data counts - idsToIsolate: ${idsToIsolate.length}, currentBatchIds: ${currentBatchIds.length}, allRows: ${this.allRows.length}, externalIds: ${this.externalIds.length}`);
        console.log(`Visual: Viewer state - viewer exists: ${!!this.viewer}, isViewerReady: ${this.isViewerReady}, idMapping exists: ${!!this.idMapping}`);
        console.log(`Visual: ============================================`);

        if (isDataFilterApplied && currentBatchIds.length > 0) {
            console.log(`Visual: Filtered data - Total IDs: ${currentBatchIds.length}, Sample (first 10):`, currentBatchIds.slice(0, 10));
        }
        
        if (currentBatchIds.length > 0 && currentBatchIds.length !== this.externalIds.length) {
            console.log(`Visual: POTENTIAL FILTER: currentBatchIds (${currentBatchIds.length}) differs from externalIds (${this.externalIds.length})`);
        }

        // Wait for viewer to be ready before processing filters
        if (!this.viewer || !this.isViewerReady || !this.idMapping) {
            // Store pending selection to apply when viewer is ready
            if (isExternalFilter && idsToIsolate.length > 0) {
                console.log(`Visual: Viewer not ready, storing ${idsToIsolate.length} IDs for later application...`);
                this.pendingSelection = idsToIsolate;
                this.lastFilteredIds = idsToIsolate;
            }
            return; // Exit early if viewer not ready
        }

        // STATE 1: External filter active - Isolate and highlight filtered elements
        // We check !isPaginating to avoid isolating partial data during load
        if (isExternalFilter && idsToIsolate.length > 0 && !isPaginating) {
            console.log(`Visual: External filter detected! Processing ${idsToIsolate.length} IDs to isolate.`);

            // Set flag to prevent selection loop
            this.isProgrammaticSelection = true;

            // Reset isDbIdSelectionActive when filter comes from external source
            this.isDbIdSelectionActive = false;

            // Apply filter (isolate, highlight with neon green, and fit to view)
            await this.syncSelectionState(true, idsToIsolate);

            // Track filtered IDs for state transition detection
            this.lastFilteredIds = idsToIsolate.slice();

            // Reset flag after operation
            this.isProgrammaticSelection = false;

            console.log(`Visual: Filter applied successfully, ${idsToIsolate.length} elements isolated and highlighted`);
        }
        // STATE 2: Filter cleared - Show all elements and restore original colors
        else if (isFilterCleared) {
            console.log(`Visual: Filter cleared detected! Showing all elements and restoring original state...`);

            // Set flag to prevent selection loop
            this.isProgrammaticSelection = true;

            // Clear isolation and selection
            this.viewer.clearSelection();
            this.viewer.clearThemingColors(this.model);
            showAll(this.viewer, this.model);

            // Reset state tracking
            this.lastFilteredIds = null;
            this.isDbIdSelectionActive = false;

            // Restore original colors from data
=======
        // 6.5. Handle incoming filters from other visuals (bidirectional interactivity)
        if (this.viewer && this.isViewerReady && this.idMapping && !isFetching) {
            await this.handleIncomingFilters(dataView);
        }

        // 7. Sync Selection & Colors
        if (this.viewer && this.isViewerReady && this.idMapping) {
            await this.syncSelectionState(isFetching);
>>>>>>> 637fefe79f416cc605a6e8b3d4f2c4a2a103b2da
            await this.syncColors();

            // Reset flag after operation
            this.isProgrammaticSelection = false;

            console.log(`Visual: Filter cleared successfully, showing all ${this.externalIds.length} elements`);
        }
        // STATE 3: No filter (initial state or no change) - Sync colors if needed
        else if (!isDataFilterApplied && !this.isDbIdSelectionActive && !isPaginating) {
            // Only sync colors if we're not in the middle of pagination
            // and there's no active selection from viewer or external filter
            if (this.lastFilteredIds === null) {
                // Initial state - sync colors once
                console.log(`Visual: Initial state - syncing colors for ${this.externalIds.length} elements`);
                await this.syncColors();
            }
        }
    }

    private async initializeViewer(): Promise<void> {
        try {
            const token = await this.fetchToken(this.currentUrn);
            if (!token) return;

            // Get performance profile from settings
            const profile = this.formattingSettings.viewerCard.performanceProfile.value.value as 'HighPerformance' | 'Balanced';

            // Get environment from settings
            const env = this.formattingSettings.viewerCard.viewerEnv.value;
            // Determine API based on Env (SVF2 -> streamingV2, SVF -> derivativeV2)
            const api = env === 'AutodeskProduction2' ? 'streamingV2' : 'derivativeV2';

            // Use the new launchViewer from utils
            const result = await launchViewer(
                this.container,
                this.currentUrn,
                token.access_token,
                this.currentGuid,
                (ids) => this.handleDbIds(ids), // Callback for selection
                profile, // Pass performance profile
                env, // Pass environment
                api // Pass API
            );

            this.viewer = result.viewer;
            this.model = result.model;
            this.idMapping = new IdMapping(this.model);
            this.isViewerReady = true;

            console.log("Visual: Viewer initialized successfully with profile:", profile);

            // CRITICAL: Build reverse mapping (dbId -> column value) for all loaded data
            // This allows us to map selected dbIds back to the values in Power BI column
            await this.buildReverseMapping();

            // Initial sync
            await this.syncColors();

            // Apply pending selection if any
            if (this.pendingSelection && this.pendingSelection.length > 0) {
                console.log(`Visual: Applying pending selection after viewer initialization: ${this.pendingSelection.length} IDs`);
                await this.syncSelectionState(true, this.pendingSelection);
                this.pendingSelection = null;
            }

        } catch (error) {
            console.error("Visual: Error initializing viewer", error);
            this.statusDiv.innerText = "Error loading viewer";
            this.statusDiv.style.color = "red";
        }
    }

    /**
     * Handles selection changes from the Viewer (Viewer -> Power BI)
     */
    private async handleDbIds(selectedDbIds: number[]) {
        if (!this.host) return;

        // Skip if this selection change came from programmatic selection (external filter)
        if (this.isProgrammaticSelection) {
            console.log("Visual: Ignoring programmatic selection change:", selectedDbIds.length, "IDs");
            return;
        }

        console.log("Visual: Selection changed in viewer (user interaction):", selectedDbIds);

        // 1. If no selection, clear filters
        if (selectedDbIds.length === 0) {
            this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.merge);
            this.isDbIdSelectionActive = false;
            return;
        }

        // 2. If selection, apply JSON filter
        // We need to find the target column for 'dbids' role
        if (!this.currentDataView || !this.currentDataView.table) return;

        const columns = this.currentDataView.table.columns;
        const dbidsIndex = columns.findIndex(c => c.roles["dbids"]);

        if (dbidsIndex === -1) return;

        const columnSource = columns[dbidsIndex];
        // Note: queryName is usually "Table.Column"
        
        // Try to get table and column names
        let target: models.IFilterColumnTarget;

        if (columnSource.queryName) {
            // Prefer queryName parsing for robustness
            const parts = columnSource.queryName.split('.');
            if (parts.length >= 2) {
                target = {
                    table: parts[0],
                    column: parts[1] // Basic splitting
                };
            } else {
                target = {
                    table: columnSource.queryName, 
                    column: columnSource.displayName
                };
            }
        } else {
            // Fallback
            target = {
                table: "Data", 
                column: columnSource.displayName
            };
        }

        // CRITICAL: Map DbIds to the exact values in the Power BI column
        // Since the column contains dbId values (as strings), we map: dbId (number) -> dbId (string)
        // We rely on dbIdToColumnValueMap built during update, or check against externalIds directly.
        let filterValues: string[] = [];
        let missingIds: number[] = [];

        for (const dbId of selectedDbIds) {
            // 1. Try reverse map
            const columnValue = this.dbIdToColumnValueMap.get(dbId);
            if (columnValue) {
                filterValues.push(columnValue);
            } else {
                // 2. Fallback: check if the stringified dbId exists in our known dataset
                const dbIdStr = String(dbId);
                if (this.externalIds.includes(dbIdStr)) {
                    filterValues.push(dbIdStr);
                } else {
                    missingIds.push(dbId);
                }
            }
        }

        console.log(`Visual: Mapped ${selectedDbIds.length} DbIds to ${filterValues.length} column values for filtering`);
        
        if (missingIds.length > 0) {
            console.warn(`Visual: Warning - Selected DbIds [${missingIds.slice(0, 5)}...] do not exist in Power BI data column. Filter will ignore them.`);
        }

        if (filterValues.length > 0) {
            console.log(`Visual: Filter values (first 10):`, filterValues.slice(0, 10));
            
            const filter = new models.BasicFilter(
                target,
                "In",
                filterValues
            );

            console.log("Visual: Applying JSON filter:", filter);
            this.host.applyJsonFilter(filter, "general", "filter", powerbi.FilterAction.merge);
            this.isDbIdSelectionActive = true;
        } else {
            console.warn("Visual: No valid filter values found for selection. Skipping filter application to avoid empty results.");
            // We do NOT clear filter here to avoid resetting other visuals to 'All' when clicking empty space or invalid objects.
            // If user wants to clear, they click empty space -> selectedDbIds.length === 0 handled above.
        }
    }

    /**
     * Build reverse mapping from dbId (number) to column value (string)
     * Since the column contains dbId values directly, we map: dbId (number) -> dbId (string from column)
     * This is critical for filtering: when user selects an element in viewer (dbId),
     * we need to map it back to the exact string value in Power BI column
     */
    private async buildReverseMapping(): Promise<void> {
        if (!this.idMapping || !this.externalIds || this.externalIds.length === 0) {
            console.log("Visual: buildReverseMapping - No data available yet");
            return;
        }

        console.log(`Visual: buildReverseMapping - Building reverse mapping for ${this.externalIds.length} dbId values from column`);

        try {
            // Map all column values (which are dbIds as strings) to actual dbIds (numbers)
            // This validates that the column values are valid dbIds in the model
            const dbIds = await this.idMapping.smartMapToDbIds(this.externalIds);
            
            // Build reverse mapping: dbId (number) -> dbId (string from column)
            // This allows us to map selected dbIds back to the exact string in the column
            // Clear existing map to avoid stale data
            this.dbIdToColumnValueMap.clear();
            
            for (let i = 0; i < this.externalIds.length; i++) {
                const columnValue = this.externalIds[i]; // This is the dbId as string from Power BI column
                if (i < dbIds.length && dbIds[i] != null) {
                    // Map: dbId (number from viewer) -> dbId (string from column)
                    this.dbIdToColumnValueMap.set(dbIds[i], columnValue);
                }
            }

            console.log(`Visual: buildReverseMapping - Built reverse mapping for ${this.dbIdToColumnValueMap.size} dbIds`);
            if (this.dbIdToColumnValueMap.size > 0) {
                const sampleEntries = Array.from(this.dbIdToColumnValueMap.entries()).slice(0, 5);
                console.log(`Visual: buildReverseMapping - Sample mappings:`, sampleEntries);
            }
        } catch (e) {
            console.error("Visual: Error building reverse mapping", e);
        }
    }

    private async fetchToken(urn: string): Promise<any> {
        if (!this.accessTokenEndpoint) {
            return null;
        }
        try {
            const response = await fetch(this.accessTokenEndpoint);
            if (!response.ok) throw new Error(response.statusText);
            return await response.json();
        } catch (e) {
            console.error("Token fetch error", e);
            return null;
        }
    }

    private async syncSelectionState(isolate: boolean, idsToIsolate: string[]) {
        // Verify viewer and model are ready
        if (!this.viewer || !this.model || !this.idMapping) {
            console.log(`Visual: syncSelectionState - Viewer/model not ready, storing pending selection for ${idsToIsolate.length} IDs`);
            if (isolate && idsToIsolate.length > 0) {
                this.pendingSelection = idsToIsolate;
            }
            return;
        }

        // Apply any pending selection if viewer is now ready
        if (this.pendingSelection && this.pendingSelection.length > 0) {
            console.log(`Visual: syncSelectionState - Applying pending selection for ${this.pendingSelection.length} IDs`);
            const pendingIds = this.pendingSelection;
            this.pendingSelection = null;
            await this.syncSelectionState(true, pendingIds);
            return;
        }

        // PATH 1: Isolate and highlight specific elements
        if (isolate && idsToIsolate.length > 0) {
            try {
                console.log(`Visual: syncSelectionState - Attempting to isolate and highlight ${idsToIsolate.length} IDs`);
                console.log(`Visual: syncSelectionState - Input IDs (first 10):`, idsToIsolate.slice(0, 10));
                console.log(`Visual: syncSelectionState - Input IDs (last 10):`, idsToIsolate.slice(-10));
                console.log(`Visual: syncSelectionState - All input IDs:`, idsToIsolate);

                // Use smart mapping that validates IDs exist in model
                // These IDs come from Power BI column and should be dbId values as strings
                const validDbIds = await this.idMapping.smartMapToDbIds(idsToIsolate);
                
                console.log(`Visual: syncSelectionState - Mapping result: ${validDbIds.length} valid DbIds from ${idsToIsolate.length} input IDs`);
                
                // CRITICAL: Build a proper mapping from column values to dbIds
                // Since smartMapToDbIds may not preserve order or may skip invalid IDs,
                // we need to map each column value to its corresponding dbId
                const columnValueToDbIdMap = new Map<string, number>();
                for (let i = 0; i < idsToIsolate.length; i++) {
                    const columnValue = idsToIsolate[i];
                    // Try to find the corresponding dbId
                    // Since we're using direct dbId mapping, the column value should be the dbId as string
                    const parsedDbId = parseInt(columnValue, 10);
                    if (!isNaN(parsedDbId) && validDbIds.includes(parsedDbId)) {
                        columnValueToDbIdMap.set(columnValue, parsedDbId);
                        // Also update reverse mapping
                        this.dbIdToColumnValueMap.set(parsedDbId, columnValue);
                    }
                }
                
                console.log(`Visual: syncSelectionState - Built column value to dbId mapping: ${columnValueToDbIdMap.size} entries`);

                console.log(`Visual: syncSelectionState - Mapped ${idsToIsolate.length} IDs to ${validDbIds.length} valid DbIds`);

                // DEBUG: Log actual DbIds to see what we're trying to isolate
                if (validDbIds.length > 0) {
                    console.log(`Visual: syncSelectionState - Mapped DbIds (first 10):`, validDbIds.slice(0, 10));
                    console.log(`Visual: syncSelectionState - Mapped DbIds (last 10):`, validDbIds.slice(-10));
                }

                if (validDbIds.length === 0) {
                    console.warn(`Visual: syncSelectionState - None of the ${idsToIsolate.length} IDs could be mapped to valid DbIds in the model.`);
                    console.warn(`Visual: syncSelectionState - Sample IDs attempted:`, idsToIsolate.slice(0, 20));
                    // Show all if no valid IDs
                    this.viewer.clearSelection();
                    this.viewer.clearThemingColors(this.model);
                    showAll(this.viewer, this.model);
                    return;
                }

                // Log mapping efficiency
                if (validDbIds.length < idsToIsolate.length) {
                    console.warn(`Visual: syncSelectionState - Only ${validDbIds.length} of ${idsToIsolate.length} IDs were valid.`);
                }

                // Verify the DbIds are actually visible/selectable in the model
                const tree = this.model.getInstanceTree();
                const verifiedDbIds: number[] = [];
                const containerNodes: number[] = []; // Nodes with children (not leaf geometry)

                for (const dbid of validDbIds) {
                    try {
                        const nodeType = tree.getNodeType(dbid);
                        if (nodeType != null) {
                            verifiedDbIds.push(dbid);

                            // Check if this is a container node (has children)
                            const childCount = tree.getChildCount(dbid);
                            if (childCount > 0) {
                                containerNodes.push(dbid);
                            }
                        }
                    } catch (e) {
                        // Node doesn't exist, skip it
                    }
                }

<<<<<<< HEAD
                console.log(`Visual: syncSelectionState - Final verified DbIds: ${verifiedDbIds.length}`);

                // CRITICAL DEBUG: Check if we're trying to isolate container nodes
                if (containerNodes.length > 0) {
                    console.warn(`Visual: syncSelectionState - WARNING: ${containerNodes.length} of ${verifiedDbIds.length} DbIds are container nodes (have children)`);
                    console.warn(`Visual: syncSelectionState - Container nodes (first 5):`, containerNodes.slice(0, 5));
                    console.warn(`Visual: syncSelectionState - Expanding container nodes to include all children for visibility`);
=======
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
>>>>>>> 637fefe79f416cc605a6e8b3d4f2c4a2a103b2da
                }

                // CRITICAL FIX: Expand container nodes to include all their children
                // This ensures that leaf nodes (with actual geometry) are included in isolation
                const expandedDbIds = new Set<number>();

                for (const dbid of verifiedDbIds) {
                    // Add the node itself
                    expandedDbIds.add(dbid);

                    // If it has children, recursively add all descendants
                    const childCount = tree.getChildCount(dbid);
                    if (childCount > 0) {
                        // Recursively get all children
                        tree.enumNodeChildren(dbid, (childId) => {
                            expandedDbIds.add(childId);
                        }, true); // true = recursive
                    }
                }

                const finalDbIds = Array.from(expandedDbIds);

<<<<<<< HEAD
                console.log(`Visual: syncSelectionState - Expanded ${verifiedDbIds.length} DbIds to ${finalDbIds.length} DbIds (including children)`);
                if (finalDbIds.length > verifiedDbIds.length) {
                    console.log(`Visual: syncSelectionState - Added ${finalDbIds.length - verifiedDbIds.length} child nodes for visibility`);
                }

                if (finalDbIds.length > 0) {
                    // Comportamiento EXACTO del "Navegador de modelo":
                    //  1. AISLAR los elementos (ocultar el resto del modelo) - CRÍTICO
                    //  2. Seleccionar los nodos
                    //  3. Aplicar color de resaltado (verde neón)
                    //  4. Enfocar la cámara

                    console.log(`Visual: syncSelectionState - Applying model-browser-like isolation and selection for ${finalDbIds.length} DbIds`);

                    // Step 1: Clear any existing selection and theming colors
                    console.log(`Visual: syncSelectionState - Clearing previous selection and theming`);
                    
                    // IMPORTANTE: Primero limpiar colores, luego selección
                    this.viewer.clearThemingColors(this.model);
                    this.viewer.clearSelection();

                    // Step 2: Programmatic selection (como si hicieras click en el Navegador de modelo)
                    // Usamos el flag isProgrammaticSelection para que SELECTION_CHANGED_EVENT
                    // no dispare filtros de salida en handleDbIds.
                    this.isProgrammaticSelection = true; // ACTIVAR FLAG ANTES DE AISLAR/SELECCIONAR
=======
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
>>>>>>> 637fefe79f416cc605a6e8b3d4f2c4a2a103b2da

                    // Step 3: ISOLATION (Model Explorer behavior)
                    // We use expanded IDs to ensure geometry is visible
                    console.log(`Visual: syncSelectionState - Isolating ${finalDbIds.length} elements (hiding all others)`);
                    isolateDbIds(this.viewer, this.model, finalDbIds);

<<<<<<< HEAD
                    // Step 4: Selecting elements
                    console.log(`Visual: syncSelectionState - Selecting ${finalDbIds.length} elements`);
                    this.viewer.select(finalDbIds, this.model);
=======
        const selectedDbids = this.viewer.getSelection();
        
        // Handle filter application to other visuals
        await this.handleDbIds(selectedDbids);

        if (selectedDbids.length === 0) {
            this.selectionManager.clear();
            this.statusDiv.innerText = `Rows: ${this.allRows.length} | Ready`;
            this.statusDiv.style.color = 'white';
            return;
        }
>>>>>>> 637fefe79f416cc605a6e8b3d4f2c4a2a103b2da

                    // Step 5: Highlight nodes with neon green color
                    const neonGreen = new THREE.Vector4(0.224, 1.0, 0.078, 1.0); // #39FF14 in RGB normalized
                    console.log(`Visual: syncSelectionState - Theming ${finalDbIds.length} nodes with neon green`);
                    
                    // Aplicar color a cada nodo (expandido)
                    finalDbIds.forEach(dbid => {
                        this.viewer.setThemingColor(dbid, neonGreen, this.model);
                    });

                    // Forzar un repintado del visor tras aislamiento, selección y theming
                    if ((this.viewer as any).impl?.invalidate) {
                        console.log(`Visual: syncSelectionState - Forcing viewer.invalidate() after isolation/selection/theming`);
                        (this.viewer as any).impl.invalidate(true, true, true);
                    }

                    // Step 6: Fit camera to view the isolated elements
                    console.log(`Visual: syncSelectionState - Fitting camera to ${finalDbIds.length} isolated elements`);
                    fitToView(this.viewer, this.model, finalDbIds);

                    // Quitar el flag programático DESPUÉS de todas las operaciones
                    // Usamos un pequeño timeout para asegurar que los eventos se hayan procesado
                    setTimeout(() => {
                        this.isProgrammaticSelection = false;
                    }, 100);

<<<<<<< HEAD
                    console.log(`Visual: syncSelectionState - Successfully isolated, selected and highlighted ${finalDbIds.length} elements (model-browser-like behavior)`);
                } else {
                    console.error(`Visual: syncSelectionState - No valid DbIds after verification, showing all`);
                    this.viewer.clearSelection();
                    this.viewer.clearThemingColors(this.model);
                    showAll(this.viewer, this.model);
                }
            } catch (e) {
                console.error("Visual: Error syncing selection", e);
                this.viewer.clearSelection();
                this.viewer.clearThemingColors(this.model);
                showAll(this.viewer, this.model);
=======
        // Iterate map to find matching keys
        this.elementDataMap.forEach((data, key) => {
            if (keysToLookup.includes(data.id)) {
                selectionIds.push(data.selectionId);
>>>>>>> 637fefe79f416cc605a6e8b3d4f2c4a2a103b2da
            }
        }
        // PATH 2: Show all elements (no isolation)
        else {
            console.log("Visual: syncSelectionState - Showing all elements (no isolation)");
            this.viewer.clearSelection();
            this.viewer.clearThemingColors(this.model);
            showAll(this.viewer, this.model);
        }
    }

    private async syncColors() {
        if (!this.viewer || !this.model || !this.elementDataMap || !this.idMapping) return;

        // Collect colors
        const colorMap: { [color: string]: string[] } = {}; // color -> ids[]

        this.elementDataMap.forEach(value => {
            if (value.color) {
                if (!colorMap[value.color]) colorMap[value.color] = [];
                colorMap[value.color].push(value.id);
            }
        });

        // Apply colors
        this.viewer.clearThemingColors(this.model);

        for (const colorHex in colorMap) {
            const ids = colorMap[colorHex];
            try {
                // Use smart mapping to get valid DbIds
                const validDbIds = await this.idMapping.smartMapToDbIds(ids);

                if (validDbIds.length > 0) {
                    const vec = this.hexToVector4(colorHex);
                    validDbIds.forEach(dbid => {
                        this.viewer.setThemingColor(dbid, vec, this.model);
                    });
                }
            } catch (e) {
                console.error("Visual: Error applying colors", e);
            }
        }
    }

    /**
     * Sync colors only for filtered elements (preserves filter highlighting)
     */
    private async syncColorsForFilteredElements(filteredIds: string[]) {
        if (!this.viewer || !this.model || !this.elementDataMap || !this.idMapping) return;

        // Create a set of filtered IDs for quick lookup
        const filteredSet = new Set(filteredIds);

        // Collect colors only for filtered elements
        const colorMap: { [color: string]: string[] } = {}; // color -> ids[]

        this.elementDataMap.forEach(value => {
            // Only process colors for filtered elements
            if (value.color && filteredSet.has(value.id)) {
                if (!colorMap[value.color]) colorMap[value.color] = [];
                colorMap[value.color].push(value.id);
            }
        });

        // Apply colors only to filtered elements (don't clear all theming, preserve filter highlight)
        for (const colorHex in colorMap) {
            const ids = colorMap[colorHex];
            try {
                // Use smart mapping to get valid DbIds
                const validDbIds = await this.idMapping.smartMapToDbIds(ids);

                if (validDbIds.length > 0) {
                    const vec = this.hexToVector4(colorHex);
                    validDbIds.forEach(dbid => {
                        // Apply color, but don't clear the neon green filter highlight
                        // The filter highlight should take precedence
                        this.viewer.setThemingColor(dbid, vec, this.model);
                    });
                }
            } catch (e) {
                console.error("Visual: Error applying colors to filtered elements", e);
            }
        }
    }

    private hexToVector4(hex: string): THREE.Vector4 {
        if (hex.startsWith('#')) hex = hex.substring(1);
        const r = parseInt(hex.substring(0, 2), 16) / 255;
        const g = parseInt(hex.substring(2, 4), 16) / 255;
        const b = parseInt(hex.substring(4, 6), 16) / 255;
        return new THREE.Vector4(r, g, b, 1);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}

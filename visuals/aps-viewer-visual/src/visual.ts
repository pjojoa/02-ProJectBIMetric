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
    private isViewerIsolated: boolean = false; // Track if viewer currently has elements isolated

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
    private externalIds: string[] = []; // ExternalIds from Power BI column (persistent identifiers)
    private dbIdToExternalIdMap: Map<number, string> = new Map(); // Map dbId -> ExternalId (for filtering Viewer -> Power BI)
    private externalIdColumnName: string = ''; // Name of the externalIds column for filtering

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
        const externalIdsIndex = columns.findIndex(c => c.roles["externalIds"]);
        const guidIndex = columns.findIndex(c => c.roles["guid"]);
        const colorIndex = columns.findIndex(c => c.roles["color"]);

        if (urnIndex === -1) {
            console.warn('Visual: URN column not found');
            return;
        }
        
        // ExternalIds is REQUIRED - no fallback to dbids
        if (externalIdsIndex === -1) {
            console.error('Visual: ExternalIds column is REQUIRED but not found. Please assign a column to the "External IDs" role.');
            this.statusDiv.innerText = 'Error: External IDs column required';
            this.statusDiv.style.color = 'red';
            return;
        }
        
        // Store column name for filtering
        const externalIdColumn = columns[externalIdsIndex];
        this.externalIdColumnName = externalIdColumn.queryName || externalIdColumn.displayName;
        
        console.log(`Visual: Using externalIds column for element identification: ${this.externalIdColumnName}`);

        // 3. Handle Pagination and Data Accumulation
        const isCreate = options.operationKind === powerbi.VisualDataChangeOperationKind.Create;
        const isAppend = options.operationKind === powerbi.VisualDataChangeOperationKind.Append;
        const hasSegment = this.currentDataView?.metadata?.segment != null;

        // Check if data is filtered BEFORE processing
        const isDataFilterApplied = this.currentDataView.metadata && this.currentDataView.metadata.isDataFilterApplied === true;

        // Process current visible rows FIRST to get currentBatchExternalIds
        // These are the filtered ExternalIds if filter is applied, or all ExternalIds if not
        const currentBatchExternalIds: string[] = [];
        const currentRowSet = new Set<string>(); // Track unique ExternalIds in current batch

        dataView.table.rows.forEach((row, rowIndex) => {
            const selectionIdBuilder = this.host.createSelectionIdBuilder();
            if (externalIdsIndex !== -1) {
                selectionIdBuilder.withTable(dataView.table, rowIndex);
            }
            const selectionId = selectionIdBuilder.createSelectionId();
            const externalIdValue = externalIdsIndex !== -1 ? row[externalIdsIndex] : null;
            const colorValue = colorIndex !== -1 ? String(row[colorIndex]) : null;

            if (externalIdValue != null) {
                const key = selectionId.getKey();
                const externalIdString = String(externalIdValue);

                // Update or add to elementDataMap (always, for selection mapping)
                this.elementDataMap.set(key, {
                    id: externalIdString, // This is an ExternalId
                    color: colorValue,
                    selectionId: selectionId
                });

                // Add to current batch ExternalIds (for filtering) - only if not already added
                if (!currentRowSet.has(externalIdString)) {
                    currentBatchExternalIds.push(externalIdString);
                    currentRowSet.add(externalIdString);
                }
            }
        });

        // CRITICAL: Only reset on Create if there's NO filter applied AND we have full dataset
        // When Power BI applies a filter, it sends Create operation, but we need to preserve the full dataset
        // Check if current batch is a subset (filtered) or full dataset (unfiltered)
        const isCurrentBatchFullDatasetForReset = 
            this.externalIds.length === 0 || // No data yet, or
            (currentBatchExternalIds.length === this.externalIds.length && 
             this.arraysEqual(currentBatchExternalIds.sort(), this.externalIds.sort())); // Same IDs
        
        // Save state before potential reset
        const hadFilterBeforeReset = this.lastFilteredIds !== null && this.lastFilteredIds.length > 0;
        
        if (isCreate && !isDataFilterApplied && isCurrentBatchFullDatasetForReset) {
            // Fresh start - no filter, full dataset, so reset everything
            console.log('Visual: Create operation without filter - resetting all data');
            this.allRows = [];
            this.elementDataMap.clear();
            this.externalIds = [];
            this.allDbIds = null;
            this.dbIdToExternalIdMap.clear();
            // Don't reset lastFilteredIds if viewer is still isolated - we need to detect filter clear
            // Only reset if we're truly starting fresh (no viewer isolation)
            if (!this.isViewerIsolated) {
                this.lastFilteredIds = null;
            }
        } else if (isCreate && (isDataFilterApplied || !isCurrentBatchFullDatasetForReset)) {
            // Create with filter OR Create with subset of data - preserve existing full dataset
            console.log('Visual: Create operation WITH filter or subset data - preserving existing full dataset');
            // Don't reset allRows, externalIds, or allDbIds
            // Don't reset lastFilteredIds - we'll detect the filter change below
        }

        // Accumulate rows and ExternalIds - ONLY when NOT filtered
        // This builds the full dataset during initial load and pagination
        if (!isDataFilterApplied) {
            // No filter: accumulate all rows to build full dataset (handles pagination)
            this.allRows = this.allRows.concat(dataView.table.rows);

            // Accumulate unique ExternalIds
            currentBatchExternalIds.forEach(externalId => {
                if (!this.externalIds.includes(externalId)) {
                    this.externalIds.push(externalId);
                }
            });

            // Update allDbIds with accumulated externalIds (represents full dataset)
            this.allDbIds = this.externalIds.slice(); // Create a copy of full dataset

            console.log(`Visual: Accumulated data - allRows: ${this.allRows.length}, externalIds: ${this.externalIds.length}`);
        }
        // When filtered: 
        // - allRows and externalIds remain as the last known full dataset (preserved from above)
        // - currentBatchExternalIds contains the filtered ExternalIds (use these for highlighting)

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

        // 8. INCOMING FILTER LOGIC (Power BI -> Viewer)
        // Enhanced logic to detect filters even when isDataFilterApplied is unreliable
        
        // Determine if we're still paginating (Create is NOT pagination unless it has segment, Append is pagination)
        // In Power BI, 'Create' operation usually means "new data view", which can be initial load or a filter change.
        // 'Append' means more data for the same view (pagination).
        const isPaginating = isAppend || hasSegment;

        // Detect filter by row count: if we have less rows than total known externalIds, it might be a filter
        const hasFilterByCount =
            currentBatchExternalIds.length > 0 &&
            this.externalIds.length > 0 &&
            currentBatchExternalIds.length < this.externalIds.length;

        // Detect if current batch is a subset (filtered) even if count matches
        // This handles cases where Power BI sends filtered data but count might be same
        const isCurrentBatchSubset = 
            this.externalIds.length > 0 &&
            currentBatchExternalIds.length > 0 &&
            currentBatchExternalIds.length < this.externalIds.length &&
            !this.arraysEqual(currentBatchExternalIds.sort(), this.externalIds.sort());
        
        // Detect if current batch matches full dataset (no filter, all elements)
        const isCurrentBatchFullDataset = 
            this.externalIds.length > 0 &&
            currentBatchExternalIds.length === this.externalIds.length &&
            this.arraysEqual(currentBatchExternalIds.sort(), this.externalIds.sort());
        
        // Detect if ExternalIds changed from last filtered set
        // BUT: Only consider it a filter change if it's NOT the full dataset
        const externalIdsChanged = 
            this.lastFilteredIds !== null && 
            this.lastFilteredIds.length > 0 &&
            !isCurrentBatchFullDataset && // Don't consider full dataset as a filter change
            (currentBatchExternalIds.length !== this.lastFilteredIds.length ||
             !this.arraysEqual(currentBatchExternalIds.sort(), this.lastFilteredIds.sort()));

        // Determine if there's an external filter (from another visual)
        // External filter = we didn't initiate it AND:
        //   - Data is filtered (metadata) OR
        //   - Row count is reduced (subset) OR
        //   - ExternalIds changed AND it's a subset (not full dataset)
        // We exclude pagination to avoid flickering during data load
        // CRITICAL: Full dataset (99 elements) should NEVER be considered a filter
        const isExternalFilter =
            !this.isDbIdSelectionActive &&
            !isPaginating &&
            !isCurrentBatchFullDataset && // NEVER treat full dataset as a filter
            (isDataFilterApplied || hasFilterByCount || isCurrentBatchSubset || externalIdsChanged);

        // Get the ExternalIds to isolate
        const externalIdsToIsolate = isExternalFilter ? currentBatchExternalIds : [];

        // Detect state transitions
        const wasFiltered = this.lastFilteredIds !== null && this.lastFilteredIds.length > 0;

        // Filter cleared = was filtered OR viewer is isolated, and now:
        //   - No filter metadata AND
        //   - Current batch matches full dataset (all elements restored) AND
        //   - NOT a subset (not filtered)
        // Also detect if viewer is isolated but Power BI has no filter (needs restoration)
        const isFilterCleared =
            (wasFiltered || this.isViewerIsolated) &&
            !this.isDbIdSelectionActive &&
            !isDataFilterApplied &&
            !hasFilterByCount &&
            !isCurrentBatchSubset &&
            isCurrentBatchFullDataset; // Current batch matches full dataset = filter cleared

        // Enhanced logging
        console.log(`Visual: ========== FILTER DETECTION DEBUG ==========`);
        console.log(`Visual: operationKind: ${options.operationKind}, type: ${options.type}`);
        console.log(`Visual: isCreate: ${isCreate}, isAppend: ${isAppend}, hasSegment: ${hasSegment}`);
        console.log(`Visual: isDataFilterApplied (from metadata): ${isDataFilterApplied}`);
        console.log(`Visual: isDbIdSelectionActive (internal flag): ${this.isDbIdSelectionActive}`);
        console.log(`Visual: isPaginating: ${isPaginating}`);
        console.log(`Visual: hasFilterByCount: ${hasFilterByCount} (current: ${currentBatchExternalIds.length}, total: ${this.externalIds.length})`);
        console.log(`Visual: Filter detection - isExternalFilter: ${isExternalFilter}`);
        console.log(`Visual: Filter detection - hasFilterByCount: ${hasFilterByCount}, externalIdsChanged: ${externalIdsChanged}, isCurrentBatchSubset: ${isCurrentBatchSubset}`);
        console.log(`Visual: State transition - wasFiltered: ${wasFiltered}, isFilterCleared: ${isFilterCleared}`);
        console.log(`Visual: Data counts - externalIdsToIsolate: ${externalIdsToIsolate.length}, currentBatchExternalIds: ${currentBatchExternalIds.length}, allRows: ${this.allRows.length}, externalIds: ${this.externalIds.length}`);
        if (this.lastFilteredIds) {
            console.log(`Visual: Last filtered ExternalIds count: ${this.lastFilteredIds.length}`);
        }
        console.log(`Visual: Viewer state - viewer exists: ${!!this.viewer}, isViewerReady: ${this.isViewerReady}, idMapping exists: ${!!this.idMapping}`);
        console.log(`Visual: ============================================`);

        if (isDataFilterApplied && currentBatchExternalIds.length > 0) {
            console.log(`Visual: Filtered data - Total ExternalIds: ${currentBatchExternalIds.length}, Sample (first 10):`, currentBatchExternalIds.slice(0, 10));
        }

        if (currentBatchExternalIds.length > 0 && currentBatchExternalIds.length !== this.externalIds.length) {
            console.log(`Visual: POTENTIAL FILTER: currentBatchExternalIds (${currentBatchExternalIds.length}) differs from externalIds (${this.externalIds.length})`);
        }

        // Wait for viewer to be ready before processing filters
        if (!this.viewer || !this.isViewerReady || !this.idMapping) {
            // Store pending selection to apply when viewer is ready
            if (isExternalFilter && externalIdsToIsolate.length > 0) {
                console.log(`Visual: Viewer not ready, storing ${externalIdsToIsolate.length} ExternalIds for later application...`);
                this.pendingSelection = externalIdsToIsolate;
                this.lastFilteredIds = externalIdsToIsolate;
            }
            return; // Exit early if viewer not ready
        }

        // STATE 1: External filter active - Isolate and highlight filtered elements
        // We check !isPaginating to avoid isolating partial data during load
        if (isExternalFilter && externalIdsToIsolate.length > 0 && !isPaginating) {
            console.log(`Visual: External filter detected! Processing ${externalIdsToIsolate.length} ExternalIds to isolate.`);

            // Set flag to prevent selection loop
            this.isProgrammaticSelection = true;

            // Reset isDbIdSelectionActive when filter comes from external source
            this.isDbIdSelectionActive = false;

            // Apply filter (isolate, highlight with neon green, and fit to view)
            // Pass ExternalIds, syncSelectionState will convert them to dbIds internally
            await this.syncSelectionState(true, externalIdsToIsolate);

            // Track filtered ExternalIds for state transition detection
            this.lastFilteredIds = externalIdsToIsolate.slice();
            this.isViewerIsolated = true; // Mark viewer as isolated

            // Reset flag after operation
            this.isProgrammaticSelection = false;

            console.log(`Visual: Filter applied successfully, ${externalIdsToIsolate.length} elements isolated and highlighted`);
        }
        // STATE 2: Filter cleared - Show all elements and restore original colors
        else if (isFilterCleared) {
            console.log(`Visual: Filter cleared detected! Showing all elements and restoring original state...`);

            // Set flag to prevent selection loop
            this.isProgrammaticSelection = true;

            // CRITICAL: Clear ALL theming colors first (including nuclear green)
            // Clear isolation and selection
            this.viewer.clearSelection();
            
            // Clear theming colors multiple times to ensure nuclear green is completely removed
            // Sometimes clearThemingColors needs to be called multiple times to fully clear
            this.viewer.clearThemingColors(this.model);
            this.viewer.clearThemingColors(this.model);
            
            // Show all elements (removes isolation)
            showAll(this.viewer, this.model);
            
            // Force another clear after showAll to ensure all colors are removed
            this.viewer.clearThemingColors(this.model);
            
            // Additional forced clear: get all visible nodes and clear them individually if needed
            try {
                const tree = this.model.getInstanceTree();
                if (tree) {
                    // Get all dbIds in the model
                    const allDbIds: number[] = [];
                    tree.enumNodeChildren(tree.getRootId(), (dbid) => {
                        allDbIds.push(dbid);
                    }, true);
                    
                    // Clear theming for all nodes individually to ensure complete removal
                    allDbIds.forEach(dbid => {
                        try {
                            // Use clearThemingColor if available, otherwise rely on clearThemingColors
                            if ((this.viewer as any).clearThemingColor) {
                                (this.viewer as any).clearThemingColor(dbid, this.model);
                            }
                        } catch (e) {
                            // Ignore errors for individual nodes
                        }
                    });
                    
                    // Final clearThemingColors call
                    this.viewer.clearThemingColors(this.model);
                }
            } catch (e) {
                console.warn('Visual: Error during forced color clear:', e);
            }

            // Reset state tracking
            this.lastFilteredIds = null;
            this.isDbIdSelectionActive = false;
            this.isViewerIsolated = false; // Mark viewer as not isolated

            // Small delay to ensure theming colors are completely cleared
            await new Promise(resolve => setTimeout(resolve, 100));

            // Restore original colors from data ONLY if colors are defined in Power BI
            // If no colors are defined, elements should remain without theming
            const hasColorsInData = Array.from(this.elementDataMap.values()).some(v => v.color != null);
            if (hasColorsInData) {
                await this.syncColors();
            } else {
                console.log('Visual: No colors defined in data, keeping elements without theming');
            }

            // Reset flag after operation
            this.isProgrammaticSelection = false;

            console.log(`Visual: Filter cleared successfully, showing all ${this.externalIds.length} elements without green highlight`);
        }
        // STATE 3: No filter (initial state or no change) - Sync colors if needed
        else if (!isDataFilterApplied && !this.isDbIdSelectionActive && !isPaginating && !this.isViewerIsolated) {
            // Only sync colors if we're not in the middle of pagination
            // and there's no active selection from viewer or external filter
            // and viewer is not isolated (to avoid applying colors when restoring)
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

            // CRITICAL: Build reverse mapping (dbId -> ExternalId) for all loaded data
            // This allows us to map selected dbIds back to ExternalIds in Power BI column
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
     * Converts selected dbIds to ExternalIds and applies filter to Power BI
     */
    private async handleDbIds(selectedDbIds: number[]) {
        if (!this.host) return;

        // Skip if this selection change came from programmatic selection (external filter)
        if (this.isProgrammaticSelection) {
            console.log("Visual: Ignoring programmatic selection change:", selectedDbIds.length, "dbIds");
            return;
        }

        console.log("Visual: Selection changed in viewer (user interaction):", selectedDbIds);

        // 1. If no selection, clear filters
        if (selectedDbIds.length === 0) {
            this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.merge);
            this.isDbIdSelectionActive = false;
            return;
        }

        // 2. If selection, apply JSON filter using ExternalIds
        // ExternalIds column is REQUIRED - no fallback
        if (!this.currentDataView || !this.currentDataView.table) return;

        const columns = this.currentDataView.table.columns;
        const externalIdsIndex = columns.findIndex(c => c.roles["externalIds"]);
        
        if (externalIdsIndex === -1) {
            console.error("Visual: ExternalIds column is REQUIRED but not found for filtering");
            return;
        }

        const columnSource = columns[externalIdsIndex];
        
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
        
        console.log(`Visual: Filtering using ExternalIds column: ${target.table}.${target.column}`);

        // CRITICAL: Map dbIds to ExternalIds for filtering
        // We work exclusively with ExternalIds - convert dbIds -> ExternalIds
        let filterValues: string[] = [];
        let missingDbIds: number[] = [];

        for (const dbId of selectedDbIds) {
            // 1. Try reverse map (built from buildReverseMapping)
            const externalId = this.dbIdToExternalIdMap.get(dbId);
            if (externalId) {
                // Verify this ExternalId exists in our Power BI data
                if (this.externalIds.includes(externalId)) {
                    filterValues.push(externalId);
                } else {
                    console.warn(`Visual: ExternalId "${externalId}" from mapping not found in Power BI data, skipping`);
                    missingDbIds.push(dbId);
                }
            } else {
                // 2. Fallback: try to get ExternalId directly from model
                if (this.idMapping) {
                    try {
                        const externalIdsFromModel = await this.idMapping.getExternalIds([dbId]);
                        if (externalIdsFromModel.length > 0 && externalIdsFromModel[0]) {
                            const modelExternalId = externalIdsFromModel[0];
                            // Check if this ExternalId exists in our Power BI column
                            if (this.externalIds.includes(modelExternalId)) {
                                filterValues.push(modelExternalId);
                                // Update mapping for future use
                                this.dbIdToExternalIdMap.set(dbId, modelExternalId);
                            } else {
                                console.warn(`Visual: ExternalId "${modelExternalId}" from model not found in Power BI data, skipping`);
                                missingDbIds.push(dbId);
                            }
                        } else {
                            console.warn(`Visual: No ExternalId found for dbId ${dbId}, skipping`);
                            missingDbIds.push(dbId);
                        }
                    } catch (e) {
                        console.warn(`Visual: Failed to get ExternalId for dbId ${dbId}:`, e);
                        missingDbIds.push(dbId);
                    }
                } else {
                    console.warn(`Visual: No idMapping available for dbId ${dbId}, skipping`);
                    missingDbIds.push(dbId);
                }
            }
        }

        console.log(`Visual: Mapped ${selectedDbIds.length} dbIds to ${filterValues.length} ExternalIds for filtering`);

        if (missingDbIds.length > 0) {
            console.warn(`Visual: Warning - ${missingDbIds.length} selected dbIds could not be mapped to ExternalIds in Power BI data. Filter will ignore them.`);
            if (missingDbIds.length <= 10) {
                console.warn(`Visual: Missing dbIds:`, missingDbIds);
            }
        }

        if (filterValues.length > 0) {
            console.log(`Visual: Filter values (ExternalIds, first 10):`, filterValues.slice(0, 10));

            const filter = new models.BasicFilter(
                target,
                "In",
                filterValues
            );

            console.log("Visual: Applying JSON filter using ExternalIds:", filter);
            this.host.applyJsonFilter(filter, "general", "filter", powerbi.FilterAction.merge);
            this.isDbIdSelectionActive = true;
        } else {
            console.warn("Visual: No valid ExternalIds found for selection. Skipping filter application to avoid empty results.");
            // We do NOT clear filter here to avoid resetting other visuals to 'All' when clicking empty space or invalid objects.
            // If user wants to clear, they click empty space -> selectedDbIds.length === 0 handled above.
        }
    }

    /**
     * Build reverse mapping from dbId (number) to ExternalId (string)
     * This is critical for filtering: when user selects an element in viewer (dbId),
     * we need to map it back to the ExternalId value in Power BI column
     * 
     * Since we work exclusively with ExternalIds, this mapping allows us to:
     * 1. Convert dbIds from viewer selection -> ExternalIds for Power BI filtering
     * 2. Ensure we only filter by ExternalIds that exist in our Power BI data
     */
    private async buildReverseMapping(): Promise<void> {
        if (!this.idMapping || !this.externalIds || this.externalIds.length === 0) {
            console.log("Visual: buildReverseMapping - No data available yet");
            return;
        }

        console.log(`Visual: buildReverseMapping - Building reverse mapping (dbId -> ExternalId) for ${this.externalIds.length} ExternalIds from column`);

        try {
            // Map all ExternalIds from Power BI column to actual dbIds (numbers) in the model
            // Use mapExternalIdsToDbIds to convert ExternalIds -> dbIds using External ID mapping
            const dbIds = await this.idMapping.mapExternalIdsToDbIds(this.externalIds);

            // Build reverse mapping: dbId (number) -> ExternalId (string from Power BI column)
            // Clear existing map to avoid stale data
            this.dbIdToExternalIdMap.clear();

            // Get ExternalIds from model for all dbIds
            // This ensures we're mapping to the actual ExternalIds in the model
            const externalIdsFromModel = await this.idMapping.getExternalIds(dbIds);
            
            // Build the mapping: dbId -> ExternalId (from Power BI column)
            for (let i = 0; i < dbIds.length; i++) {
                const dbId = dbIds[i];
                if (dbId != null) {
                    const modelExternalId = externalIdsFromModel[i];
                    if (modelExternalId) {
                        // Find the ExternalId in our Power BI column that matches this model ExternalId
                        // This ensures we map to the exact value that exists in Power BI
                        if (this.externalIds.includes(modelExternalId)) {
                            this.dbIdToExternalIdMap.set(dbId, modelExternalId);
                        } else {
                            console.warn(`Visual: buildReverseMapping - Model ExternalId "${modelExternalId}" not found in Power BI column`);
                        }
                    }
                }
            }
            
            console.log(`Visual: buildReverseMapping - Built reverse mapping (dbId -> ExternalId) for ${this.dbIdToExternalIdMap.size} dbIds`);
            
            if (this.dbIdToExternalIdMap.size > 0) {
                const sampleEntries = Array.from(this.dbIdToExternalIdMap.entries()).slice(0, 5);
                console.log(`Visual: buildReverseMapping - Sample mappings:`, sampleEntries);
            }
            
            if (this.dbIdToExternalIdMap.size < dbIds.length) {
                console.warn(`Visual: buildReverseMapping - Only mapped ${this.dbIdToExternalIdMap.size} of ${dbIds.length} dbIds. Some elements may not have ExternalIds or don't match Power BI data.`);
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

    private async syncSelectionState(isolate: boolean, externalIdsToIsolate: string[]) {
        // Verify viewer and model are ready
        if (!this.viewer || !this.model || !this.idMapping) {
            console.log(`Visual: syncSelectionState - Viewer/model not ready, storing pending selection for ${externalIdsToIsolate.length} ExternalIds`);
            if (isolate && externalIdsToIsolate.length > 0) {
                this.pendingSelection = externalIdsToIsolate;
            }
            return;
        }

        // Apply any pending selection if viewer is now ready
        if (this.pendingSelection && this.pendingSelection.length > 0) {
            console.log(`Visual: syncSelectionState - Applying pending selection for ${this.pendingSelection.length} ExternalIds`);
            const pendingExternalIds = this.pendingSelection;
            this.pendingSelection = null;
            await this.syncSelectionState(true, pendingExternalIds);
            return;
        }

        // PATH 1: Isolate and highlight specific elements
        if (isolate && externalIdsToIsolate.length > 0) {
            try {
                console.log(`Visual: syncSelectionState - Attempting to isolate and highlight ${externalIdsToIsolate.length} ExternalIds`);
                console.log(`Visual: syncSelectionState - Input ExternalIds (first 10):`, externalIdsToIsolate.slice(0, 10));
                console.log(`Visual: syncSelectionState - Input ExternalIds (last 10):`, externalIdsToIsolate.slice(-10));

                // Convert ExternalIds to dbIds using external ID mapping
                // Use mapExternalIdsToDbIds to convert ExternalIds -> dbIds using External ID mapping
                const validDbIds = await this.idMapping.mapExternalIdsToDbIds(externalIdsToIsolate);

                console.log(`Visual: syncSelectionState - Mapping result: ${validDbIds.length} valid dbIds from ${externalIdsToIsolate.length} ExternalIds`);

                // Build mapping from ExternalIds to dbIds for reference
                // This helps us track which ExternalIds were successfully mapped
                const externalIdToDbIdMap = new Map<string, number>();
                const externalIdsFromModel = await this.idMapping.getExternalIds(validDbIds);
                
                for (let i = 0; i < validDbIds.length; i++) {
                    const dbId = validDbIds[i];
                    const modelExternalId = externalIdsFromModel[i];
                    if (modelExternalId && externalIdsToIsolate.includes(modelExternalId)) {
                        externalIdToDbIdMap.set(modelExternalId, dbId);
                        // Update reverse mapping for future use
                        this.dbIdToExternalIdMap.set(dbId, modelExternalId);
                    }
                }

                console.log(`Visual: syncSelectionState - Built ExternalId to dbId mapping: ${externalIdToDbIdMap.size} entries`);

                console.log(`Visual: syncSelectionState - Mapped ${externalIdsToIsolate.length} ExternalIds to ${validDbIds.length} valid dbIds`);

                // DEBUG: Log actual dbIds to see what we're trying to isolate
                if (validDbIds.length > 0) {
                    console.log(`Visual: syncSelectionState - Mapped dbIds (first 10):`, validDbIds.slice(0, 10));
                    console.log(`Visual: syncSelectionState - Mapped dbIds (last 10):`, validDbIds.slice(-10));
                }

                if (validDbIds.length === 0) {
                    console.warn(`Visual: syncSelectionState - None of the ${externalIdsToIsolate.length} ExternalIds could be mapped to valid dbIds in the model.`);
                    console.warn(`Visual: syncSelectionState - Sample ExternalIds attempted:`, externalIdsToIsolate.slice(0, 20));
                    // Show all if no valid IDs
                    this.viewer.clearSelection();
                    this.viewer.clearThemingColors(this.model);
                    showAll(this.viewer, this.model);
                    return;
                }

                // Log mapping efficiency
                if (validDbIds.length < externalIdsToIsolate.length) {
                    console.warn(`Visual: syncSelectionState - Only ${validDbIds.length} of ${externalIdsToIsolate.length} ExternalIds were mapped to valid dbIds.`);
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

                console.log(`Visual: syncSelectionState - Final verified DbIds: ${verifiedDbIds.length}`);

                // CRITICAL DEBUG: Check if we're trying to isolate container nodes
                if (containerNodes.length > 0) {
                    console.warn(`Visual: syncSelectionState - WARNING: ${containerNodes.length} of ${verifiedDbIds.length} DbIds are container nodes (have children)`);
                    console.warn(`Visual: syncSelectionState - Container nodes (first 5):`, containerNodes.slice(0, 5));
                    console.warn(`Visual: syncSelectionState - Expanding container nodes to include all children for visibility`);
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

                    // Step 3: ISOLATION (Model Explorer behavior)
                    // We use expanded IDs to ensure geometry is visible
                    console.log(`Visual: syncSelectionState - Isolating ${finalDbIds.length} elements (hiding all others)`);
                    isolateDbIds(this.viewer, this.model, finalDbIds);

                    // Step 4: NO selecting elements (only isolate + focus for incoming filters)
                    // We skip selection to avoid triggering outgoing filter events from programmatic actions
                    console.log(`Visual: syncSelectionState - Skipping selection (incoming filter, not user click)`);
                    // this.viewer.select(finalDbIds, this.model); // REMOVED - no programmatic selection

                    // Step 5: Highlight nodes with NUCLEAR GREEN color (highly visible)
                    // Define color inline to avoid THREE is not defined error at module load time
                    // Nuclear green: #00FF41 (Matrix-style, RGB: 0, 255, 65)
                    const NUCLEAR_GREEN = new THREE.Vector4(0.0, 1.0, 0.255, 1.0);
                    console.log(`Visual: syncSelectionState - Theming ${finalDbIds.length} nodes with NUCLEAR GREEN`);

                    // Apply nuclear green color to each node (expanded)
                    finalDbIds.forEach(dbid => {
                        this.viewer.setThemingColor(dbid, NUCLEAR_GREEN, this.model);
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

                    console.log(`Visual: syncSelectionState - Successfully isolated and highlighted ${finalDbIds.length} elements with NUCLEAR GREEN (NO selection)`);
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
                // Use mapExternalIdsToDbIds to convert ExternalIds -> dbIds
                const validDbIds = await this.idMapping.mapExternalIdsToDbIds(ids);

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
                // Use mapExternalIdsToDbIds to convert ExternalIds -> dbIds
                const validDbIds = await this.idMapping.mapExternalIdsToDbIds(ids);

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

    /**
     * Helper method to compare two arrays for equality
     */
    private arraysEqual(a: string[], b: string[]): boolean {
        if (a.length !== b.length) return false;
        for (let i = 0; i < a.length; i++) {
            if (a[i] !== b[i]) return false;
        }
        return true;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}

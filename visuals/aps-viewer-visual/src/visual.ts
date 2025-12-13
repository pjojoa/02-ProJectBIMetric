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
import {
    IdMapping,
    launchViewer,
    showAll,
    isolateDbIds,
    fitToView,
    isolateExternalIds,
    fitToExternalIds,
    getSelectionAsExternalIds
} from "./viewer.utils";
import * as models from "powerbi-models";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import DataView = powerbi.DataView;
import DataViewTableRow = powerbi.DataViewTableRow;

// Strongly-typed aliases to make the ExternalId-only design explicit
type ExternalId = string;
type DbId = number;

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

    // Selection handling improvements
    private selectionDebounceTimer: number | null = null; // Timer for debouncing selection changes
    private isProcessingSelection: boolean = false; // Flag to prevent concurrent selection processing
    private readonly SELECTION_DEBOUNCE_MS: number = 30; // Very short delay to let viewer update selection state after Ctrl+Click

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
    /** Maps SelectionId key -> { externalId, color, selectionId } */
    private elementDataMap: Map<string, { id: ExternalId, color: string | null, selectionId: ISelectionId }> = new Map();
    /** All ExternalIds coming from Power BI (full dataset, accumulated with pagination) */
    private externalIds: ExternalId[] = [];
    /** Reverse mapping: dbId -> ExternalId (used only when we need to go from Viewer selection back to ExternalId) */
    private dbIdToColumnValueMap: Map<DbId, ExternalId> = new Map();
    /** Reverse mapping: ExternalId -> row index in Power BI table (for validation and selection) */
    private externalIdToRowIndexMap: Map<ExternalId, number> = new Map();

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
            if (externalIdsIndex !== -1) {
                selectionIdBuilder.withTable(dataView.table, rowIndex);
            }
            const selectionId = selectionIdBuilder.createSelectionId();
            const externalIdValue = externalIdsIndex !== -1 ? row[externalIdsIndex] : null;
            const colorValue = colorIndex !== -1 ? String(row[colorIndex]) : null;

            if (externalIdValue != null) {
                const key = selectionId.getKey();
                const externalId: ExternalId = String(externalIdValue).trim();
                if (!externalId) {
                    return;
                }

                // Update or add to elementDataMap (always, for selection mapping)
                this.elementDataMap.set(key, {
                    id: externalId,
                    color: colorValue,
                    selectionId: selectionId
                });

                // Track reverse mapping ExternalId -> row index
                if (!this.externalIdToRowIndexMap.has(externalId)) {
                    this.externalIdToRowIndexMap.set(externalId, rowIndex);
                }

                // Add to current batch IDs (for filtering) - only if not already added
                if (!currentRowSet.has(externalId)) {
                    currentBatchIds.push(externalId);
                    currentRowSet.add(externalId);
                }
            }
        });

        // Accumulate rows and IDs - ONLY when NOT filtered
        // This builds the full dataset during initial load and pagination
        if (!isDataFilterApplied) {
            // No filter: accumulate all rows to build full dataset (handles pagination)
            this.allRows = this.allRows.concat(dataView.table.rows);

            // Accumulate unique ExternalIds - OPTIMIZED
            // Using Set for O(1) lookups instead of O(N) array.includes
            const existingIdsSet = new Set(this.externalIds);
            let newIdsAdded = false;

            currentBatchIds.forEach(externalId => {
                if (!existingIdsSet.has(externalId)) {
                    this.externalIds.push(externalId);
                    existingIdsSet.add(externalId); // Keep Set in sync for this loop
                    newIdsAdded = true;
                }
            });

            // Update allDbIds with accumulated externalIds (represents full dataset)
            this.allDbIds = this.externalIds.slice(); // Create a copy of full dataset

            if (newIdsAdded) {
                console.log(`Visual: Accumulated data - Total Rows: ${this.allRows.length}, Unique IDs: ${this.externalIds.length}`);
            }
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

        // 6. Handle Pagination Fetch - AGGRESSIVE
        // We want to fetch ALL data to ensure the viewer can map everything
        let isFetching = false;
        const HARD_LIMIT = 200000; // Safety cap to prevent browser crash

        if (this.currentDataView.metadata.segment) {
            if (this.allRows.length < HARD_LIMIT) {
                const moreData = this.host.fetchMoreData();
                if (moreData) {
                    isFetching = true;
                    this.statusDiv.innerText = `Loading data... (${this.allRows.length} rows loaded)`;
                    this.statusDiv.style.backgroundColor = '#d9534f'; // Red/Orange for "Busy"
                    this.statusDiv.style.color = 'white';
                }
            } else {
                this.statusDiv.innerText = `Warning: Data limit reached (${HARD_LIMIT} rows). Some objects may not be selectable.`;
                this.statusDiv.style.backgroundColor = 'orange';
                this.statusDiv.style.color = 'black';
            }
        } else {
            // Pagination complete
            const uniqueIdCount = this.externalIds.length;
            this.statusDiv.innerText = `Ready | Rows: ${this.allRows.length} | Unique Objects: ${uniqueIdCount}`;
            this.statusDiv.style.backgroundColor = '#333';
            this.statusDiv.style.color = 'lightgreen';
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
            // Use debounced version to improve responsiveness and prevent race conditions
            const result = await launchViewer(
                this.container,
                this.currentUrn,
                token.access_token,
                this.currentGuid,
                () => this.handleSelectionChangeDebounced(), // Debounced callback for selection
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
     * CRITICAL: The Power BI column contains External IDs, not dbIds.
     * We need to convert dbIds (from viewer selection) to External IDs (for filtering).
     */
    /**
     * Handles selection changes from the Viewer (Viewer -> Power BI)
     * Uses the new ExternalID-first pattern.
     * Now with debouncing and concurrent call prevention for better responsiveness.
     */
    private async handleSelectionChange() {
        if (!this.host || !this.viewer || !this.idMapping) return;

        // Skip if this selection change came from programmatic selection (external filter)
        if (this.isProgrammaticSelection) {
            // console.log("Visual: Ignoring programmatic selection change");
            return;
        }

        // Prevent concurrent processing
        if (this.isProcessingSelection) {
            console.log("Visual: Selection change already being processed, skipping...");
            return;
        }

        this.isProcessingSelection = true;

        try {
            // 1. Get current selection from viewer - this should include ALL selected elements
            // The viewer maintains all selected elements when using Ctrl+Click
            const currentDbIds = this.viewer.getSelection();
            console.log(`Visual: Raw selection from viewer: ${currentDbIds.length} dbIds`, currentDbIds.slice(0, 10), currentDbIds.length > 10 ? `... and ${currentDbIds.length - 10} more` : '');

            // 2. Convert dbIds to ExternalIDs
            const selectedExternalIds = await getSelectionAsExternalIds(this.viewer, this.idMapping);

            if (selectedExternalIds.length > 1) {
                console.log(`Visual: Multi-selection detected! Selected ${selectedExternalIds.length} elements. ExternalIDs:`, selectedExternalIds.slice(0, 10), selectedExternalIds.length > 10 ? `... and ${selectedExternalIds.length - 10} more` : '');
            } else if (selectedExternalIds.length === 1) {
                console.log(`Visual: Single selection. ExternalID: ${selectedExternalIds[0]}`);
            } else {
                console.log(`Visual: Selection cleared (no elements selected)`);
            }

            // Debug: Check if dbIds count matches ExternalIds count
            if (currentDbIds.length !== selectedExternalIds.length) {
                console.warn(`Visual: Selection count mismatch! dbIds: ${currentDbIds.length}, ExternalIds: ${selectedExternalIds.length}`);
            }

            // 2. If no selection, clear filters
            if (selectedExternalIds.length === 0) {
                console.log("Visual: No selection - clearing all filters");
                this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.merge);
                this.isDbIdSelectionActive = false;

                // Also clear selection in SelectionManager
                try {
                    this.selectionManager.clear();
                } catch (clearError) {
                    console.warn("Visual: Could not clear selection via SelectionManager:", clearError);
                }
                return;
            }

            // 3. Validate that we have the necessary components
            if (!this.currentDataView || !this.currentDataView.table) {
                console.warn("Visual: No data view available for filtering");
                return;
            }

            const columns = this.currentDataView.table.columns;
            const externalIdsIndex = columns.findIndex(c => c.roles["externalIds"]);

            if (externalIdsIndex === -1) {
                console.warn("Visual: 'externalIds' role column not found");
                return;
            }

            const columnSource = columns[externalIdsIndex];

            // 4. Get table and column names for filter target
            let target: models.IFilterColumnTarget;

            if (columnSource.queryName) {
                const parts = columnSource.queryName.split('.');
                if (parts.length >= 2) {
                    target = {
                        table: parts[0],
                        column: parts[1]
                    };
                } else {
                    target = {
                        table: columnSource.queryName,
                        column: columnSource.displayName
                    };
                }
            } else {
                target = {
                    table: "Data",
                    column: columnSource.displayName
                };
            }

            // 5. Verify that External IDs exist in our known dataset
            // This ensures we only filter with IDs that exist in Power BI data
            const filterValues: string[] = [];
            const unknownExternalIds: string[] = [];

            for (const extId of selectedExternalIds) {
                // Check if this External ID exists in our dataset
                if (this.externalIds.includes(extId)) {
                    filterValues.push(extId);
                } else {
                    unknownExternalIds.push(extId);
                }
            }

            if (unknownExternalIds.length > 0) {
                // console.warn(`Visual: ${unknownExternalIds.length} Unknown External IDs (not in dataset). Sample:`, unknownExternalIds.slice(0, 5));
            }

            // 6. Apply filter if we have valid values
            if (filterValues.length > 0) {
                if (filterValues.length > 1) {
                    console.log(`Visual: Applying multi-selection filter with ${filterValues.length} External IDs`);
                    console.log(`Visual: Filter will show data for ${filterValues.length} selected elements`);
                } else {
                    console.log(`Visual: Applying filter with 1 External ID: ${filterValues[0]}`);
                }

                const filter = new models.BasicFilter(
                    target,
                    "In",
                    filterValues
                );

                // CRITICAL: Use FilterAction.merge to allow cross-filtering with other visuals
                // This filter will show rows where the External ID column matches ANY of the selected IDs
                this.host.applyJsonFilter(filter, "general", "filter", powerbi.FilterAction.merge);
                this.isDbIdSelectionActive = true;

                console.log(`Visual: Filter applied successfully! Power BI will now filter other visuals to show data for ${filterValues.length} selected element(s)`);

                // Also update selection manager to ensure proper cross-visual interaction
                try {
                    // Build selection IDs for the filtered items to maintain selection state
                    const selectionIds: ISelectionId[] = [];
                    for (const extId of filterValues) {
                        // Find the selection ID for this External ID
                        this.elementDataMap.forEach((value, key) => {
                            if (value.id === extId) {
                                selectionIds.push(value.selectionId);
                            }
                        });
                    }

                    if (selectionIds.length > 0) {
                        // Apply selection to maintain state consistency
                        this.selectionManager.select(selectionIds, true); // true = multiSelect
                    }
                } catch (selectionError) {
                    console.warn("Visual: Could not apply selection via SelectionManager:", selectionError);
                }
            } else {
                console.warn("Visual: No valid External IDs found in dataset. Cannot apply filter.");
            }
        } finally {
            // Always release the processing flag
            this.isProcessingSelection = false;
        }
    }

    /**
     * Debounced version of handleSelectionChange to prevent rapid-fire calls
     * Uses requestAnimationFrame + small timeout to ensure viewer has updated
     * its selection state, especially important for Ctrl+Click multi-selection.
     */
    private handleSelectionChangeDebounced() {
        // Clear any existing timer
        if (this.selectionDebounceTimer !== null) {
            clearTimeout(this.selectionDebounceTimer);
        }

        // Use requestAnimationFrame to wait for the next browser frame,
        // then add a small delay to ensure viewer has updated selection state
        // This is critical for Ctrl+Click where viewer needs to add element to selection array
        requestAnimationFrame(() => {
            this.selectionDebounceTimer = window.setTimeout(() => {
                this.handleSelectionChange();
                this.selectionDebounceTimer = null;
            }, this.SELECTION_DEBOUNCE_MS);
        });
    }

    /**
     * Build reverse mapping from dbId (number) to External ID (string)
     * CRITICAL: The Power BI column contains External IDs, not dbIds.
     * This mapping allows us to quickly look up External IDs from dbIds.
     * However, the primary method for filtering is getExternalIds() which is more reliable.
     */
    private async buildReverseMapping(): Promise<void> {
        if (!this.idMapping || !this.externalIds || this.externalIds.length === 0) {
            console.log("Visual: buildReverseMapping - No data available yet");
            return;
        }

        console.log(`Visual: buildReverseMapping - Building reverse mapping (dbId → ExternalId) for ${this.externalIds.length} ExternalIds from column`);

        try {
            // Step 1: Convert ExternalIds (from Power BI column) to dbIds
            console.log("Visual: buildReverseMapping - Converting ExternalIds to dbIds...");
            const dbIds = await this.idMapping.getDbids(this.externalIds);
            console.log(`Visual: buildReverseMapping - Mapped ${dbIds.length} of ${this.externalIds.length} ExternalIds to dbIds`);

            const allDbIds = dbIds.filter(dbId => dbId != null && !isNaN(dbId));
            if (allDbIds.length === 0) {
                console.warn("Visual: buildReverseMapping - No valid dbIds found. Cannot build reverse mapping.");
                return;
            }

            // Step 2: Build reverse mapping: dbId -> ExternalId using index alignment
            this.dbIdToColumnValueMap.clear();

            for (let i = 0; i < allDbIds.length && i < this.externalIds.length; i++) {
                const dbId = allDbIds[i];
                const externalId = this.externalIds[i];
                if (dbId != null && !isNaN(dbId) && externalId) {
                    this.dbIdToColumnValueMap.set(dbId, externalId);
                }
            }

            console.log(`Visual: buildReverseMapping - Built reverse mapping for ${this.dbIdToColumnValueMap.size} dbIds`);
            if (this.dbIdToColumnValueMap.size > 0) {
                const sampleEntries = Array.from(this.dbIdToColumnValueMap.entries()).slice(0, 5);
                console.log(`Visual: buildReverseMapping - Sample mappings (dbId → ExternalId):`, sampleEntries);
            }
        } catch (e) {
            console.error("Visual: Error building reverse mapping", e);
            console.error("Visual: Stack trace:", e.stack);
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

                // Map ExternalIds (from Power BI) to dbIds for viewer operations
                const validDbIds = await this.idMapping.getDbids(idsToIsolate);

                console.log(`Visual: syncSelectionState - Mapping result: ${validDbIds.length} valid DbIds from ${idsToIsolate.length} input External IDs`);

                // Update reverse mapping: dbId → External ID (for potential future use)
                // This is a best-effort update to keep the cache fresh
                // Note: The primary method for filtering (Viewer → Power BI) uses getExternalIds() directly in handleDbIds()
                try {
                    const externalIdsForDbIds = await this.idMapping.getExternalIds(validDbIds);
                    for (let i = 0; i < validDbIds.length; i++) {
                        if (i < externalIdsForDbIds.length && externalIdsForDbIds[i]) {
                            this.dbIdToColumnValueMap.set(validDbIds[i], externalIdsForDbIds[i]);
                        }
                    }
                    console.log(`Visual: syncSelectionState - Updated reverse mapping for ${validDbIds.length} dbIds`);
                } catch (e) {
                    console.warn("Visual: syncSelectionState - Could not update reverse mapping:", e);
                }

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
                // Map ExternalIds (from Power BI) to dbIds for theming
                const validDbIds = await this.idMapping.getDbids(ids);

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
                // Map ExternalIds (from Power BI) to dbIds for theming
                const validDbIds = await this.idMapping.getDbids(ids);

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

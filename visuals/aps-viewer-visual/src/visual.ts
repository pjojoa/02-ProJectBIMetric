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
import { FormattingSettingsService, formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import ColorPicker = formattingSettings.ColorPicker;
import ToggleSwitch = formattingSettings.ToggleSwitch;
import {
    IdMapping,
    launchViewer,
    showAll,
    isolateDbIds,
    fitToView,
    getSelectionAsExternalIds
} from "./viewer.utils";
import { showLoadingAnimation, hideLoadingAnimation } from "./loading-animation";
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

    // Access Token Endpoint: valor por defecto (oculto de la UI pero usado internamente)
    private accessTokenEndpoint: string = 'https://zero2-projectbimetric.onrender.com/token';
    private currentUrn: string = "";


    private viewer: Autodesk.Viewing.Viewer3D = null;
    private model: Autodesk.Viewing.Model = null;
    private idMapping: IdMapping = null;
    private isViewerReady: boolean = false;

    // Loading animation
    private loadingAnimationHide: (() => void) | null = null;
    private loadingAnimationTimeout: number | null = null;
    private loadingAnimationStartTime: number | null = null;
    private readonly LOADING_ANIMATION_MIN_DURATION: number = 6000; // 6 segundos mínimos (2 segundos más que antes)
    private readonly LOADING_ANIMATION_MAX_DURATION: number = 12000; // 12 segundos máximo como seguridad

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
    /** Maps Category Value -> Hex Color (resolved from settings) */
    private categoryColorMap: Map<string, string> = new Map();


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
    // eslint-disable-next-line max-lines-per-function
    public async update(options: VisualUpdateOptions): Promise<void> {
        // IMPORTANTE: Populate settings ANTES de procesar datos para capturar cambios del usuario
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualSettingsModel, options.dataViews[0]);
        
        // Guardar estado anterior de categoryColorMap para detectar cambios
        const previousCategoryColorMap = new Map(this.categoryColorMap);
        // Access Token Endpoint: usar valor de settings si está disponible, sino usar valor por defecto
        const { accessTokenEndpoint } = this.formattingSettings.viewerCard;
        const endpointValue = accessTokenEndpoint?.value || 'https://zero2-projectbimetric.onrender.com/token';
        if (endpointValue !== this.accessTokenEndpoint) {
            this.accessTokenEndpoint = endpointValue;
            // If endpoint changes, we might need to re-init, but for now just update prop
        }

        this.currentDataView = options.dataViews[0];

        // 1. Get the Table DataView
        const dataView = options.dataViews[0];
        if (!dataView || !dataView.table || !dataView.table.rows) {
            this.currentUrn = '';
            return;
        }

        // 2. Identify Column Indices
        const columns = dataView.table.columns;
        const urnIndex = columns.findIndex(c => c.roles["urn"]);
        const externalIdsIndex = columns.findIndex(c => c.roles["externalIds"]);
        // const guidIndex = columns.findIndex(c => c.roles["guid"]); // Removed
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

        // --- CATEGORICAL COLORING LOGIC ---
        // 1. Identify unique categories from the 'color' column
        // We only do this if a color column is mapped
        if (colorIndex !== -1) {
            // 1. Identify unique categories and find the BEST row to use for settings
            // Estrategia mejorada: buscar TODAS las filas de cada categoría para encontrar colores guardados
            // Map: Category Value -> { rowIndex (primera fila encontrada), savedColor (de cualquier fila de esa categoría) }
            const categoryMetaMap = new Map<string, { rowIndex: number, savedColor: string | null }>();

            // Primera pasada: identificar todas las categorías únicas y encontrar colores guardados
            dataView.table.rows.forEach((row, rowIndex) => {
                const colorValue = String(row[colorIndex]).trim();
                if (!colorValue) return;

                const objects = row.objects;
                let sColor: string | null = null;

                if (objects && objects['dataPoint']) {
                    const dp = objects['dataPoint'] as any;
                    if (dp.fill && dp.fill.solid && dp.fill.solid.color) {
                        sColor = dp.fill.solid.color;
                        console.log(`Visual: Found saved color for category '${colorValue}' at row ${rowIndex}: ${sColor}`);
                    }
                }

                // Si ya tenemos esta categoría, actualizar el color guardado SIEMPRE que encontremos uno
                // (el usuario puede haber cambiado el color, así que debemos usar el más reciente)
                const existing = categoryMetaMap.get(colorValue);
                if (existing) {
                    // Si encontramos un color guardado, usarlo (puede ser nuevo o actualizado)
                    if (sColor) {
                        existing.savedColor = sColor;
                        console.log(`Visual: Updated saved color for '${colorValue}' from row ${rowIndex}: ${sColor}`);
                    }
                    // Si no encontramos color en esta fila pero ya teníamos uno guardado, mantenerlo
                } else {
                    // Primera vez que vemos esta categoría: guardar rowIndex y color (si existe)
                    categoryMetaMap.set(colorValue, {
                        rowIndex: rowIndex,
                        savedColor: sColor
                    });
                }
            });

            // 2. Populate Formatting Settings with Dynamic Slices
            this.formattingSettings.dataPointCard.slices = [];
            this.categoryColorMap.clear();

            // Estrategia de color mejorada:
            // - Colores por defecto consistentes usando hash de la categoría (mismo nombre = mismo color)
            // - Si el usuario elige un color, ese color se guarda y persiste al cambiar de página
            const defaultColors = ['#01B8AA', '#374649', '#FD625E', '#F2C80F', '#5F6B6D', '#8AD4EB', '#FE9666', '#A66999'];
            
            // Función para obtener un color por defecto consistente basado en el nombre de la categoría
            const getDefaultColorForCategory = (category: string): string => {
                // Usar un hash simple del nombre para obtener siempre el mismo color por defecto
                let hash = 0;
                for (let i = 0; i < category.length; i++) {
                    hash = ((hash << 5) - hash) + category.charCodeAt(i);
                    hash = hash & hash; // Convert to 32bit integer
                }
                const index = Math.abs(hash) % defaultColors.length;
                return defaultColors[index];
            };

            categoryMetaMap.forEach((meta, category) => {
                const hasUserColor = !!meta.savedColor;

                // Color final que usaremos tanto en la UI como en el visor:
                // - Si el usuario definió un color, usamos ese (prioridad absoluta).
                // - Si no, usamos un color por defecto CONSISTENTE basado en el nombre de la categoría.
                const finalColor = hasUserColor
                    ? (meta.savedColor as string)
                    : getDefaultColorForCategory(category);

                // DETAILED LOGGING
                console.log(`Visual: Processing category '${category}':`, {
                    rowIndex: meta.rowIndex,
                    savedColor: meta.savedColor,
                    finalColor: finalColor,
                    hasUserColor: hasUserColor
                });

                // Actualizar Map para el visor SIEMPRE (siempre habrá un color final)
                this.categoryColorMap.set(category, finalColor);
                console.log(`Visual: Added '${category}' to categoryColorMap with color ${finalColor}`);

                // Create Selector tied to the specific row
                const selectionIdBuilder = this.host.createSelectionIdBuilder();
                selectionIdBuilder.withTable(dataView.table, meta.rowIndex);
                const selectionId = selectionIdBuilder.createSelectionId();

                // Create Slices
                // ÚNICO control: Color Picker "[Category]" - el usuario puede sobrescribir el color
                const colorSlice = new ColorPicker({
                    name: "fill",
                    displayName: `${category} Color`,
                    value: { value: finalColor },
                    selector: selectionId.getSelector()
                });

                console.log(`Visual: Created slices for '${category}' - Color value: ${finalColor}`);

                this.formattingSettings.dataPointCard.slices.push(colorSlice);
            });
            
            // Detectar cambios en los colores y forzar actualización del visor si es necesario
            let colorsChanged = false;
            if (previousCategoryColorMap.size !== this.categoryColorMap.size) {
                colorsChanged = true;
                console.log(`Visual: Category count changed: ${previousCategoryColorMap.size} -> ${this.categoryColorMap.size}`);
            } else {
                for (const [category, color] of this.categoryColorMap) {
                    const previousColor = previousCategoryColorMap.get(category);
                    if (previousColor !== color) {
                        colorsChanged = true;
                        console.log(`Visual: Color changed for category '${category}': ${previousColor || 'none'} -> ${color}`);
                        break;
                    }
                }
            }
            
            // Si los colores cambiaron y el botón global está activo, actualizar inmediatamente
            if (colorsChanged && this.isGlobalColoringEnabled && this.viewer && this.model && this.idMapping) {
                console.log('Visual: Colors changed, forcing immediate syncColors() update');
                void this.syncColors();
            }

        }


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
        // 5. Handle View GUID - REMOVED
        // let viewGuid: string | null = null;
        // if (guidIndex !== -1 && dataView.table.rows.length > 0) {
        //     viewGuid = String(dataView.table.rows[0][guidIndex]);
        // }

        // 6. Handle Pagination Fetch - AGGRESSIVE
        // We want to fetch ALL data to ensure the viewer can map everything
        const HARD_LIMIT = 200000; // Safety cap to prevent browser crash

        if (this.currentDataView.metadata.segment) {
            if (this.allRows.length < HARD_LIMIT) {
                const moreData = this.host.fetchMoreData();
                if (moreData) {
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
        // Detectar: nuevo URN O reinicio del visor (viewer es null pero hay URN)
        const isNewUrn = modelUrn && modelUrn !== this.currentUrn;
        const needsInitialization = modelUrn && (!this.viewer || isNewUrn);
        
        if (needsInitialization) {
            // Si hay un visor anterior, destruirlo primero
            if (this.viewer) {
                console.log("Visual: Destroying previous viewer before initialization (new URN detected)");
                this.destroyViewer();
            } else if (modelUrn && !this.viewer) {
                console.log("Visual: Viewer restart detected (viewer is null but URN exists)");
            }
            
            this.currentUrn = modelUrn;
            await this.initializeViewer();
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
        } else if (!isDataFilterApplied && !this.isDbIdSelectionActive && !isPaginating) {
            // Only sync colors if we're not in the middle of pagination
            // and there's no active selection from viewer or external filter
            if (this.lastFilteredIds === null) {
                // Initial state - sync colors once
                // CRITICAL: Build reverse mapping (dbId -> column value) for all loaded data
                // This allows us to map selected dbIds back to the values in Power BI column
                await this.buildReverseMapping();

                // Initial sync - only if coloring is enabled (default is OFF)
                // syncColors() will check isGlobalColoringEnabled and do nothing if false
                // Note: setupToolbar() is now called in initializeViewer() to ensure it always loads
                await this.syncColors();
            }
        }

        // --- FINAL STEP: sincronizar colores si el usuario tiene el botón global ACTIVO ---
        // Esto asegura que cualquier cambio en el panel de formato ("Data Colors")
        // se refleje inmediatamente en el visor, incluso si no hay cambios de filtro/selección.
        if (this.isGlobalColoringEnabled) {
            // No esperamos el resultado para no bloquear el pipeline de actualización
            // (syncColors internamente valida viewer/model/idMapping antes de aplicar).
            void this.syncColors();
        }
    }

    private async syncColors(): Promise<void> {
        if (!this.viewer || !this.model || !this.idMapping) return;

        // Respect global toggle
        if (!this.isGlobalColoringEnabled) {
            this.viewer.clearThemingColors(this.model);
            return;
        }

        // Clear any existing theming
        this.viewer.clearThemingColors(this.model);

        // Group External Ids by Hex Color
        const colorToExternalIds = new Map<string, string[]>();

        this.elementDataMap.forEach((entry) => {
            const category = entry.color;
            if (category && this.categoryColorMap.has(category)) {
                const hexColor = this.categoryColorMap.get(category);
                if (!colorToExternalIds.has(hexColor)) {
                    colorToExternalIds.set(hexColor, []);
                }
                colorToExternalIds.get(hexColor).push(entry.id);
            }
        });

        // Apply colors in parallel
        const promises = [];
        for (const [hexColor, externalIds] of colorToExternalIds) {
            promises.push(this.applyColorToExternalIds(hexColor, externalIds));
        }
        await Promise.all(promises);
    }

    private async applyColorToExternalIds(hexColor: string, externalIds: string[]) {
        if (!this.viewer || !this.model || !this.idMapping) return;

        // Use IdMapping to get DbIds efficiently
        const dbIds = await this.idMapping.getDbids(externalIds);

        // Apply Color
        const THREE = (window as any).THREE;
        const colorVector = new THREE.Vector4(
            parseInt(hexColor.substr(1, 2), 16) / 255,
            parseInt(hexColor.substr(3, 2), 16) / 255,
            parseInt(hexColor.substr(5, 2), 16) / 255,
            1 // Alpha
        );

        if (dbIds.length > 0) {
            dbIds.forEach(dbId => {
                this.viewer.setThemingColor(dbId, colorVector, this.model);
            });
        }
    }
    /**
     * Destruye el visor anterior si existe, limpiando todos los recursos
     */
    private destroyViewer(): void {
        // Limpiar animación anterior si existe
        if (this.loadingAnimationHide) {
            this.loadingAnimationHide();
            this.loadingAnimationHide = null;
        }
        if (this.loadingAnimationTimeout) {
            clearTimeout(this.loadingAnimationTimeout);
            this.loadingAnimationTimeout = null;
        }
        this.loadingAnimationStartTime = null;
        hideLoadingAnimation(this.container);

        // Destruir visor si existe
        if (this.viewer) {
            try {
                // Limpiar selección y theming
                if (this.model) {
                    this.viewer.clearSelection();
                    this.viewer.clearThemingColors(this.model);
                }
                // Destruir el visor
                try {
                    // Intentar destruir el visor usando el método disponible
                    if (this.viewer && (this.viewer as any).destroy) {
                        (this.viewer as any).destroy();
                    } else if (this.viewer && (this.viewer as any).impl && (this.viewer as any).impl.terminate) {
                        (this.viewer as any).impl.terminate();
                    }
                } catch (destroyError) {
                    // Si no se puede destruir, simplemente continuar
                    console.warn("Visual: Could not destroy viewer:", destroyError);
                }
            } catch (error) {
                console.warn("Visual: Error destroying viewer:", error);
            }
        }

        // Limpiar contenedor (solo elementos del visor, no el contenedor mismo)
        // El statusDiv está fuera del contenedor, así que no se elimina
        while (this.container.firstChild) {
            this.container.removeChild(this.container.firstChild);
        }

        // Resetear estado
        this.viewer = null;
        this.model = null;
        this.idMapping = null;
        this.isViewerReady = false;
        this.isDbIdSelectionActive = false;
        this.lastFilteredIds = null;
        this.pendingSelection = null;

        // Asegurar que el contenedor tenga las propiedades básicas
        this.container.style.width = '100%';
        this.container.style.height = 'calc(100% - 20px)';
        this.container.style.position = 'relative';
    }

    private async initializeViewer(): Promise<void> {
        try {
            const token = await this.fetchToken(this.currentUrn);
            if (!token) {
                return;
            }

            // Get performance profile from settings

            // Get environment from settings
            // Get performance profile from settings
            const profile = this.formattingSettings.viewerCard.performanceProfile.value.value as 'HighPerformance' | 'Balanced';

            // Get environment from settings
            const env = this.formattingSettings.viewerCard.viewerEnv.value;
            // Determine API based on Env (SVF2 -> streamingV2, SVF -> derivativeV2)
            const api = env === 'AutodeskProduction2' ? 'streamingV2' : 'derivativeV2';
            
            // Get skipPropertyDb setting for faster loading
            const skipPropertyDb = this.formattingSettings.viewerCard.skipPropertyDb.value;

            // Use the new launchViewer from utils
            // Callback se ejecuta DESPUÉS de viewer.start() para mostrar animación y tapar "Powered By Autodesk"
            const { viewer, model } = await launchViewer(
                this.container, // Use container instead of targetDiv if changed, or ensure definition
                this.currentUrn,
                token.access_token,
                null, // Explicitly pass null for GUID as it was removed
                () => {
                    // Debounce logic:
                    // When selection changes, wait 150ms (reduced for better fluidity).
                    // If another change happens, reset timer.
                    if (this.selectionDebounceTimer) {
                        clearTimeout(this.selectionDebounceTimer);
                    }
                    this.selectionDebounceTimer = setTimeout(() => {
                        this.handleSelectionChange();
                        this.selectionDebounceTimer = null;
                    }, 150);
                },
                profile, // Pass performance profile
                env, // Pass environment
                api, // Pass API
                () => {
                    // IMPORTANTE: Mostrar animación DESPUÉS de viewer.start()
                    // Esto tapa el mensaje "Powered By Autodesk" que aparece durante la carga
                    console.log("Visual: Showing loading animation after viewer.start() to hide 'Powered By Autodesk' message");
                    
                    // Guardar tiempo de inicio
                    this.loadingAnimationStartTime = Date.now();
                    
                    const hideFn = showLoadingAnimation(
                        this.container,
                        this.LOADING_ANIMATION_MAX_DURATION, // Usar duración máxima como timeout de seguridad
                        "Preparando tu modelo...",
                        "SKYDATABIM S.A.S."
                    );
                    this.loadingAnimationHide = hideFn;
                    
                    // Timeout de seguridad: ocultar después del máximo si no se ha ocultado antes
                    this.loadingAnimationTimeout = window.setTimeout(() => {
                        console.log("Visual: Maximum animation duration reached. Forcing hide.");
                        if (this.loadingAnimationHide) {
                            this.loadingAnimationHide();
                            this.loadingAnimationHide = null;
                        }
                        this.loadingAnimationTimeout = null;
                        this.loadingAnimationStartTime = null;
                        hideLoadingAnimation(this.container);
                    }, this.LOADING_ANIMATION_MAX_DURATION);
                },
                skipPropertyDb // Pass skipPropertyDb for faster loading
            );

            this.viewer = viewer;
            this.model = model;
            this.idMapping = new IdMapping(this.model);
            this.isViewerReady = true;

            console.log("Visual: Viewer initialized successfully with profile:", profile);

            // CRITICAL: Set up toolbar IMMEDIATELY after viewer is ready
            // This ensures the button is always created, regardless of model size or other conditions
            // Use a small delay to ensure toolbar is fully initialized
            setTimeout(() => {
                this.setupToolbar();
            }, 100);
            
            // IMPORTANTE: Ocultar animación cuando el modelo esté completamente cargado
            // Esto asegura que siempre tape "Powered By Autodesk" hasta que el modelo esté listo
            const hideAnimationWhenReady = () => {
                if (!this.loadingAnimationStartTime) {
                    // No hay animación activa, salir
                    return;
                }
                
                // Calcular tiempo transcurrido desde que se inició la animación
                const elapsed = Date.now() - this.loadingAnimationStartTime;
                const remaining = this.LOADING_ANIMATION_MIN_DURATION - elapsed;
                
                if (elapsed >= this.LOADING_ANIMATION_MIN_DURATION) {
                    // Han pasado al menos 6 segundos, ocultar animación inmediatamente
                    console.log(`Visual: Model loaded. Hiding loading animation (${elapsed}ms elapsed, minimum 6s satisfied)`);
                    if (this.loadingAnimationHide) {
                        this.loadingAnimationHide();
                        this.loadingAnimationHide = null;
                    }
                    if (this.loadingAnimationTimeout) {
                        clearTimeout(this.loadingAnimationTimeout);
                        this.loadingAnimationTimeout = null;
                    }
                    this.loadingAnimationStartTime = null;
                    hideLoadingAnimation(this.container);
                } else {
                    // Aún no han pasado 6 segundos, esperar el tiempo restante
                    console.log(`Visual: Model loaded but animation must continue for ${remaining}ms more to reach minimum 6s duration`);
                    setTimeout(() => {
                        if (this.loadingAnimationHide) {
                            this.loadingAnimationHide();
                            this.loadingAnimationHide = null;
                        }
                        if (this.loadingAnimationTimeout) {
                            clearTimeout(this.loadingAnimationTimeout);
                            this.loadingAnimationTimeout = null;
                        }
                        this.loadingAnimationStartTime = null;
                        hideLoadingAnimation(this.container);
                        console.log("Visual: Minimum duration satisfied. Animation hidden.");
                    }, remaining);
                }
            };

            // Escuchar eventos de carga del modelo
            // GEOMETRY_LOADED_EVENT se dispara cuando la geometría está lista
            viewer.addEventListener(Autodesk.Viewing.GEOMETRY_LOADED_EVENT, () => {
                console.log("Visual: GEOMETRY_LOADED_EVENT fired - geometry is loaded");
                hideAnimationWhenReady();
            });

            // También escuchar OBJECT_TREE_CREATED_EVENT como respaldo
            viewer.addEventListener(Autodesk.Viewing.OBJECT_TREE_CREATED_EVENT, () => {
                console.log("Visual: OBJECT_TREE_CREATED_EVENT fired - object tree created");
                // Solo ocultar si ya pasó el tiempo mínimo
                if (this.loadingAnimationStartTime) {
                    const elapsed = Date.now() - this.loadingAnimationStartTime;
                    if (elapsed >= this.LOADING_ANIMATION_MIN_DURATION) {
                        hideAnimationWhenReady();
                    }
                }
            });

            // Verificar periódicamente si el modelo está cargado (por si los eventos no se disparan)
            const checkModelLoaded = setInterval(() => {
                if (model && model.isLoadDone && model.isLoadDone()) {
                    console.log("Visual: Model load done detected via isLoadDone()");
                    clearInterval(checkModelLoaded);
                    hideAnimationWhenReady();
                }
            }, 500);

            // Limpiar el intervalo después del tiempo máximo
            setTimeout(() => {
                clearInterval(checkModelLoaded);
            }, this.LOADING_ANIMATION_MAX_DURATION);

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
            
            // Ocultar animación en caso de error
            if (this.loadingAnimationHide) {
                this.loadingAnimationHide();
                this.loadingAnimationHide = null;
            }
            if (this.loadingAnimationTimeout) {
                clearTimeout(this.loadingAnimationTimeout);
                this.loadingAnimationTimeout = null;
            }
            this.loadingAnimationStartTime = null;
            hideLoadingAnimation(this.container);
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

    // Flag to track if we need to run another update after the current one finishes
    private isUpdatePending: boolean = false;
    private isGlobalColoringEnabled: boolean = false; // Default to OFF (user must activate)
    private paintBucketButton: Autodesk.Viewing.UI.Button | null = null; // Store reference

    // eslint-disable-next-line max-lines-per-function
    private async handleSelectionChange() {
        if (!this.host || !this.viewer || !this.idMapping) return;

        // Skip if this selection change came from programmatic selection (external filter)
        if (this.isProgrammaticSelection) {
            return;
        }

        // Concurrency Management:
        // If already processing, mark that we need another update (because selection has changed again)
        // and return. The running process will pick this up when it finishes.
        if (this.isProcessingSelection) {
            console.log("Visual: Selection change already running, queuing next update...");
            this.isUpdatePending = true;
            return;
        }

        this.isProcessingSelection = true;
        this.isUpdatePending = false; // Clear pending flag as we are starting now

        try {
            // Start the processing loop
            // We loop as long as isUpdatePending becomes true during our execution
            do {
                // If we are looping back, it means a new selection happened while we were working.
                // We clear the flag and process the *latest* state.
                if (this.isUpdatePending) {
                    console.log("Visual: Processing queued selection update...");
                    this.isUpdatePending = false;
                }

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
                        // Use empty array to clear selection, BUT false to NOT ensure absolute?
                        // Actually, select([]) clears it.
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


                // 5. Apply filter directly with ALL selected External IDs
                // We trust the Viewer: if an object is selected, we want to filter by it.
                // Power BI will handle the matching against its data model.
                // We do NOT validate against this.externalIds because that array might be incomplete
                // (pagination in progress or limits reached).
                const filterValues = selectedExternalIds;

                // 6. Apply filter if we have valid values
                if (filterValues.length > 0) {
                    if (filterValues.length > 1) {
                        console.log(`Visual: Applying multi-selection filter with ${filterValues.length} External IDs`);
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
                            this.elementDataMap.forEach((value) => {
                                if (value.id === extId) {
                                    selectionIds.push(value.selectionId);
                                }
                            });
                        }

                        if (selectionIds.length > 0) {
                            // Apply selection to maintain state consistency
                            this.selectionManager.select(selectionIds, false); // false = replace (sync with absolute viewer state)
                        }
                    } catch (selectionError) {
                        console.warn("Visual: Could not apply selection via SelectionManager:", selectionError);
                    }
                } else {
                    console.warn("Visual: No valid External IDs found in dataset. Cannot apply filter.");
                }

            } while (this.isUpdatePending);

        } catch (error) {
            console.error("Visual: Error handling selection change", error);
        } finally {
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
        if (this.selectionDebounceTimer) {
            clearTimeout(this.selectionDebounceTimer);
        }

        // Set new timer
        this.selectionDebounceTimer = setTimeout(() => {
            this.handleSelectionChange();
            this.selectionDebounceTimer = null;
        }, this.SELECTION_DEBOUNCE_MS);
    }

    private setupToolbar() {
        if (!this.viewer) {
            console.warn('Visual: setupToolbar called but viewer is not available');
            return;
        }

        // Helper to add button
        const createButton = () => {
            if (this.paintBucketButton) {
                console.log('Visual: Paint bucket button already exists, skipping creation');
                return; // Already created
            }

            const toolbar = (this.viewer as any).getToolbar();
            if (!toolbar) {
                console.warn('Visual: Toolbar not available yet, will retry...');
                // Retry after a short delay if toolbar is not ready
                setTimeout(() => {
                    if (!this.paintBucketButton) {
                        this.setupToolbar();
                    }
                }, 500);
                return;
            }

            // Create Button
            this.paintBucketButton = new Autodesk.Viewing.UI.Button('paint-bucket-btn');
            this.paintBucketButton.setToolTip('Toggle Colors');
            this.paintBucketButton.addClass('paint-bucket-icon');

            // Color palette icon SVG (similar design to ghosting button)
            // Designed to be clear and visible at 28x28px - represents a color palette
            const paintBucketIcon = `
                <svg width="28" height="28" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                    <!-- Palette base (oval shape) -->
                    <ellipse cx="12" cy="16" rx="8" ry="5" 
                             fill="currentColor" opacity="0.9"/>
                    
                    <!-- Palette thumb hole (top center) -->
                    <ellipse cx="12" cy="12" rx="2.5" ry="3" 
                             fill="white" opacity="0.3"/>
                    <ellipse cx="12" cy="12" rx="1.8" ry="2.2" 
                             fill="currentColor" opacity="0.6"/>
                    
                    <!-- Color spots on palette (representing different colors) -->
                    <!-- Red spot -->
                    <circle cx="8" cy="14" r="1.8" fill="#FF4444" opacity="0.9"/>
                    <circle cx="8" cy="14" r="1.2" fill="#FF6666"/>
                    
                    <!-- Green spot -->
                    <circle cx="16" cy="14" r="1.8" fill="#44FF44" opacity="0.9"/>
                    <circle cx="16" cy="14" r="1.2" fill="#66FF66"/>
                    
                    <!-- Blue spot -->
                    <circle cx="10" cy="18" r="1.8" fill="#4444FF" opacity="0.9"/>
                    <circle cx="10" cy="18" r="1.2" fill="#6666FF"/>
                    
                    <!-- Yellow spot -->
                    <circle cx="14" cy="18" r="1.8" fill="#FFDD44" opacity="0.9"/>
                    <circle cx="14" cy="18" r="1.2" fill="#FFEE66"/>
                    
                    <!-- Brush handle (extending from palette) -->
                    <path d="M 12 11 L 12 6 L 13.5 6 L 13.5 11" 
                          fill="currentColor" opacity="0.7"/>
                    <path d="M 12 6 L 13.5 6 L 13.2 4.5 L 12.3 4.5 Z" 
                          fill="currentColor" opacity="0.8"/>
                    
                    <!-- Brush tip -->
                    <ellipse cx="12.75" cy="5" rx="0.8" ry="1.2" 
                             fill="currentColor" opacity="0.6"/>
                    
                    <!-- Highlight on palette -->
                    <ellipse cx="10" cy="14" rx="3" ry="2" 
                             fill="white" opacity="0.15"/>
                </svg>
            `;

            // Set the icon directly (same pattern as ghosting button)
            if (this.paintBucketButton.icon) {
                // eslint-disable-next-line powerbi-visuals/no-inner-outer-html
                this.paintBucketButton.icon.innerHTML = paintBucketIcon;
                // Ensure the icon is visible and properly styled
                this.paintBucketButton.icon.style.display = 'flex';
                this.paintBucketButton.icon.style.alignItems = 'center';
                this.paintBucketButton.icon.style.justifyContent = 'center';
                this.paintBucketButton.icon.style.width = '28px';
                this.paintBucketButton.icon.style.height = '28px';
            }

            // Set initial state based on isGlobalColoringEnabled
            const initialState = this.isGlobalColoringEnabled ? 
                Autodesk.Viewing.UI.Button.State.ACTIVE : 
                Autodesk.Viewing.UI.Button.State.INACTIVE;
            this.paintBucketButton.setState(initialState);

            // On Click
            this.paintBucketButton.onClick = () => {
                this.isGlobalColoringEnabled = !this.isGlobalColoringEnabled;

                // Update button state (same pattern as ghosting button)
                const newState = this.isGlobalColoringEnabled ? 
                    Autodesk.Viewing.UI.Button.State.ACTIVE : 
                    Autodesk.Viewing.UI.Button.State.INACTIVE;
                this.paintBucketButton.setState(newState);

                // Update visual state
                if (this.isGlobalColoringEnabled) {
                    this.syncColors();
                } else {
                    this.viewer.clearThemingColors(this.model);
                }
                
                console.log('Visual: Paint bucket (colors) set to ' + (this.isGlobalColoringEnabled ? 'ACTIVE' : 'INACTIVE'));
            };

            // Add to toolbar (Model Tools sub-toolbar usually)
            // Find 'modelTools' or create new group
            let subToolbar = toolbar.getControl('modelTools');
            if (!subToolbar) {
                subToolbar = new Autodesk.Viewing.UI.ControlGroup('my-custom-tools');
                toolbar.addControl(subToolbar);
            }
            subToolbar.addControl(this.paintBucketButton);
        };

        // Try to create button immediately
        if ((this.viewer as any).getToolbar()) {
            try {
                createButton();
                console.log('Visual: Paint bucket button created successfully');
            } catch (error) {
                console.error('Visual: Error creating paint bucket button:', error);
                // Retry after delay
                setTimeout(() => {
                    if (!this.paintBucketButton) {
                        this.setupToolbar();
                    }
                }, 1000);
            }
        } else {
            // Toolbar not ready yet, wait for it
            console.log('Visual: Waiting for toolbar to be created...');
            const toolbarListener = () => {
                try {
                    createButton();
                    console.log('Visual: Paint bucket button created after toolbar event');
                    this.viewer.removeEventListener(Autodesk.Viewing.TOOLBAR_CREATED_EVENT, toolbarListener);
                } catch (error) {
                    console.error('Visual: Error creating paint bucket button after toolbar event:', error);
                }
            };
            this.viewer.addEventListener(Autodesk.Viewing.TOOLBAR_CREATED_EVENT, toolbarListener);
            
            // Also set a timeout as backup in case the event doesn't fire
            setTimeout(() => {
                if (!this.paintBucketButton && (this.viewer as any).getToolbar()) {
                    console.log('Visual: Creating paint bucket button via timeout backup');
                    try {
                        createButton();
                        this.viewer.removeEventListener(Autodesk.Viewing.TOOLBAR_CREATED_EVENT, toolbarListener);
                    } catch (error) {
                        console.error('Visual: Error in timeout backup for paint bucket button:', error);
                    }
                }
            }, 2000);
        }
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

    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    private async fetchToken(_urn: string): Promise<any> {
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

    // eslint-disable-next-line max-lines-per-function
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

                    // OPTIMIZATION: Apply theming in batch for better performance
                    // Use setThemingColor with array for faster batch operation if available
                    // Otherwise, apply individually but optimize with requestAnimationFrame
                    if (finalDbIds.length > 100) {
                        // For large selections, batch the theming operation
                        // Apply theming in chunks to avoid blocking the UI
                        const chunkSize = 100;
                        for (let i = 0; i < finalDbIds.length; i += chunkSize) {
                            const chunk = finalDbIds.slice(i, i + chunkSize);
                            chunk.forEach(dbid => {
                                this.viewer.setThemingColor(dbid, NUCLEAR_GREEN, this.model);
                            });
                            // Yield to browser for smoother rendering
                            if (i + chunkSize < finalDbIds.length) {
                                await new Promise(resolve => requestAnimationFrame(resolve));
                            }
                        }
                    } else {
                        // For smaller selections, apply all at once
                        finalDbIds.forEach(dbid => {
                            this.viewer.setThemingColor(dbid, NUCLEAR_GREEN, this.model);
                        });
                    }

                    // OPTIMIZATION: Use requestAnimationFrame to synchronize visual updates
                    // This ensures smoother rendering and better bidirectional interaction speed
                    await new Promise(resolve => requestAnimationFrame(resolve));

                    // Step 6: Fit camera to view the isolated elements
                    console.log(`Visual: syncSelectionState - Fitting camera to ${finalDbIds.length} isolated elements`);
                    fitToView(this.viewer, this.model, finalDbIds);
                    
                    // OPTIMIZATION: Only invalidate if necessary (avoid unnecessary repaints)
                    // The viewer will automatically repaint after isolation and theming
                    // Only force invalidate for very large models if needed
                    if (finalDbIds.length > 1000 && (this.viewer as any).impl?.invalidate) {
                        console.log(`Visual: syncSelectionState - Forcing viewer.invalidate() for large selection (${finalDbIds.length} elements)`);
                        (this.viewer as any).impl.invalidate(true, true, true);
                    }

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


    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}

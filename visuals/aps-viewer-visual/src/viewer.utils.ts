/// import * as Autodesk from "@types/forge-viewer";

'use strict';

import { registerGhostingExtension } from "./ghosting.extension"; // Import to register extension

const runtime: { options: Autodesk.Viewing.InitializerOptions; ready: Promise<void> | null } = {
    options: {},
    ready: null
};

declare global {
    interface Window { DISABLE_INDEXED_DB: boolean; }
}

/**
 * Initializes the Viewer Runtime. Uses a Singleton pattern to avoid re-initialization.
 * @param options Initializer options
 */
export function initializeViewerRuntime(options: Autodesk.Viewing.InitializerOptions): Promise<void> {
    if (!runtime.ready) {
        runtime.options = { ...options };
        runtime.ready = (async function () {
            // Enable OPFS if supported by browser (Viewer 7.98+ does this by default, but good to ensure)
            // window.DISABLE_INDEXED_DB = true; // Removed to allow OPFS/IndexedDB cache

            await loadScript('https://developer.api.autodesk.com/modelderivative/v2/viewers/7.*/viewer3D.js');
            await loadStylesheet('https://developer.api.autodesk.com/modelderivative/v2/viewers/7.*/style.css');

            // Register extensions now that Autodesk namespace is available
            registerGhostingExtension();

            return new Promise((resolve) => Autodesk.Viewing.Initializer(runtime.options, resolve));
        })() as Promise<void>;
    } else {
        // Check if critical options changed (though we can't really re-init)
        if (['accessToken', 'getAccessToken', 'env', 'api', 'language'].some(prop => options[prop] !== runtime.options[prop])) {
            console.warn('Visual: Cannot initialize another viewer runtime with different settings. Using existing runtime.');
            // return Promise.reject('Cannot initialize another viewer runtime with different settings.');
        }
    }
    return runtime.ready;
}

export function loadModel(viewer: Autodesk.Viewing.Viewer3D, urn: string, guid?: string, skipPropertyDb: boolean = false): Promise<Autodesk.Viewing.Model> {
    return new Promise(function (resolve, reject) {
        const loadOptions: any = {
            skipPropertyDb: skipPropertyDb
        };

        Autodesk.Viewing.Document.load(
            'urn:' + urn,
            (doc) => {
                const view = guid ? doc.getRoot().findByGuid(guid) : doc.getRoot().getDefaultGeometry();
                viewer.loadDocumentNode(doc, view, loadOptions).then(m => resolve(m));
            },
            (code, message, args) => reject({ code, message, args })
        );
    });
}

export function getVisibleNodes(model: Autodesk.Viewing.Model): number[] {
    const tree = model.getInstanceTree();
    const dbids: number[] = [];
    tree.enumNodeChildren(tree.getRootId(), dbid => {
        if (tree.getChildCount(dbid) === 0 && !tree.isNodeHidden(dbid) && !tree.isNodeOff(dbid)) {
            dbids.push(dbid);
        }
    }, true);
    return dbids;
}

/**
 * Isolates specific elements in the viewer by their dbIds.
 * @param viewer The Autodesk Viewer instance.
 * @param model The model to apply isolation to.
 * @param dbIds Array of dbIds to isolate. If empty or undefined, shows all elements.
 */
export function isolateDbIds(viewer: Autodesk.Viewing.Viewer3D, model: Autodesk.Viewing.Model, dbIds: number[]): void {
    if (!viewer || !model) return;
    viewer.isolate(dbIds, model);
}

/**
 * Fits the view to show specific elements by their dbIds.
 * @param viewer The Autodesk Viewer instance.
 * @param model The model to apply fit to.
 * @param dbIds Array of dbIds to fit to. If empty or undefined, fits to entire model.
 */
export function fitToView(viewer: Autodesk.Viewing.Viewer3D, model: Autodesk.Viewing.Model, dbIds?: number[]): void {
    if (!viewer || !model) return;
    viewer.fitToView(dbIds, model);
}

/**
 * Shows all elements in the viewer (clears isolation).
 * @param viewer The Autodesk Viewer instance.
 * @param model The model to apply show all to.
 */
export function showAll(viewer: Autodesk.Viewing.Viewer3D, model: Autodesk.Viewing.Model): void {
    if (!viewer || !model) return;
    viewer.isolate(undefined, model);
    viewer.fitToView(undefined, model);
}

/**
 * Helper class for mapping between "dbIds" (sequential numbers assigned to each design element,
 * typically used by the Viewer APIs) and "externalId" values (persistent IDs from the authoring
 * application, for example Revit GUIDs).
 *
 * IMPORTANT (externalIdOnly mode):
 * --------------------------------
 * The public contract of this class is **ExternalId-first**:
 *  - Power BI and the custom visual work EXCLUSIVELY with ExternalId strings.
 *  - dbIds are used only internally when calling Viewer APIs (isolate, fitToView, theming, etc.).
 *
 * External consumers (for example `visual.ts`) must:
 *  - Use ExternalId as the primary identifier when talking to Power BI.
 *  - Use `getDbids(externalIds)` ONLY when they need to call Viewer APIs that require dbIds.
 *  - Never depend on dbId values as part of any filter that goes back to Power BI.
 */
export class IdMapping {
    /** When true, the system assumes ExternalId is the primary identifier exposed to the outside world. */
    private readonly externalIdOnly: boolean = true;

    private readonly externalIdMappingPromise: Promise<{ [externalId: string]: number; }>;
    private readonly reverseMappingPromise: Promise<{ [dbid: number]: string; }>;

    constructor(private model: Autodesk.Viewing.Model) {
        console.log("Visual: IdMapping - Initializing in externalIdOnly mode");
        this.externalIdMappingPromise = new Promise((resolve, reject) => {
            model.getExternalIdMapping(resolve, reject);
        });

        // Create reverse mapping for faster lookup
        this.reverseMappingPromise = this.externalIdMappingPromise.then(externalIdMapping => {
            const reverse: { [dbid: number]: string; } = {};
            for (const [externalId, dbid] of Object.entries(externalIdMapping)) {
                if (dbid != null) {
                    reverse[dbid] = externalId;
                }
            }
            return reverse;
        });
    }

    /**
     * Resolves an array of ExternalIds against the model's ExternalId mapping.
     * Returns which ExternalIds were found and which were missing.
     */
    async resolveExternalIds(externalIds: string[]): Promise<{ found: string[]; missing: string[] }> {
        const mapping = await this.externalIdMappingPromise;

        const found: string[] = [];
        const missing: string[] = [];

        externalIds.forEach(rawId => {
            const externalId = (rawId ?? "").toString().trim();
            if (!externalId) {
                return;
            }
            if (Object.prototype.hasOwnProperty.call(mapping, externalId)) {
                found.push(externalId);
            } else {
                missing.push(externalId);
            }
        });

        if (missing.length > 0) {
            console.warn(`Visual: IdMapping.resolveExternalIds - ${missing.length} ExternalIds not found in model. Sample:`, missing.slice(0, 20));
        }

        return { found, missing };
    }

    /**
     * Returns dbIds for a set of ExternalIds. This is meant for INTERNAL use only when
     * calling Viewer APIs that require dbIds (isolation, fitToView, theming, etc.).
     */
    /**
     * Returns dbIds for a set of ExternalIds. This is meant for INTERNAL use only when
     * calling Viewer APIs that require dbIds (isolation, fitToView, theming, etc.).
     */
    async getDbids(externalIds: string[]): Promise<number[]> {
        const mapping = await this.externalIdMappingPromise;

        const dbIds: number[] = [];
        const missing: string[] = [];

        externalIds.forEach(rawId => {
            const externalId = (rawId ?? "").toString().trim();
            if (!externalId) {
                return;
            }
            const dbId = mapping[externalId];
            if (dbId != null && !isNaN(dbId)) {
                dbIds.push(dbId);
            } else {
                missing.push(externalId);
            }
        });

        if (missing.length > 0) {
            // Only warn if significant number missing to avoid log spam
            if (missing.length > 20) {
                console.warn(`Visual: IdMapping.getDbids - ${missing.length} ExternalIds have no dbId in model. Sample:`, missing.slice(0, 10));
            }
        }

        return dbIds;
    }

    /**
     * Returns ExternalIds for a set of dbIds.
     *
     * This is primarily used when the Viewer reports selections as dbIds and we need to
     * convert them back to ExternalIds so that Power BI can be filtered using ExternalId strings.
     *
     * Behaviour:
     *  - First tries to resolve dbId → ExternalId using the precomputed reverse mapping.
     *  - If a direct mapping is not found, it walks up the instance tree to find the nearest
     *    ancestor that has an ExternalId.
     *  - If no ancestor has an ExternalId, the entry will be returned as undefined.
     *
     * The returned array is aligned with the input array (same length, same order).
     */
    async getExternalIds(dbids: number[]): Promise<string[]> {
        const reverse = await this.reverseMappingPromise;
        const tree = this.model.getInstanceTree();

        const result: string[] = [];
        const unresolved: number[] = [];

        const resolveSingle = (dbid: number): string | undefined => {
            if (dbid == null || isNaN(dbid)) {
                return undefined;
            }

            // 1) Direct reverse mapping
            if (reverse[dbid]) {
                return reverse[dbid];
            }

            // 2) Walk up the instance tree to find an ancestor that has an ExternalId
            let current: number | undefined = dbid;
            try {
                // Safety check for infinite loops
                let depth = 0;
                while (current != null && !isNaN(current) && depth < 100) {
                    const parent = tree.getNodeParentId(current);
                    if (parent == null || isNaN(parent) || parent === current) {
                        break;
                    }
                    if (reverse[parent]) {
                        return reverse[parent];
                    }
                    current = parent;
                    depth++;
                }
            } catch (e) {
                // If anything goes wrong while walking the tree, fall back to undefined
            }

            return undefined;
        };

        for (const dbid of dbids) {
            const extId = resolveSingle(dbid);
            if (extId) {
                result.push(extId);
            } else {
                unresolved.push(dbid);
            }
        }

        if (unresolved.length > 0) {
            // Only warn if significant number missing
            if (unresolved.length > 20) {
                console.warn(`Visual: IdMapping.getExternalIds - ${unresolved.length} dbIds have no ExternalId mapping. Sample:`, unresolved.slice(0, 10));
            }
        }

        return result;
    }

    /**
     * Validates if DbIds exist in the model
     */
    validateDbIds(dbids: number[]): Promise<number[]> {
        return new Promise((resolve) => {
            const tree = this.model.getInstanceTree();
            const validDbIds: number[] = [];

            for (const dbid of dbids) {
                if (dbid != null && !isNaN(dbid)) {
                    try {
                        // Check if node exists by trying to get its type
                        const nodeType = tree.getNodeType(dbid);
                        if (nodeType != null) {
                            validDbIds.push(dbid);
                        }
                    } catch (e) {
                        // Node doesn't exist
                    }
                }
            }

            resolve(validDbIds);
        });
    }

    /**
     * DEPRECATED: smartMapToDbIds
     *
     * Historically this method tried to be smart by accepting either dbIds or ExternalIds and
     * figuring out what they were. In the new ExternalId-only design this behaviour is confusing
     * and can lead to a large number of "Invalid/Unmapped IDs" logs.
     *
     * It is kept for backwards compatibility, but:
     *  - It now simply assumes that all inputs are ExternalIds.
     *  - It delegates to getDbids(externalIds).
     *
     * New code should prefer:
     *  - getDbids(externalIds)            // when you need dbIds for Viewer APIs
     *  - resolveExternalIds(externalIds)  // when you only need to validate ExternalIds
     */
    async smartMapToDbIds(ids: string[]): Promise<number[]> {
        console.warn("Visual: IdMapping.smartMapToDbIds is deprecated. Prefer getDbids(...) with ExternalIds.");
        return this.getDbids(ids);
    }
}

/**
 * Helper to get the current viewer selection as a list of ExternalIds.
 * This abstracts the dbId -> ExternalId conversion from the main visual logic.
 */
export async function getSelectionAsExternalIds(viewer: Autodesk.Viewing.Viewer3D, mapping: IdMapping): Promise<string[]> {
    if (!viewer || !mapping) return [];

    // Get ALL selected elements from the viewer
    // When using Ctrl+Click, the viewer maintains all selected elements in its selection array
    const selection = viewer.getSelection();
    if (!selection || selection.length === 0) return [];

    // Convert all dbIds to ExternalIds
    // This should include all elements selected with Ctrl+Click
    const externalIds = await mapping.getExternalIds(selection);
    
    return externalIds;
}

/**
 * Convenience helper: isolate elements by ExternalId.
 * Uses IdMapping.getDbids under the hood and then calls isolateDbIds.
 */
export async function isolateExternalIds(
    viewer: Autodesk.Viewing.Viewer3D,
    model: Autodesk.Viewing.Model,
    mapping: IdMapping,
    externalIds: string[]
): Promise<void> {
    if (!viewer || !model || !mapping) return;
    const dbIds = await mapping.getDbids(externalIds);
    if (!dbIds || dbIds.length === 0) {
        console.warn("Visual: isolateExternalIds - No dbIds found for given ExternalIds. Showing all.");
        showAll(viewer, model);
        return;
    }
    isolateDbIds(viewer, model, dbIds);
}

/**
 * Convenience helper: fit camera to elements by ExternalId.
 * Uses IdMapping.getDbids under the hood and then calls fitToView.
 */
export async function fitToExternalIds(
    viewer: Autodesk.Viewing.Viewer3D,
    model: Autodesk.Viewing.Model,
    mapping: IdMapping,
    externalIds?: string[]
): Promise<void> {
    if (!viewer || !model || !mapping) return;
    if (!externalIds || externalIds.length === 0) {
        fitToView(viewer, model, undefined);
        return;
    }
    const dbIds = await mapping.getDbids(externalIds);
    if (!dbIds || dbIds.length === 0) {
        console.warn("Visual: fitToExternalIds - No dbIds found for given ExternalIds. Fitting to whole model.");
        fitToView(viewer, model, undefined);
        return;
    }
    fitToView(viewer, model, dbIds);
}

function loadScript(src: string): Promise<void> {
    return new Promise((resolve, reject) => {
        const el = document.createElement("script");
        el.onload = () => resolve();
        el.onerror = (err) => reject(err);
        el.type = 'application/javascript';
        el.src = src;
        document.head.appendChild(el);
    });
}

function loadStylesheet(href: string): Promise<void> {
    return new Promise((resolve, reject) => {
        const el = document.createElement('link');
        el.onload = () => resolve();
        el.onerror = (err) => reject(err);
        el.rel = 'stylesheet';
        el.href = href;
        document.head.appendChild(el);
    });
}

// --- NEW FUNCTIONS FOR BIMSualize PATTERN & PERFORMANCE ---

export async function launchViewer(
    container: HTMLElement,
    urn: string,
    token: string,
    guid: string | null,
    onDbIdsChanged: (ids: number[]) => void,
    performanceProfile: 'HighPerformance' | 'Balanced' = 'HighPerformance',
    env: string = 'AutodeskProduction2',
    api: string = 'streamingV2',
    onViewerStarted?: () => void,
    skipPropertyDb: boolean = false
): Promise<{ viewer: Autodesk.Viewing.Viewer3D, model: Autodesk.Viewing.Model }> {

    // Initialize runtime if not already (Singleton)
    // Use streamingV2 for SVF2 support
    await initializeViewerRuntime({
        env: env,
        api: api,
        getAccessToken: (callback) => callback(token, 3600)
    });

    // Create Viewer instance
    // We use GuiViewer3D to get the toolbar
    // Optimize: Disable extensions that are not needed for better performance
    const config: any = {
        extensions: ['BIMetric.GhostingExtension'], // Load our ghosting extension
        disabledExtensions: {
            measure: false, // Keep measure for user interaction
            viewcube: false, // Keep viewcube for navigation
            explode: true, // Disable explode (not needed)
            section: false, // Keep section for analysis
            bimwalk: false, // Enable first-person navigation
            fusionOrbit: true, // Disable fusionOrbit (not needed for BIM)
            modelBrowser: false, // Keep modelBrowser (useful)
            propertiesPanel: false, // Keep propertiesPanel (useful)
            layerManager: true, // Disable layerManager for performance (can be enabled if needed)
            hyperlink: true, // Disable hyperlink (not needed)
            // Disable MixpanelExtension to prevent localStorage errors in Power BI sandbox
            MixpanelExtension: true,
            // Disable other heavy extensions for better performance
            Debug: true, // Disable debug extension
            MarkupsCore: true, // Disable markups if not needed
            MarkupsGui: true // Disable markups GUI if not needed
        }
    };

    const viewer = new Autodesk.Viewing.GuiViewer3D(container, config as Autodesk.Viewing.ViewerConfig);
    
    // Suppress MixpanelExtension errors by preventing its registration
    // This extension tries to use localStorage which is not available in Power BI sandbox
    const originalRegisterExtension = Autodesk.Viewing.theExtensionManager.registerExtension;
    Autodesk.Viewing.theExtensionManager.registerExtension = function(name: string, extension: any) {
        if (name === 'Autodesk.Viewing.MixpanelExtension') {
            console.warn('Visual: Suppressing MixpanelExtension registration to prevent localStorage errors');
            return; // Don't register this extension
        }
        return originalRegisterExtension.call(this, name, extension);
    };
    
    viewer.start();
    
    // Restore original function after viewer starts
    setTimeout(() => {
        Autodesk.Viewing.theExtensionManager.registerExtension = originalRegisterExtension;
    }, 1000);

    // IMPORTANTE: Ejecutar callback DESPUÉS de viewer.start()
    // Esto permite mostrar la animación para tapar el mensaje "Powered By Autodesk"
    if (onViewerStarted) {
        onViewerStarted();
    }

    // Load VisualClusters extension for better performance with large models
    // Wrap in try-catch to prevent errors from blocking viewer initialization
    try {
        await viewer.loadExtension('Autodesk.VisualClusters');
    } catch (error) {
        console.warn('Visual: Could not load VisualClusters extension:', error);
        // Continue without the extension - it's optional for performance
    }

    // Load Model
    // skipPropertyDb: true for faster loading (skip property database)
    // Note: If skipPropertyDb is true, properties won't be available but loading is much faster
    // ExternalId mapping (needed for bidirectional interaction) is NOT affected by skipPropertyDb
    const model = await loadModel(viewer, urn, guid || undefined, skipPropertyDb);

    // Apply Performance Profile (this may set lighting/background for performance)
    applyPerformanceProfile(viewer, performanceProfile);

    // Apply default viewer configuration (pass the container for proper positioning)
    // IMPORTANT: This is called AFTER applyPerformanceProfile to ensure natural environment (preset 2) is always active
    applyDefaultViewerConfiguration(viewer, container);

    // Setup Selection Listener
    viewer.addEventListener(Autodesk.Viewing.SELECTION_CHANGED_EVENT, (ev: any) => {
        const ids = ev?.dbIdArray ?? [];
        onDbIdsChanged(ids);
    });

    return { viewer, model };
}

export function applyPerformanceProfile(viewer: Autodesk.Viewing.GuiViewer3D, profile: 'HighPerformance' | 'Balanced') {
    if (profile === 'HighPerformance') {
        // Maximum performance settings for best FPS and bidirectional interaction speed
        viewer.setProgressiveRendering(true);
        viewer.setQualityLevel(false, false); // No ambient shadows, no antialiasing
        viewer.setGroundShadow(false);
        viewer.setGroundReflection(false);
        viewer.setEnvMapBackground(false);
        viewer.setLightPreset(0); // Simple lighting
        
        // Additional performance optimizations
        if (viewer.prefs) {
            // Enable optimizeNavigation for smoother navigation
            try {
                viewer.prefs.set('optimizeNavigation', true);
            } catch (e) {
                // Ignore if not available
            }
            
            // Reduce render quality during navigation for better FPS
            try {
                viewer.prefs.set('navigationQuality', 'low');
            } catch (e) {
                // Ignore if not available
            }
            
            // Enable frustum culling optimization
            try {
                viewer.prefs.set('frustumCulling', true);
            } catch (e) {
                // Ignore if not available
            }
            
            // Enable aggressive culling for better frame rate
            try {
                viewer.prefs.set('aggressiveCulling', true);
            } catch (e) {
                // Ignore if not available
            }
            
            // Reduce update frequency during navigation
            try {
                viewer.prefs.set('navigationUpdateFrequency', 'low');
            } catch (e) {
                // Ignore if not available
            }
        }
        
        // Disable expensive features for maximum performance
        // Note: Background color will be set by applyDefaultViewerConfiguration for natural environment
        // Only reduce texture quality for faster loading and rendering
        if ((viewer as any).setTextureQuality) {
            (viewer as any).setTextureQuality(0.5); // Lower texture quality (50%)
        }
        
    } else {
        // Balanced - good performance with some visual quality
        viewer.setProgressiveRendering(true);
        viewer.setQualityLevel(true, true); // Ambient shadows, AA
        viewer.setGroundShadow(true);
        
        // Moderate optimizations
        if (viewer.prefs) {
            try {
                viewer.prefs.set('optimizeNavigation', true);
            } catch (e) {
                // Ignore if not available
            }
            
            try {
                viewer.prefs.set('frustumCulling', true);
            } catch (e) {
                // Ignore if not available
            }
        }
    }
    
    // Common optimizations for both profiles
    // Enable occlusion culling if available (hides objects behind others)
    if ((viewer as any).setOcclusionCulling) {
        (viewer as any).setOcclusionCulling(true);
    }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
export function applyDefaultViewerConfiguration(viewer: Autodesk.Viewing.GuiViewer3D, _customVisualContainer: HTMLElement) {
    // Enable large model mode (BETA) - CRITICAL for performance with large models
    viewer.prefs.set('largemodelmode', true);

    // Set natural environment for better visual experience
    // Preset 2 provides a natural, balanced lighting that looks good and performs well
    viewer.setEnvMapBackground(true); // Enable environment background for natural look
    viewer.setLightPreset(2); // Natural lighting preset (always active)

    // Configure selection mode for multi-selection support
    if (viewer.prefs) {
        // Ensure selection mode allows multiple selections
        console.log('Visual: Multi-selection enabled (Ctrl+Click to select multiple elements)');
        
        // Additional performance preferences
        try {
            // Enable aggressive culling for better frame rate
            viewer.prefs.set('aggressiveCulling', true);
        } catch (e) {
            // Ignore if not available
        }
        
        try {
            // Reduce update frequency during navigation
            viewer.prefs.set('navigationUpdateFrequency', 'low');
        } catch (e) {
            // Ignore if not available
        }
    }

    // Toolbar mantiene su configuración por defecto (centro, horizontal)
    // Las extensiones se cargarán automáticamente sin modificar la posición/orientación de la toolbar

    console.log('Visual: Applied optimized viewer configuration (large model mode, natural lighting preset 2, environment background enabled)');
}

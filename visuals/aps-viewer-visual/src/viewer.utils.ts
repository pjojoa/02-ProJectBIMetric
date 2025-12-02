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
 * Helper class for mapping between "dbIDs" (sequential numbers assigned to each design element;
 * typically used by the Viewer APIs) and "external IDs" (typically based on persistent IDs
 * from the authoring application, for example, Revit GUIDs).
 */
export class IdMapping {
    private readonly externalIdMappingPromise: Promise<{ [externalId: string]: number; }>;
    private readonly reverseMappingPromise: Promise<{ [dbid: number]: string; }>;

    constructor(private model: Autodesk.Viewing.Model) {
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

    getDbids(externalIds: string[]): Promise<number[]> {
        return this.externalIdMappingPromise
            .then(externalIdMapping => externalIds.map(externalId => externalIdMapping[externalId]));
    }

    getExternalIds(dbids: number[]): Promise<string[]> {
        return new Promise((resolve, reject) => {
            this.model.getBulkProperties(dbids, { propFilter: ['externalId'] }, results => {
                resolve(results.map(result => result.externalId))
            }, reject);
        });
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
     * Check if a string ID is clearly an ExternalId (not a numeric dbId)
     * ExternalIds typically contain letters, hyphens, or are longer than typical dbIds
     */
    private isExternalId(id: string): boolean {
        // If it contains non-numeric characters (except leading/trailing whitespace), it's an ExternalId
        // Also check if it's a GUID format (contains hyphens)
        return /[^0-9]/.test(id.trim()) || id.includes('-') || id.length > 10;
    }

    /**
     * Map ExternalIds to dbIds using External ID mapping
     * This is the PRIMARY method when working exclusively with ExternalIds
     */
    async mapExternalIdsToDbIds(externalIds: string[]): Promise<number[]> {
        const validDbIds: number[] = [];
        const tree = this.model.getInstanceTree();
        const invalidExternalIds: string[] = [];

        console.log(`Visual: IdMapping - Mapping ${externalIds.length} ExternalIds to dbIds using External ID mapping`);
        if (externalIds.length > 0) {
            console.log(`Visual: IdMapping - Sample ExternalIds (first 10):`, externalIds.slice(0, 10));
        }

        try {
            const externalIdMapping = await this.externalIdMappingPromise;
            const totalExternalIdsInMapping = Object.keys(externalIdMapping).length;
            console.log(`Visual: IdMapping - External ID mapping loaded with ${totalExternalIdsInMapping} entries.`);

            for (const externalId of externalIds) {
                const mappedDbId = externalIdMapping[externalId];
                if (mappedDbId != null && !isNaN(mappedDbId)) {
                    // Validate it exists in model
                    try {
                        if (tree && tree.getNodeType(mappedDbId) != null) {
                            validDbIds.push(mappedDbId);
                        } else {
                            invalidExternalIds.push(`${externalId} (mapped to DbId ${mappedDbId} but not in model)`);
                        }
                    } catch (e) {
                        invalidExternalIds.push(`${externalId} (mapped to DbId ${mappedDbId} but not in model)`);
                    }
                } else {
                    invalidExternalIds.push(`${externalId} (not found in External ID mapping)`);
                }
            }
        } catch (error) {
            console.error("Visual: IdMapping - Failed to load or use External ID mapping.", error);
        }

        console.log(`Visual: IdMapping - Mapped ${validDbIds.length} of ${externalIds.length} ExternalIds to dbIds`);
        if (invalidExternalIds.length > 0 && invalidExternalIds.length < 20) {
            console.warn(`Visual: IdMapping - Invalid/Unmapped ExternalIds:`, invalidExternalIds);
        } else if (invalidExternalIds.length >= 20) {
            console.warn(`Visual: IdMapping - ${invalidExternalIds.length} Invalid/Unmapped ExternalIds. Sample:`, invalidExternalIds.slice(0, 20));
        }

        return validDbIds;
    }

    /**
     * Smart mapping: tries External ID mapping first, then validates direct DbId
     * NOTE: When working exclusively with ExternalIds, use mapExternalIdsToDbIds instead
     */
    async smartMapToDbIds(ids: string[]): Promise<number[]> {
        const validDbIds: number[] = [];
        const tree = this.model.getInstanceTree();
        const unmappedIds: string[] = [];
        const invalidDbIds: string[] = [];
        const mappedViaExternalId: string[] = [];
        const mappedViaDirectDbId: string[] = [];

        console.log(`Visual: IdMapping - Starting smart mapping for ${ids.length} IDs.`);
        if (ids.length > 0) {
            console.log(`Visual: IdMapping - Sample input IDs (first 10):`, ids.slice(0, 10));
        }

        // Pass 1: Try direct dbId mapping ONLY for IDs that are clearly numeric (not ExternalIds)
        // This iterates all IDs and tries to validate them as numbers against the InstanceTree
        for (const id of ids) {
            // Skip if it's clearly an ExternalId (contains non-numeric chars, hyphens, or is too long)
            if (this.isExternalId(id)) {
                unmappedIds.push(id);
                continue;
            }

            const directDbId = parseInt(id, 10);
            if (!isNaN(directDbId) && directDbId > 0) { // dbIds are positive integers
                try {
                    // Check if node exists in the tree. getNodeType returns null/undefined if not found.
                    if (tree && tree.getNodeType(directDbId) != null) {
                        validDbIds.push(directDbId);
                        mappedViaDirectDbId.push(`${id} (used as DbId ${directDbId})`);
                        continue; // Success, move to next ID
                    } else {
                        // Numeric, but not found in tree. Might be an ExternalId that looks like a number?
                        unmappedIds.push(id);
                    }
                } catch (e) {
                    unmappedIds.push(id);
                }
            } else {
                // Not a number, must be ExternalId
                unmappedIds.push(id);
            }
        }

        // Pass 2: Fallback to External ID mapping (SLOW path, depends on promise)
        // Only executed if there are unmapped IDs
        if (unmappedIds.length > 0) {
            console.log(`Visual: IdMapping - ${unmappedIds.length} IDs not found as direct DbIds, attempting External ID mapping...`);
            try {
                const externalIdMapping = await this.externalIdMappingPromise;
                const totalExternalIdsInMapping = Object.keys(externalIdMapping).length;
                console.log(`Visual: IdMapping - External ID mapping loaded with ${totalExternalIdsInMapping} entries.`);

                for (const id of unmappedIds) {
                    const mappedDbId = externalIdMapping[id];
                    if (mappedDbId != null && !isNaN(mappedDbId)) {
                        // Validate it exists in model
                        try {
                            if (tree && tree.getNodeType(mappedDbId) != null) {
                                validDbIds.push(mappedDbId);
                                mappedViaExternalId.push(`${id} -> DbId ${mappedDbId}`);
                            } else {
                                invalidDbIds.push(`${id} (mapped to DbId ${mappedDbId} but not in model)`);
                            }
                        } catch (e) {
                            invalidDbIds.push(`${id} (mapped to DbId ${mappedDbId} but not in model)`);
                        }
                    } else {
                        invalidDbIds.push(`${id} (not found in External ID mapping)`);
                    }
                }
            } catch (error) {
                console.warn("Visual: IdMapping - Failed to load or use External ID mapping.", error);
                // If external mapping fails, we just return the valid direct dbIds we found
            }
        }

        // Logging results
        console.log(`Visual: IdMapping - Results: ${validDbIds.length} valid total.`);
        
        if (mappedViaDirectDbId.length > 0) {
            console.log(`Visual: IdMapping - Mapped via Direct DbId: ${mappedViaDirectDbId.length}. Sample:`, mappedViaDirectDbId.slice(0, 5));
        }
        if (mappedViaExternalId.length > 0) {
            console.log(`Visual: IdMapping - Mapped via External ID: ${mappedViaExternalId.length}. Sample:`, mappedViaExternalId.slice(0, 5));
        }
        if (invalidDbIds.length > 0 && invalidDbIds.length < 20) {
            console.warn(`Visual: IdMapping - Invalid/Unmapped IDs:`, invalidDbIds);
        } else if (invalidDbIds.length >= 20) {
            console.warn(`Visual: IdMapping - ${invalidDbIds.length} Invalid/Unmapped IDs. Sample:`, invalidDbIds.slice(0, 20));
        }

        return validDbIds;
    }
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
    api: string = 'streamingV2'
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
    const config: any = {
        extensions: ['BIMetric.GhostingExtension'], // Load our ghosting extension
        disabledExtensions: {
            measure: false,
            viewcube: false,
            explode: true,
            section: false,
            bimwalk: false, // Enable first-person navigation
            fusionOrbit: true,
            modelBrowser: false, // Enable Model Browser
            propertiesPanel: false,
            layerManager: true,
            hyperlink: true
        }
    };

    const viewer = new Autodesk.Viewing.GuiViewer3D(container, config as Autodesk.Viewing.ViewerConfig);
    viewer.start();

    // Load VisualClusters extension for better performance with large models
    await viewer.loadExtension('Autodesk.VisualClusters');

    // Load Model
    // skipPropertyDb: false by default, can be optimized later
    const model = await loadModel(viewer, urn, guid || undefined, false);

    // Apply Performance Profile
    applyPerformanceProfile(viewer, performanceProfile);

    // Apply default viewer configuration (pass the container for proper positioning)
    // Use setTimeout to ensure toolbar is created before scaling
    setTimeout(() => {
        applyDefaultViewerConfiguration(viewer, container);
    }, 100);

    // Setup Selection Listener
    viewer.addEventListener(Autodesk.Viewing.SELECTION_CHANGED_EVENT, (ev: any) => {
        const ids = ev?.dbIdArray ?? [];
        onDbIdsChanged(ids);
    });

    return { viewer, model };
}

export function applyPerformanceProfile(viewer: Autodesk.Viewing.GuiViewer3D, profile: 'HighPerformance' | 'Balanced') {
    if (profile === 'HighPerformance') {
        viewer.setProgressiveRendering(true);
        viewer.setQualityLevel(false, false); // No ambient shadows, no antialiasing
        viewer.setGroundShadow(false);
        viewer.setGroundReflection(false);
        viewer.setEnvMapBackground(false);
        viewer.setLightPreset(0); // Simple lighting

        // Optimize navigation
        if (viewer.prefs) {
            // Try to enable optimizeNavigation if available in prefs
            // viewer.prefs.set('optimizeNavigation', true); 
        }
    } else {
        // Balanced
        viewer.setProgressiveRendering(true);
        viewer.setQualityLevel(true, true); // Ambient shadows, AA
        viewer.setGroundShadow(true);
    }
}

export function applyDefaultViewerConfiguration(viewer: Autodesk.Viewing.GuiViewer3D, customVisualContainer: HTMLElement) {
    // Enable large model mode (BETA)
    viewer.prefs.set('largemodelmode', true);

    // Set environment background visible
    viewer.setEnvMapBackground(true);

    // Set environment to "Plaza" (ID: 8)
    viewer.setLightPreset(8);

    // Reduce toolbar size to 65%
    // Try multiple ways to access toolbar element
    let toolbarElement: HTMLElement | null = null;
    
    if (viewer.toolbar) {
        // Try container property
        toolbarElement = (viewer.toolbar as any).container as HTMLElement;
        
        // If not found, try to find toolbar in DOM
        if (!toolbarElement) {
            const toolbarInDom = document.querySelector('.adsk-viewing-viewer-toolbar') as HTMLElement;
            if (toolbarInDom) {
                toolbarElement = toolbarInDom;
            }
        }
        
        // If still not found, try viewer's toolbar container
        if (!toolbarElement && (viewer as any).toolbarContainer) {
            toolbarElement = (viewer as any).toolbarContainer as HTMLElement;
        }
    }
    
    if (toolbarElement) {
        toolbarElement.style.transform = 'scale(0.65)';
        toolbarElement.style.transformOrigin = 'center center';
        // Ensure toolbar remains visible and functional
        toolbarElement.style.opacity = '1';
        console.log('Visual: Toolbar scaled to 65%');
    } else {
        console.warn('Visual: Could not find toolbar element to scale');
    }

    console.log('Visual: Applied default viewer configuration (large model mode, Plaza, toolbar scaled to 65%)');
}

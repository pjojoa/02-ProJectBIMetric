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
import { initializeViewerRuntime, loadModel, IdMapping } from './viewer.utils';

/**
 * Custom visual wrapper for the Autodesk Platform Services Viewer.
 */
export class Visual implements IVisual {
    // Visual state
    private host: IVisualHost;
    private container: HTMLElement;
    private formattingSettings: VisualSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private currentDataView: DataView = null;
    private selectionManager: ISelectionManager = null;

    // Visual inputs
    private accessTokenEndpoint: string = '';

    // Viewer runtime
    private viewer: Autodesk.Viewing.GuiViewer3D = null;
    private urn: string = '';
    private guid: string = '';
    private externalIds: string[] = [];
    private model: Autodesk.Viewing.Model = null;
    private idMapping: IdMapping = null;

    /**
     * Initializes the viewer visual.
     * @param options Additional visual initialization options.
     */
    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();
        this.container = options.element;
        this.getAccessToken = this.getAccessToken.bind(this);
        this.onPropertiesLoaded = this.onPropertiesLoaded.bind(this);
        this.onSelectionChanged = this.onSelectionChanged.bind(this);
    }

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
        if (this.currentDataView.table?.rows?.length > 0) {
            const columns = this.currentDataView.table.columns;
            const urnIndex = columns.findIndex(c => c.roles['modelUrn']);
            const idIndex = columns.findIndex(c => c.roles['elementId']);

            if (urnIndex === -1 || idIndex === -1) {
                // Data roles not properly mapped
                return;
            }

            const rows = this.currentDataView.table.rows;
            const urns = this.collectDesignUrns(this.currentDataView, urnIndex);
            if (urns.length > 1) {
                this.showNotification('Multiple design URNs detected. Only the first one will be displayed.');
            }
            if (urns[0] !== this.urn) {
                this.urn = urns[0];
                this.updateModel();
            }
            this.externalIds = rows.map(r => r[idIndex].valueOf() as string);
        } else {
            this.urn = '';
            this.externalIds = [];
            this.updateModel();
        }

        if (this.idMapping) {
            const isDataFilterApplied = this.currentDataView.metadata?.isDataFilterApplied;
            if (this.externalIds.length > 0 && isDataFilterApplied) {
                const dbids = await this.idMapping.getDbids(this.externalIds);
                this.viewer.isolate(dbids);
                this.viewer.fitToView(dbids);
            } else {
                this.viewer.isolate();
                this.viewer.fitToView();
            }
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
            const env = this.formattingSettings?.viewerCard?.viewerEnv?.value || 'AutodeskProduction2';
            const region = this.formattingSettings?.viewerCard?.viewerRegion?.value || 'US';

            // API depends on Env: SVF2 -> streamingV2, SVF -> derivativeV2
            const api = env.includes('Production2') ? 'streamingV2' : 'derivativeV2';

            await initializeViewerRuntime({
                env: env,
                api: api,
                region: region,
                getAccessToken: this.getAccessToken
            });
            this.container.innerText = '';
            this.viewer = new Autodesk.Viewing.GuiViewer3D(this.container);
            this.viewer.start();
            this.viewer.addEventListener(Autodesk.Viewing.OBJECT_TREE_CREATED_EVENT, this.onPropertiesLoaded);
            this.viewer.addEventListener(Autodesk.Viewing.SELECTION_CHANGED_EVENT, this.onSelectionChanged);
            if (this.urn) {
                this.updateModel();
            }
        } catch (err) {
            this.showNotification('Could not initialize viewer runtime. Please see console for more details.');
            console.error(err);
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

        if (this.model && this.model.getData().urn !== this.urn) {
            this.viewer.unloadModel(this.model);
            this.model = null;
            this.idMapping = null;
        }

        try {
            if (this.urn) {
                // Sanitize URN: remove whitespace and 'urn:' prefix if present
                let safeUrn = this.urn.trim();
                if (safeUrn.toLowerCase().startsWith('urn:')) {
                    safeUrn = safeUrn.substring(4);
                }
                this.model = await loadModel(this.viewer, safeUrn, this.guid);
            }
        } catch (err) {
            let decodedUrn = '';
            try {
                decodedUrn = atob(this.urn);
            } catch (e) {
                decodedUrn = 'Invalid Base64';
            }
            let msg = `Could not load model. URN: ${this.urn.substring(0, 10)}... Decoded: ${decodedUrn.substring(0, 50)}... `;
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

    private async onPropertiesLoaded() {
        this.idMapping = new IdMapping(this.model);
    }

    private async onSelectionChanged() {
        const allExternalIds = this.currentDataView.table.rows;
        if (!allExternalIds) {
            return;
        }

        const columns = this.currentDataView.table.columns;
        const idIndex = columns.findIndex(c => c.roles['elementId']);
        if (idIndex === -1) return;

        const selectedDbids = this.viewer.getSelection();
        const selectedExternalIds = await this.idMapping.getExternalIds(selectedDbids);
        const selectionIds: powerbi.extensibility.ISelectionId[] = [];
        for (const selectedExternalId of selectedExternalIds) {
            const rowIndex = allExternalIds.findIndex(row => row[idIndex] === selectedExternalId);
            if (rowIndex !== -1) {
                const selectionId = this.host.createSelectionIdBuilder()
                    .withTable(this.currentDataView.table, rowIndex)
                    .createSelectionId();
                selectionIds.push(selectionId);
            }
        }
        this.selectionManager.select(selectionIds);
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

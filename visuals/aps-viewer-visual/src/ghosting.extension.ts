/// <reference types="@types/forge-viewer" />

export function registerGhostingExtension() {
    if (typeof Autodesk === 'undefined' || !Autodesk.Viewing) return;

    function GhostingExtension(viewer: Autodesk.Viewing.GuiViewer3D, options: any) {
        Autodesk.Viewing.Extension.call(this, viewer, options);
        this.button = null;
        this._group = null;
    }

    GhostingExtension.prototype = Object.create(Autodesk.Viewing.Extension.prototype);
    GhostingExtension.prototype.constructor = GhostingExtension;

    GhostingExtension.prototype.load = function () {
        console.log('Visual: GhostingExtension loaded');
        return true;
    };

    GhostingExtension.prototype.unload = function () {
        if (this.button) {
            this.removeToolbarButton(this.button);
            this.button = null;
        }
        return true;
    };

    GhostingExtension.prototype.onToolbarCreated = function (toolbar: Autodesk.Viewing.UI.ToolBar) {
        this.button = new Autodesk.Viewing.UI.Button('ghosting-toggle-button');

        // Initial state check
        const isGhosting = this.viewer.prefs.get(Autodesk.Viewing.Private.Prefs3D.GHOSTING);
        this.updateButtonState(isGhosting);

        this.button.setToolTip('Objetos fantasma');

        // Create a Pac-Man style ghost icon - simple sheet with eyes
        const ghostIcon = `
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <!-- Ghost body - simple rounded top like a sheet -->
                <path d="M12 3C8.13 3 5 6.13 5 10c0 2.5 1.25 4.7 3.2 6V20c0 .55.45 1 1 1h3.6v-3h.4v3H16c.55 0 1-.45 1-1v-4c1.95-1.3 3.2-3.5 3.2-6 0-3.87-3.13-7-7-7z" fill="currentColor"/>
                <!-- Large oval left eye -->
                <ellipse cx="9" cy="10" rx="2.5" ry="3" fill="#ffffff"/>
                <!-- Large oval right eye -->
                <ellipse cx="15" cy="10" rx="2.5" ry="3" fill="#ffffff"/>
                <!-- Bottom wavy edge - classic Pac-Man ghost style with 3 waves -->
                <path d="M5 19c1 0 2-.3 3-.3s2 .3 3 .3 2-.3 3-.3 2 .3 3 .3" stroke="currentColor" stroke-width="2" stroke-linecap="round" fill="none"/>
            </svg>
        `;
        
        // Set the icon directly
        if (this.button.icon) {
            // eslint-disable-next-line powerbi-visuals/no-inner-outer-html
            this.button.icon.innerHTML = ghostIcon;
            // Ensure the icon is visible and properly styled
            this.button.icon.style.display = 'flex';
            this.button.icon.style.alignItems = 'center';
            this.button.icon.style.justifyContent = 'center';
            this.button.icon.style.width = '24px';
            this.button.icon.style.height = '24px';
        }

        const self = this;
        this.button.onClick = function () {
            const current = self.viewer.prefs.get(Autodesk.Viewing.Private.Prefs3D.GHOSTING);
            const newState = !current;
            self.viewer.setGhosting(newState);
            self.updateButtonState(newState);
            console.log('Visual: Ghosting set to ' + newState);
        };

        // Add to toolbar
        // Try to add to 'modelTools' group first, or create our own
        this._group = toolbar.getControl('modelTools') as Autodesk.Viewing.UI.ControlGroup;
        if (!this._group) {
            this._group = new Autodesk.Viewing.UI.ControlGroup('bimetric-tools');
            toolbar.addControl(this._group);
        }

        this._group.addControl(this.button);
    };

    GhostingExtension.prototype.updateButtonState = function (isGhosting: boolean) {
        if (this.button) {
            const state = isGhosting ? Autodesk.Viewing.UI.Button.State.ACTIVE : Autodesk.Viewing.UI.Button.State.INACTIVE;
            this.button.setState(state);
        }
    };

    GhostingExtension.prototype.removeToolbarButton = function (button: Autodesk.Viewing.UI.Button) {
        if (this._group) {
            this._group.removeControl(button);
            if (this._group.getNumberOfControls() === 0) {
                this.viewer.toolbar.removeControl(this._group);
            }
        }
    };

    // Register the extension
    // Autodesk.Viewing.theExtensionManager.unregisterExtension('BIMetric.GhostingExtension'); // Optional: unregister first to be safe
    Autodesk.Viewing.theExtensionManager.registerExtension('BIMetric.GhostingExtension', GhostingExtension);
}

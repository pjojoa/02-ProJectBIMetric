/// <reference types="@types/forge-viewer" />

// eslint-disable-next-line max-lines-per-function
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

        // Improved ghost icon with better design and visibility
        const ghostIcon = `
            <svg width="28" height="28" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                <!-- Ghost body (sheet-like with rounded top) -->
                <path d="M 7 4 Q 12 2 17 4 Q 19 5 19 8 L 19 18 Q 18 17.5 17 18 Q 16 18.5 15 18 Q 14 17.5 13 18 Q 12 18.5 11 18 Q 10 17.5 9 18 Q 8 18.5 7 18 Q 6 17.5 5 18 L 5 8 Q 5 5 7 4 Z" 
                      fill="currentColor" opacity="0.9"/>
                
                <!-- Shadow/depth on right side -->
                <path d="M 15 4 Q 17.5 4.5 18.5 6.5 Q 19 7.5 19 8 L 19 18 Q 18 17.5 17 18 Q 16.5 18.3 16 18.2 L 16 5.5 Q 15.8 4.5 15 4 Z" 
                      fill="currentColor" opacity="0.3"/>
                
                <!-- Left eye (white background) -->
                <circle cx="9.5" cy="9.5" r="2" fill="white"/>
                <!-- Right eye (white background) -->
                <circle cx="14.5" cy="9.5" r="2" fill="white"/>
                
                <!-- Left pupil -->
                <circle cx="9.7" cy="9.4" r="1.2" fill="#1a1a1a"/>
                <!-- Right pupil -->
                <circle cx="14.7" cy="9.4" r="1.2" fill="#1a1a1a"/>
                
                <!-- Eye highlights -->
                <circle cx="9.3" cy="9" r="0.4" fill="white"/>
                <circle cx="14.3" cy="9" r="0.4" fill="white"/>
                
                <!-- Mouth (curved smile) -->
                <path d="M 10 13 Q 12 14.5 14 13" fill="none" stroke="currentColor" stroke-width="1" stroke-linecap="round" opacity="0.8"/>
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
            this.button.icon.style.width = '28px';
            this.button.icon.style.height = '28px';
        }

        // eslint-disable-next-line @typescript-eslint/no-this-alias
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
            // Set button state (ACTIVE = blue background, INACTIVE = default)
            const state = isGhosting ? Autodesk.Viewing.UI.Button.State.ACTIVE : Autodesk.Viewing.UI.Button.State.INACTIVE;
            this.button.setState(state);

            // The icon color is controlled by CSS via 'currentColor'
            // When ACTIVE, the toolbar automatically applies blue background
            // The icon will inherit the appropriate color from the button state
            console.log('Visual: Ghosting button state updated to ' + (isGhosting ? 'ACTIVE (blue)' : 'INACTIVE'));
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

'use strict';

import { formattingSettings } from 'powerbi-visuals-utils-formattingmodel';

import Card = formattingSettings.SimpleCard;
import Slice = formattingSettings.Slice;
import Model = formattingSettings.Model;

class ViewerCard extends Card {
    accessTokenEndpoint = new formattingSettings.TextInput({
        name: 'accessTokenEndpoint',
        displayName: 'Access Token Endpoint',
        description: 'URL that the viewer can call to generate access tokens.',
        placeholder: '',
        value: 'https://zero2-projectbimetric.onrender.com/token'
    });
    viewerEnv = new formattingSettings.TextInput({
        name: 'viewerEnv',
        displayName: 'Viewer Environment',
        description: 'AutodeskProduction (Universal) or AutodeskProduction2 (SVF2 only)',
        placeholder: 'AutodeskProduction',
        value: 'AutodeskProduction'
    });
    viewerRegion = new formattingSettings.TextInput({
        name: 'viewerRegion',
        displayName: 'Region',
        description: 'US or EMEA',
        placeholder: 'US',
        value: 'US'
    });
    performanceProfile = new formattingSettings.ItemDropdown({
        name: 'performanceProfile',
        displayName: 'Performance Profile',
        description: 'Select performance profile for the viewer',
        value: {
            value: 'HighPerformance',
            displayName: 'High Performance'
        },
        items: [
            {
                value: 'HighPerformance',
                displayName: 'High Performance'
            },
            {
                value: 'Balanced',
                displayName: 'Balanced'
            }
        ]
    });
    skipPropertyDb = new formattingSettings.ToggleSwitch({
        name: 'skipPropertyDb',
        displayName: 'Fast Load (Skip Property DB)',
        description: 'Skip loading property database for faster initial load (30-50% faster). Properties will not be available but visualization and bidirectional interaction work normally.',
        value: false
    });
    name: string = 'viewer';
    displayName: string = 'Viewer Runtime';
    slices: Array<Slice> = [this.accessTokenEndpoint, this.viewerEnv, this.viewerRegion, this.performanceProfile, this.skipPropertyDb];
}

export class DataPointCard extends Card {
    name: string = "dataPoint";
    displayName: string = "Data Colors";

    // Slices will be populated dynamically in visual.ts
    slices: Array<Slice> = [];
}

export class VisualSettingsModel extends Model {
    viewerCard = new ViewerCard();
    dataPointCard = new DataPointCard();
    cards = [this.viewerCard, this.dataPointCard];
}

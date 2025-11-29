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
        description: 'AutodeskProduction2 (SVF2) or AutodeskProduction (SVF)',
        placeholder: 'AutodeskProduction2',
        value: 'AutodeskProduction2'
    });
    viewerRegion = new formattingSettings.TextInput({
        name: 'viewerRegion',
        displayName: 'Region',
        description: 'US or EMEA',
        placeholder: 'US',
        value: 'US'
    });
    name: string = 'viewer';
    displayName: string = 'Viewer Runtime';
    slices: Array<Slice> = [this.accessTokenEndpoint, this.viewerEnv, this.viewerRegion];
}

export class VisualSettingsModel extends Model {
    viewerCard = new ViewerCard();
    cards = [this.viewerCard];
}

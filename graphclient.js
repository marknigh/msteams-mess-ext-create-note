// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { Client } = require('@microsoft/microsoft-graph-client');

/**
 * This class is a wrapper for the Microsoft Graph API.
 * See: https://developer.microsoft.com/en-us/graph for more information.
 */
class GraphClient {
    constructor(token) {
        if (!token || !token.trim()) {
            throw new Error('GraphClient: Invalid token received.');
        }

        this._token = token;

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
        this.graphClient = Client.init({
            authProvider: (done) => {
                done(null, this._token); // First parameter takes an error if you can't get an access token.
            },
        });
    }

    async GetMyProfile() {
        return await this.graphClient.api('/me').get();
    }

    async GetMyNoteBooks() {
        return await this.graphClient.api('/me/onenote/notebooks').get()
    }

    async GetMySections(noteBookId) {
        return await this.graphClient.api('/me/onenote/notebooks/' + noteBookId + '/sections').get()
    }

    async CreatePage(action) {
        const date = new Date().toDateString();
        const messageText = action.messagePayload.body.content

        // check if invoke came from message or commandbox. If message, also include message content into oneNote page.
        const note = action.commandContext == 'commandbox' ? 'no notes, from commandbox' : messageText 
        let body = 
            `<!DOCTYPE html>
            <html>
            <head>
                <title>${date}</title>
            </head>
            <body>
                <p>${action.data.Note}</p>
                <p>${note}</p>
            </body>
            </html>`
        
        return await this.graphClient.api('me/onenote/sections/' + action.data.SectionNames + '/pages').headers({ 'Content-type': 'application/xhtml+xml' }).post(body)
    
    }
    
    // Gets the user's photo
    async GetPhotoAsync(token) {
        let graphPhotoEndpoint = 'https://graph.microsoft.com/v1.0/me/photos/240x240/$value';
        let graphRequestParams = {
            method: 'GET',
            headers: {
                'Content-Type': 'image/png',
                "authorization": "bearer " + token
            }            
        }

        let response = await fetch(graphPhotoEndpoint, graphRequestParams).catch(this.unhandledFetchError);
        if (!response.ok) {
            console.error("ERROR: ", response);
        }

        let imageBuffer = await response.arrayBuffer().catch(this.unhandledFetchError); //Get image data as raw binary data

        //Convert binary data to an image URL and set the url in state
        const imageUri = 'data:image/png;base64,' + Buffer.from(imageBuffer).toString('base64');
        return imageUri;
    }
}

exports.GraphClient = GraphClient;

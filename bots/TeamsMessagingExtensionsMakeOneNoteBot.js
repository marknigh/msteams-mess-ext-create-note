// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
// Modified by Mark Nigh, January 2023 as a training to understand MS teams applicatoin development specifically around
// messaging extension - action

const {
    TeamsActivityHandler,
    CardFactory,
} = require('botbuilder');

const { polyfills } = require('isomorphic-fetch');

const { GraphClient } = require('../graphclient.js');

// Import Adaptive Cards
const { createNotebookCard } = require('../cards/NotebookCard')
const { createSectionCard } = require('../cards/SectionsCard')
const { createSuccessCard } = require('../cards/SuccessMessageCard')

// User Configuration property name
const USER_CONFIGURATION = 'userConfigurationProperty';

class TeamsMessagingExtensionsMakeOneNoteBot extends TeamsActivityHandler {
    /**
     *
     * @param {UserState} User state to persist configuration settings
     */
    constructor(userState) {
        super();
        // Creates a new user property accessor.
        // See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors.
        this.userConfigurationProperty = userState.createProperty(
            USER_CONFIGURATION
        );

        // 
        this.onMessage(async (context, next) => {
            // Sends a message activity to the sender of the incoming activity.
            await context.sendActivity(`Echo: '${context.activity.text}'`);
            await next();
        });

        this.connectionName = process.env.ConnectionName;
        this.userState = userState;
    }

    /**
     * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
     */
    async run(context) {
        await super.run(context);

        // Save state changes
        await this.userState.saveChanges(context);
    }

    // Overloaded function. Receives invoke activities with the name 'composeExtension/setting'
    async handleTeamsMessagingExtensionConfigurationSetting(context, settings) {
        // When the user submits the settings page, this event is fired.
        if (settings.state != null) {
            await this.userConfigurationProperty.set(context, settings.state);
        }
    }

    // Overloaded function. Receives invoke activities with the name 'composeExtension/fetchTask'
    async handleTeamsMessagingExtensionFetchTask(context, action) {

        if (action.commandId === 'MAKENOTE') {
            const magicCode =
                action.state && Number.isInteger(Number(action.state))
                    ? action.state
                    : '';
            const tokenResponse = await context.adapter.getUserToken(
                context,
                this.connectionName,
                magicCode
            );

            if (!tokenResponse || !tokenResponse.token) {
                // There is no token, so the user has not signed in yet.
                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions

                const signInLink = await context.adapter.getSignInLink(
                    context,
                    this.connectionName
                );

                return {
                    composeExtension: {
                        type: 'silentAuth',
                        suggestedActions: {
                            actions: [
                                {
                                    type: 'openUrl',
                                    value: signInLink,
                                    title: 'Bot Service OAuth'
                                },
                            ],
                        },
                    },
                };
            }

            const graphClient = new GraphClient(tokenResponse.token);
            const myNoteBooks = await graphClient.GetMyNoteBooks()
            const myNoteBookNames = []
            for (const noteBook of myNoteBooks.value) {
                myNoteBookNames.push({'value': noteBook.id, 'title': noteBook.displayName})
            }

            const noteBookCard = createNotebookCard(myNoteBookNames)

            return {
                task: {
                    type: 'continue',
                    value: {
                        card: noteBookCard,
                        heigth: 250,
                        width: 400,
                        title: 'Your oneNote NoteBooks'
                    },
                },
            };
        }
        if (action.commandId === 'SignOutCommand') {
            const adapter = context.adapter;
            await adapter.signOutUser(context, this.connectionName);

            const card = CardFactory.adaptiveCard({
                version: '1.0.0',
                type: 'AdaptiveCard',
                body: [
                    {
                        type: 'TextBlock',
                        text: 'You have been signed out.'
                    },
                ],
                actions: [
                    {
                        type: 'Action.Submit',
                        title: 'Close',
                        data: {
                            key: 'close'
                        },
                    },
                ],
            });

            return {
                task: {
                    type: 'continue',
                    value: {
                        card: card,
                        heigth: 200,
                        width: 400,
                        title: 'Adaptive Card: Inputs'
                    },
                },
            };
        }

        return null;
    }

    async handleTeamsMessagingExtensionSubmitAction(context, action) {
        // This method is to handle the 'Close' button on the confirmation Task Module after the user signs out.
        
        const magicCode =
                action.state && Number.isInteger(Number(action.state))
                    ? action.state
                    : '';
        const tokenResponse = await context.adapter.getUserToken(
            context,
            this.connectionName,
            magicCode
        );

        if (!tokenResponse || !tokenResponse.token) {
            // There is no token, so the user has not signed in yet.
            // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions

            const signInLink = await context.adapter.getSignInLink(
                context,
                this.connectionName
            );

            return {
                composeExtension: {
                    type: 'silentAuth',
                    suggestedActions: {
                        actions: [
                            {
                                type: 'openUrl',
                                value: signInLink,
                                title: 'Bot Service OAuth'
                            },
                        ],
                    },
                },
            };
        }

        const graphClient = new GraphClient(tokenResponse.token);

        switch(action.data.response) {
            case "showSections": 

                const mySections = await graphClient.GetMySections(action.data.NoteBookNames)
                const mySectionsNames = []
                for (const noteBook of mySections.value) {
                    mySectionsNames.push({'value': noteBook.id, 'title': noteBook.displayName})
                }
                
                const sectionsCard = createSectionCard(mySectionsNames)

                return {
                    task: {
                        type: 'continue',
                        value: {
                            card: sectionsCard,
                            heigth: 250,
                            width: 400,
                            title: 'Your Sections'
                        },
                    },
                };
            
            case 'createPage':

                await graphClient.CreatePage(action)
                
                const successCard = createSuccessCard()

                return {
                    task: {
                        type: 'continue',
                        value: {
                            card: successCard,
                            heigth: 200,
                            width: 300,
                            title: 'Page was successfully created'
                        },
                    },
                };
            
            default:
                return {}
        }
    }


    async onInvokeActivity(context) {
        console.log('onInvoke, ' + context.activity.name);
        const valueObj = context.activity.value;
        if (valueObj.authentication) {
            const authObj = valueObj.authentication;
            if (authObj.token) {
                // If the token is NOT exchangeable, then do NOT deduplicate requests.
                if (await this.tokenIsExchangeable(context)) {
                    return await super.onInvokeActivity(context);
                }
                else {
                    const response = {
                        status: 412
                    };
                    return response;
                }
            }
        }

        return await super.onInvokeActivity(context);
             
    }

    async tokenIsExchangeable(context) {
        let tokenExchangeResponse = null;
        try {
            const valueObj = context.activity.value;
            const tokenExchangeRequest = valueObj.authentication;
            console.log("tokenExchangeRequest.token: " + tokenExchangeRequest.token);

            tokenExchangeResponse = await context.adapter.exchangeToken(context,
                process.env.connectionName,
                context.activity.from.id,
                { token: tokenExchangeRequest.token });
            console.log('tokenExchangeResponse: ' + JSON.stringify(tokenExchangeResponse));
        } catch (err) {
            console.log('tokenExchange error: ' + err);
            // Ignore Exceptions
            // If token exchange failed for any reason, tokenExchangeResponse above stays null , and hence we send back a failure invoke response to the caller.
        }
        if (!tokenExchangeResponse || !tokenExchangeResponse.token) {
            return false;
        }

        console.log('Exchanged token: ' + tokenExchangeResponse.token);
        return true;
    }
}

module.exports.TeamsMessagingExtensionsMakeOneNoteBot = TeamsMessagingExtensionsMakeOneNoteBot;
/**
 * Copyright(c) Microsoft Corporation.All rights reserved.
 * Licensed under the MIT License.
 */

import {
    ActivityHandler,
    BotTelemetryClient,
    ConversationState,
    EndOfConversationCodes,
    Severity,
    TurnContext,
    UserState,
    ConversationReference,
    StatePropertyAccessor,
} from 'botbuilder';
import {
    Dialog,
    DialogContext,
    DialogSet,
    DialogState,
} from 'botbuilder-dialogs';
import { ADAuth } from '../auth/adAuth';
import { IOnboardingState } from '../models/onboardingState';
import { MongoDBConversationService, LogConversationReferenceMsg } from '../services/mongoDBConversationService';
import { DefaultAdapter } from '../adapters/defaultAdapter';
import { AnalyticsService, TimeoutMsg } from '../services/analyticsService';
import { DIALOG_CONTEXT_KEY, TIMEOUT_MESSAGE, NOT_AVAILABLE } from '../config/constants';
import { inspect } from 'util'


const WELCOMED_USER = 'welcomedUserProperty';
const WELCOMEDBACK_USER = 'welcomedBackUserProperty';

export class DialogBot<T extends Dialog> extends ActivityHandler {

    private adapter: DefaultAdapter;

    private readonly telemetryClient: BotTelemetryClient;
    private readonly solutionName: string = 'ginaCore';
    private readonly rootDialogId: string;
    private readonly dialogs: DialogSet;
    private readonly userState: UserState;
    private readonly conversationState: ConversationState;
    private readonly mongoDbConversationService: MongoDBConversationService;

    // Are those (welcomedUserProperty & welcomedBackUserProperty) being used anywhere?
    // user state property used to track if the user was ever welcomed
    private welcomedUserProperty: StatePropertyAccessor<boolean>;
    // conversation state property used to track if the user was welcomed in the active conversation
    private welcomedBackUserProperty: StatePropertyAccessor<boolean>;

    private adAuth: ADAuth;

    constructor(
        adapter: DefaultAdapter, 
        userState: UserState, 
        conversationState: ConversationState, 
        telemetryClient: BotTelemetryClient,
        dialog: T, 
        adAuth: ADAuth, 
        mongoDbConversationService: MongoDBConversationService
    ) {
        super();

        this.adapter = adapter;

        this.rootDialogId = dialog.id;
        this.telemetryClient = telemetryClient;
        this.userState = userState;
        this.conversationState = conversationState;

        this.welcomedUserProperty = userState.createProperty(WELCOMED_USER);
        this.welcomedBackUserProperty = conversationState.createProperty(WELCOMEDBACK_USER);

        this.adAuth = adAuth;
        this.mongoDbConversationService = mongoDbConversationService;

        this.dialogs = new DialogSet(conversationState.createProperty<DialogState>(this.solutionName));
        this.dialogs.add(dialog);

        this.onTurn(this.turn.bind(this));
        this.onMessage(this.message.bind(this));
    }

    /**
     * On every message track the conversation reference obj
     * @param turnContext 
     * @param next 
     */
    public async message(turnContext: TurnContext, next: () => Promise<void>): Promise<any> {
        this.addConversationReferenceToMongo(turnContext);
        await next();
    }

    //tslint:disable-next-line: no-any
    public async turn(turnContext: TurnContext, next: () => Promise<void>): Promise<any> {
        // logs needed for testing
        console.log(`DialogBot: Turn context before setting dialog context: ${inspect(turnContext)}`);
        console.log(`DialogBot: Telementry client: ${inspect(this.telemetryClient)}`);

        // Client notifying this bot took to long to respond (timed out)
        if (turnContext.activity.code === EndOfConversationCodes.BotTimedOut) {
			if (this.telemetryClient) {
				this.telemetryClient.trackTrace({
					message: `Timeout in ${turnContext.activity.channelId} channel: Bot took too long to respond`,
					severityLevel: Severity.Information
				});

				return;
			}
		}

        try {
            await this.userAuthenticate(turnContext);
        } catch (err) {
            console.log('Error during user authentication: ', JSON.stringify(err));
            await turnContext.sendActivity('There was a problem during authentication. Contact system administrator');
        }

        const dc: DialogContext = await this.dialogs.createContext(turnContext);

        // logs needed for testing
        console.log(`DialogBot: Dialog context: ${inspect(dc)}`);

        // Dialog context needs to be persisted in order to be later used
        // in TimeoutMiddleware.
        turnContext.turnState.set(DIALOG_CONTEXT_KEY, dc);

        // logs needed for testing
        console.log(`DialogBot: Turn context after setting dialog context: ${inspect(turnContext)}`);

        try {
            if (dc.activeDialog !== undefined) {
                console.log('Continue active dialog: ', dc.activeDialog.id);
                await dc.continueDialog();
            } else {
                console.log('No active dialog, starting root dialog: ', this.rootDialogId);
                await dc.beginDialog(this.rootDialogId);
            }

            await next();
        } catch (err) {
            console.log('Error in dialog turn', JSON.stringify(err));
            throw err;
        }
    }

    public async userAuthenticate(turnContext: TurnContext) {
        // check if user is authenticated
        let onboardingState: IOnboardingState | undefined;

        try {
            if ((onboardingState = await this.adAuth.getOnboardingState(turnContext)) === undefined) {
                onboardingState = await this.adAuth.loadUser(turnContext);
 
                // logs needed for testing
                console.log(`DialogBot: Onboarding state: ${inspect(onboardingState)}`);
            }

            console.log('Setting current user to user:', onboardingState.userId);
            turnContext.turnState.set('currentUser', onboardingState);
        }
        catch (err) {
            throw new Error(`There was an error in DialogBot.userAuthenticate(): ${JSON.stringify(err)}`);
        }
    }

    /**
     * For every user log the conversation reference obj to send push notifications when needed
     * @param context 
     */
    public async addConversationReferenceToMongo(context: TurnContext) {
        try {
            const conversationReference = TurnContext.getConversationReference(context.activity);

            // logs needed for testing
            console.log(`DialogBot: Conversation reference: ${inspect(conversationReference)}`);

            if (this.mongoDbConversationService) {
                await this.mongoDbConversationService.saveConversationReferenceObj(new LogConversationReferenceMsg(conversationReference, context));
            }
        } catch (err) {
            console.log('Error while saving the conversationObj', JSON.stringify(err));
        }
    }
    /**
     * Sends a proactive timeout message to a user that has timed out.
     * Also, it will cancel all the dialogs that are waiting for a 
     * response from the timed out user, as well as log an analytics
     * document containing timeout information.
     * 
     * This method will be called by the timeout API only!
     *
     * @param {Partial<ConversationReference>} conversationReference
     * @param {AnalyticsService} analyticsService
     * @param {string} userId
     * @returns {Promise<void>}
     * @memberof DialogBot
     */
    public async sendProactiveTimeoutMessage(
        conversationReference: Partial<ConversationReference>,
        analyticsService: AnalyticsService,
        userId: string,
        skillInstanceId: string
    ): Promise<void> {
        await this.adapter.continueConversation(conversationReference, async context => {
            try {
                const dialogContext: DialogContext = await this.dialogs.createContext(context);

                // logs needed for testing
                console.log(`DialogBot Timeout: Dialog context: ${inspect(conversationReference)}`);

                // If the user hasn't explicitly ended the latest 
                // running dialog, it will be canceled here.
                if (dialogContext.activeDialog?.state.dialogs?.dialogStack?.length > 0) {
                    // Log skill status document to analytics.
                    await analyticsService.saveTimeout(
                        new TimeoutMsg(
                            conversationReference.conversation?.id ?? NOT_AVAILABLE,
                            conversationReference.channelId ?? NOT_AVAILABLE,
                            userId,
                            skillInstanceId,
                            TIMEOUT_MESSAGE
                        )
                    );

                    // Canceling all the active dialogs.
                    await dialogContext.cancelAllDialogs();

                    // Sending the Adaptive Card reply.
                    await context.sendActivity(TIMEOUT_MESSAGE);

                    // Saving context changes.
                    await this.conversationState.saveChanges(context);
                } else {
                    console.log("The active dialogs stack is empty, no need to cancel any dialog.");
                }
            } catch (e) {
                console.error(`Error when sending timeout message or logging in analytics with details ${JSON.stringify(e)}`);
            }
        });
    }
}

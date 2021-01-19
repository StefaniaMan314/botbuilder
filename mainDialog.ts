/**
 * Copyright(c) Microsoft Corporation.All rights reserved.
 * Licensed under the MIT License.
 */

import { BotTelemetryClient, StatePropertyAccessor } from "botbuilder";
import { DialogContext, DialogTurnResult, DialogTurnStatus, DialogInstance } from "botbuilder-dialogs";
import { ISkillManifest, SkillDialog, SkillRouter, SkillContext } from "botbuilder-skills";
import { InterruptionAction, RouterDialog, TokenEvents } from "botbuilder-solutions";
import { Activity, ActivityTypes } from "botframework-schema";
import short from "short-uuid";
import { IOnboardingState } from "../models/onboardingState";
import { BotServices } from "../services/botServices";
import { IBotSettings } from "../services/botSettings";
import { CancelDialog } from "./cancelDialog";

import { SkillRecognizer, SkillRecognizerTelemetryClient } from "../dispatcher/skillRecognizer";
import { SmartIntentDialog } from "../dialogs/smartintent/smartIntentDialog";
import { AnalyticsService, UserInputMsg, NlpDataMsg, SkillstatusMsg } from "../services/analyticsService";
import { ControlSkillEnum, SmartIntentSkillEnum,GeneralReplies, SkillCompletionFlag, CANCEL_TEXT_LABEL } from "../config/constants";
import { NoneIntentDialog } from "./noneIntentDialog";
import { inspect } from "util";

export class MainDialog extends RouterDialog {
	// Fields
	private readonly settings: Partial<IBotSettings>;
	private readonly analyticsSvc: AnalyticsService;
	private readonly skillContextAccessor: StatePropertyAccessor<SkillContext>;
	public skillRecognizer: SkillRecognizerTelemetryClient;

	private skillInstanceId: string;

	// Constructor
	constructor(
		settings: Partial<IBotSettings>,
		services: BotServices,
		cancelDialog: CancelDialog,
		skillDialogs: SkillDialog[],
		skillContextAccessor: StatePropertyAccessor<SkillContext>,
		onboardingAccessor: StatePropertyAccessor<IOnboardingState>,
		telemetryClient: BotTelemetryClient,
		analyticsSvc: AnalyticsService
	) {
		super(MainDialog.name, telemetryClient);
		this.settings = settings;
		this.skillContextAccessor = skillContextAccessor;
		this.telemetryClient = telemetryClient;

		this.addDialog(cancelDialog);

		skillDialogs.forEach((skillDialog: SkillDialog) => {
			this.addDialog(skillDialog);
		});
		this.skillRecognizer = new SkillRecognizer(false, this.telemetryClient);
		this.analyticsSvc = analyticsSvc;

		this.skillInstanceId = "";
	}

	protected async route(dc: DialogContext): Promise<void> {
		// logs needed for testing
		console.log(`MainDialog route: Dialog Context: ${inspect(dc)}`);

		// Check dispatch result
		let responseObj = await this.skillRecognizer.recognize(dc.context);

		const intent: string = responseObj.intent;
		console.log(`MainDialog:route Detected intent is ${intent} for ${dc.context.activity.text} `);

		if (!this.settings?.skills) {
			throw new Error("There is no skills in settings value");
		}

		// Identify if the dispatch intent matches any Action within a Skill if so, we pass to the appropriate SkillDialog to hand-off
		const identifiedSkill: ISkillManifest | undefined = SkillRouter.isSkill(this.settings.skills, intent);
		if (identifiedSkill) {

			// We have identified a skill so initialize the skill connection with the target skill
			const sc: SkillContext = await this.skillContextAccessor.get(dc.context, new SkillContext());
			const skillContext: SkillContext = Object.assign(new SkillContext(), sc);
			skillContext.setObj("nlpResult", responseObj);
			skillContext.setObj("currentUser", <IOnboardingState>dc.context.turnState.get("currentUser").userId);
			await this.skillContextAccessor.set(dc.context, skillContext);

			const result: DialogTurnResult = await dc.beginDialog(identifiedSkill.id);
			console.log("MainDialog.route, beginDialog(skillID) result: ", result);

			if (result.status === DialogTurnStatus.complete) {
				await this.complete(dc);
			}
		} else {
			let normalizedEntities: Array<string | undefined> | undefined =
				responseObj?.metadata?.luisResponse?.entities?.map(
					(entry: any) => entry?.resolution?.values?.[0]
				);
				
			let controlType: string | undefined = 
				normalizedEntities?.find(entry => entry === ControlSkillEnum.STOP);

			if (intent === ControlSkillEnum.INTENT && controlType) {
				await dc.context.sendActivity(GeneralReplies.Cancel_text);
				await this.analyticsSvc.saveSkillStatus(
					new SkillstatusMsg(
						dc.context,
						GeneralReplies.Cancel_text,
						CANCEL_TEXT_LABEL,
						SkillCompletionFlag.CANCEL,
						intent
					)
				);

				await this.complete(dc);
			} else {
				this.skillInstanceId = short.generate();
				dc.context.turnState.set("skillInstanceId", this.skillInstanceId);

				await dc.beginDialog(NoneIntentDialog.Name);
			}
		}
	}

	protected async onEvent(dc: DialogContext): Promise<void> {
		// logs needed for testing
		console.log(`MainDialog onEvent: Dialog Context: ${inspect(dc)}`);

		// Check if there was an action submitted from an adaptive card
		if (dc.context.activity.value) {
			// tslint:disable-next-line: no-unsafe-any
			if (dc.context.activity.value.intent) {
				if (this.settings.skills === undefined) {
					throw new Error("There is no skills in settings value");
				}

				const identifiedSkill: ISkillManifest | undefined = SkillRouter.isSkill(
					this.settings.skills,
					dc.context.activity.value.intent
				);

				if (identifiedSkill !== undefined) {
					let skillInstanceIdTemp = await this.getSkillInstanceId(dc);
					this.skillInstanceId = skillInstanceIdTemp ? skillInstanceIdTemp : this.skillInstanceId;
					dc.context.turnState.set("skillInstanceId", this.skillInstanceId);
					dc.context.turnState.set("v4Skill", identifiedSkill.id);
					await this.analyticsSvc.saveUserInput(new UserInputMsg(dc.context));
					const result: DialogTurnResult = await this.runDialog(dc, identifiedSkill.id);

					if (result.status === DialogTurnStatus.complete) {
						await this.complete(dc);
					}

					return;
				}
			}
		}

		let forward: boolean = true;
		const ev: Activity = dc.context.activity;
		if (ev.name !== undefined && ev.name.trim().length > 0) {
			switch (ev.name) {
				case TokenEvents.tokenResponseEventName: {
					forward = true;
					break;
				}
				default: {
					await dc.context.sendActivity({
						type: ActivityTypes.Trace,
						text: `"Unknown Event ${ev.name} was received but not processed."`
					});
					forward = false;
				}
			}
		}

		if (forward) {
			const result: DialogTurnResult = await dc.continueDialog();

			if (result.status === DialogTurnStatus.complete) {
				await this.complete(dc);
			}

			if (dc?.context?.activity?.value?.type?.toLowerCase().includes("cancel")) {
				await dc.cancelAllDialogs();
			}
		}
	}

	private async runDialog(dc: DialogContext, dialogId: string): Promise<DialogTurnResult> {
		let dialogResult: DialogTurnResult;

		// logs needed for testing
		console.log(`MainDialog runDialog: Dialog Id: ${dialogId}`);
		console.log(`MainDialog runDialog: Dialog Context: ${inspect(dc)}`);

		if (!dialogId) {
			throw new Error("Exception: MainDialog.runDialog(), parameter dialogId null");
		}

		if (dc?.activeDialog?.id === dialogId) {
			console.log("Continue active dialog: ", dc.activeDialog.id);
			dialogResult = await dc.continueDialog();
		} else {
			console.log("No active ", dialogId, " dialog. Starting dialog ", dialogId);
			dialogResult = await dc.beginDialog(dialogId);
		}

		// once the child dialog of MainDialog ends, it's possible that other dialogs are on stack and
		// the returned status from continueDialog or beginDialog might be DialogTurnStatus.waiting; in
		// this case, we need to return "completed" if dialogId is not on the stack anymore
		if (dialogResult.status === DialogTurnStatus.waiting && dc?.activeDialog?.id !== dialogId) {
			let dialogIdOnStack: DialogInstance | undefined = dc.stack.find((dialogInstance: DialogInstance) => {
				return dialogInstance.id === dialogId;
			});

			if (dialogIdOnStack) {
				// No dialogId found on stack, need to return "completed"
				dialogResult = { status: DialogTurnStatus.complete, result: dialogResult.result };
			}
		}

		return dialogResult;
	}

	protected async complete(dc: DialogContext, result?: DialogTurnResult): Promise<void> {
		// The active dialog's stack ended with a complete status

		// User can receive a message like "What else can I help with"

		// For SkillDialog, there's no need to end the dialog since this is done automatically in forwardToSkill called from continue/begin

		// logs needed for testing
		console.log(`MainDialog complete: Dialog Context: ${inspect(dc)}`);

		// Clean skillStatus object's nlpResult value
		const sc: SkillContext = await this.skillContextAccessor.get(dc.context, new SkillContext());
		const skillContext: SkillContext = Object.assign(new SkillContext(), sc);
		skillContext.setObj("nlpResult", <any>null);
		skillContext.setObj("variables", <any>null);
		skillContext.setObj("currentUser", <any>null);
		this.skillInstanceId = "";
		await this.skillContextAccessor.set(dc.context, skillContext);
	}

	protected async onInterruptDialog(dc: DialogContext): Promise<InterruptionAction> {
		// logs needed for testing
		console.log(`MainDialog onInterruptDialog: Dialog context: ${inspect(dc)}`);

		// Check dispatch result
		if (dc?.context?.activity?.text) {
			let responseObj = await this.skillRecognizer.recognize(dc.context);

			const intent: string = responseObj.intent;
			console.log(`MainDialog:onInterruptDialog Detected intent is ${intent} for ${dc.context.activity.text} `);

			// Mircea: Is this even needed? We can have va with no v4 skills
			if (!this.settings.skills) {
				throw new Error("There is no skills in settings value");
			}

			let identifiedSkill: ISkillManifest | undefined = SkillRouter.isSkill(this.settings.skills, intent);
			if (intent === ControlSkillEnum.INTENT && responseObj?.metadata?.luisResponse?.entities?.length > 0) {
				let normalizedEntities = responseObj.metadata.luisResponse.entities
					.map((el: any) => {
						if (el.resolution?.values?.length > 0 && el.type) {
							return { value: el.resolution.values[0], type: el.type };
						}
						return undefined;
					})
					.filter((el: any) => el);
				let appName = normalizedEntities.find((el: { type: string }) => el.type === ControlSkillEnum.APP_NAME_TAG)?.value;
				let controlType = normalizedEntities.find((el: { type: string }) => el.type === ControlSkillEnum.COMMAND_ENTITY_TAG)?.value;

				if (controlType === ControlSkillEnum.LAUNCH) {
					if (appName) {
						identifiedSkill = SkillRouter.isSkill(this.settings.skills, appName);
						await dc.cancelAllDialogs();
					}
				} else if (controlType === ControlSkillEnum.STOP) {
					identifiedSkill = undefined;
					await dc.cancelAllDialogs();
				}
			}

			// This if executes only if there is an identifiedSkill and it can happen only at the begininng of a new skill
			if (identifiedSkill && dc.stack.length < 1) {
				// We have identified a skill so initialize the skill connection with the target skill
				const sc: SkillContext = await this.skillContextAccessor.get(dc.context, new SkillContext());
				const skillContext: SkillContext = Object.assign(new SkillContext(), sc);
				const skillInstanceIdTemp = short.generate();
				this.skillInstanceId = skillInstanceIdTemp;
				skillContext.setObj("nlpResult", responseObj);
				skillContext.setObj("currentUser", <IOnboardingState>dc.context.turnState.get("currentUser").userId);
				//SkillInstanceID is need for multi-turn skills.
				skillContext.setObj("skillInstanceId", this.skillInstanceId);
				await this.skillContextAccessor.set(dc.context, skillContext);

				//Set skillInstanceId and v4Skill on the turnState as it will be used in the analyticsMiddleware
				//for capturing botOuput.
				dc.context.turnState.set("skillInstanceId", this.skillInstanceId);
				dc.context.turnState.set("v4Skill", identifiedSkill.id);
				await this.analyticsSvc.saveUserInput(new UserInputMsg(dc.context));
				await this.analyticsSvc.saveNlpData(new NlpDataMsg(dc.context, responseObj, new Date()));

				const result: DialogTurnResult = await this.runDialog(dc, identifiedSkill.id);

				if (result.status === DialogTurnStatus.complete) {
					await this.complete(dc);
				}

				return InterruptionAction.StartedDialog;
			} else if (intent === SmartIntentSkillEnum.INTENT && dc.stack.length < 1) {
				const skillInstanceIdTemp = short.generate();
				this.skillInstanceId = skillInstanceIdTemp;
				dc.context.turnState.set("skillInstanceId", this.skillInstanceId);
				dc.context.turnState.set("v4Skill", intent);

				await this.analyticsSvc.saveUserInput(new UserInputMsg(dc.context));
				await this.analyticsSvc.saveNlpData(new NlpDataMsg(dc.context, responseObj, new Date()));

				console.log("*********** skillInstanceId for SmartIntentDialog ***********"+this.skillInstanceId);
				await dc.beginDialog(SmartIntentDialog.SmartIntentDialog_DIALOG_ID, {
					data: responseObj,
					skillInstanceId: this.skillInstanceId
				});
				return InterruptionAction.StartedDialog;
			} else {
				await this.captureAnalyticsIfV4Skill(dc, this.settings.skills, responseObj);
			}
		}

		return InterruptionAction.NoAction;
	}

	//Checks if the current activeDialog is a v4 registered skill
	// and if true then captures analytics data for UserInput and NLP data
	protected async captureAnalyticsIfV4Skill(dc: DialogContext, skills: ISkillManifest[], nlpResult: any) {
		// logs needed for testing
		console.log(`MainDialog captureAnalyticsIfV4Skill: Dialog context: ${inspect(dc)}`);

		if (dc.activeDialog !== undefined) {
			const currentActiveDialog = dc.activeDialog.id;
			const currentActiveSkill = skills.some((manifest) => {
				return manifest.id === currentActiveDialog;
			});

			if (currentActiveSkill) {
				dc.context.turnState.set("v4Skill", true);
				await this.analyticsSvc.saveUserInput(new UserInputMsg(dc.context));
				await this.analyticsSvc.saveNlpData(new NlpDataMsg(dc.context, nlpResult, new Date()));
			}
		}
	}
	//Helper method to get skillIstanceId from shared state (skillContextAccessor).
	protected async getSkillInstanceId(dc: DialogContext): Promise<string> {
		const sc: SkillContext = await this.skillContextAccessor.get(dc.context, new SkillContext());
		const skillContext: SkillContext = Object.assign(new SkillContext(), sc);
		const skillInstanceId = <string>skillContext.getObj("skillInstanceId");

		return skillInstanceId;
	}

	protected async onContinueDialog(innerDc: DialogContext) {
		if (!innerDc.context.turnState.get("skillInstanceId")) {
			innerDc.context.turnState.set("skillInstanceId", this.skillInstanceId);
		}

		return super.onContinueDialog(innerDc);
	}
}

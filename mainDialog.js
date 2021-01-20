"use strict";
/**
 * Copyright(c) Microsoft Corporation.All rights reserved.
 * Licensed under the MIT License.
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.MainDialog = void 0;
const botbuilder_dialogs_1 = require("botbuilder-dialogs");
const botbuilder_skills_1 = require("botbuilder-skills");
const botbuilder_solutions_1 = require("botbuilder-solutions");
const botframework_schema_1 = require("botframework-schema");
const short_uuid_1 = __importDefault(require("short-uuid"));
const skillRecognizer_1 = require("../dispatcher/skillRecognizer");
const smartIntentDialog_1 = require("../dialogs/smartintent/smartIntentDialog");
const analyticsService_1 = require("../services/analyticsService");
const constants_1 = require("../config/constants");
const noneIntentDialog_1 = require("./noneIntentDialog");
const util_1 = require("util");
class MainDialog extends botbuilder_solutions_1.RouterDialog {
    // Constructor
    constructor(settings, services, cancelDialog, skillDialogs, skillContextAccessor, onboardingAccessor, telemetryClient, analyticsSvc) {
        super(MainDialog.name, telemetryClient);
        this.settings = settings;
        this.skillContextAccessor = skillContextAccessor;
        this.telemetryClient = telemetryClient;
        this.addDialog(cancelDialog);
        skillDialogs.forEach((skillDialog) => {
            this.addDialog(skillDialog);
        });
        this.skillRecognizer = new skillRecognizer_1.SkillRecognizer(false, this.telemetryClient);
        this.analyticsSvc = analyticsSvc;
        this.skillInstanceId = "";
    }
    async route(dc) {
        var _a, _b, _c, _d;
        // logs needed for testing
        console.log(`MainDialog route: Dialog Context: ${util_1.inspect(dc)}`);
        // Check dispatch result
        let responseObj = await this.skillRecognizer.recognize(dc.context);
        const intent = responseObj.intent;
        console.log(`MainDialog:route Detected intent is ${intent} for ${dc.context.activity.text} `);
        if (!((_a = this.settings) === null || _a === void 0 ? void 0 : _a.skills)) {
            throw new Error("There is no skills in settings value");
        }
        // Identify if the dispatch intent matches any Action within a Skill if so, we pass to the appropriate SkillDialog to hand-off
        const identifiedSkill = botbuilder_skills_1.SkillRouter.isSkill(this.settings.skills, intent);
        if (identifiedSkill) {
            // We have identified a skill so initialize the skill connection with the target skill
            const sc = await this.skillContextAccessor.get(dc.context, new botbuilder_skills_1.SkillContext());
            const skillContext = Object.assign(new botbuilder_skills_1.SkillContext(), sc);
            skillContext.setObj("nlpResult", responseObj);
            skillContext.setObj("currentUser", dc.context.turnState.get("currentUser").userId);
            await this.skillContextAccessor.set(dc.context, skillContext);
            const result = await dc.beginDialog(identifiedSkill.id);
            console.log("MainDialog.route, beginDialog(skillID) result: ", result);
            if (result.status === botbuilder_dialogs_1.DialogTurnStatus.complete) {
                await this.complete(dc);
            }
        }
        else {
            let normalizedEntities = (_d = (_c = (_b = responseObj === null || responseObj === void 0 ? void 0 : responseObj.metadata) === null || _b === void 0 ? void 0 : _b.luisResponse) === null || _c === void 0 ? void 0 : _c.entities) === null || _d === void 0 ? void 0 : _d.map((entry) => { var _a, _b; return (_b = (_a = entry === null || entry === void 0 ? void 0 : entry.resolution) === null || _a === void 0 ? void 0 : _a.values) === null || _b === void 0 ? void 0 : _b[0]; });
            let controlType = normalizedEntities === null || normalizedEntities === void 0 ? void 0 : normalizedEntities.find(entry => entry === constants_1.ControlSkillEnum.STOP);
            if (intent === constants_1.ControlSkillEnum.INTENT && controlType) {
                await dc.context.sendActivity("Okay, your request has been cancelled. Let me know if you need anything else." /* Cancel_text */);
                await this.analyticsSvc.saveSkillStatus(new analyticsService_1.SkillstatusMsg(dc.context, "Okay, your request has been cancelled. Let me know if you need anything else." /* Cancel_text */, constants_1.CANCEL_TEXT_LABEL, constants_1.SkillCompletionFlag.CANCEL, intent));
                await this.complete(dc);
            }
            else {
                this.skillInstanceId = short_uuid_1.default.generate();
                dc.context.turnState.set("skillInstanceId", this.skillInstanceId);
                await dc.beginDialog(noneIntentDialog_1.NoneIntentDialog.Name);
            }
        }
    }
    async onEvent(dc) {
        var _a, _b, _c, _d;
        // logs needed for testing
        console.log(`MainDialog onEvent: Dialog Context: ${util_1.inspect(dc)}`);
        // Check if there was an action submitted from an adaptive card
        if (dc.context.activity.value) {
            // tslint:disable-next-line: no-unsafe-any
            if (dc.context.activity.value.intent) {
                if (this.settings.skills === undefined) {
                    throw new Error("There is no skills in settings value");
                }
                const identifiedSkill = botbuilder_skills_1.SkillRouter.isSkill(this.settings.skills, dc.context.activity.value.intent);
                if (identifiedSkill !== undefined) {
                    let skillInstanceIdTemp = await this.getSkillInstanceId(dc);
                    this.skillInstanceId = skillInstanceIdTemp ? skillInstanceIdTemp : this.skillInstanceId;
                    dc.context.turnState.set("skillInstanceId", this.skillInstanceId);
                    dc.context.turnState.set("v4Skill", identifiedSkill.id);
                    await this.analyticsSvc.saveUserInput(new analyticsService_1.UserInputMsg(dc.context));
                    const result = await this.runDialog(dc, identifiedSkill.id);
                    if (result.status === botbuilder_dialogs_1.DialogTurnStatus.complete) {
                        await this.complete(dc);
                    }
                    return;
                }
            }
        }
        let forward = true;
        const ev = dc.context.activity;
        if (ev.name !== undefined && ev.name.trim().length > 0) {
            switch (ev.name) {
                case botbuilder_solutions_1.TokenEvents.tokenResponseEventName: {
                    forward = true;
                    break;
                }
                default: {
                    await dc.context.sendActivity({
                        type: botframework_schema_1.ActivityTypes.Trace,
                        text: `"Unknown Event ${ev.name} was received but not processed."`
                    });
                    forward = false;
                }
            }
        }
        if (forward) {
            const result = await dc.continueDialog();
            if (result.status === botbuilder_dialogs_1.DialogTurnStatus.complete) {
                await this.complete(dc);
            }
            if ((_d = (_c = (_b = (_a = dc === null || dc === void 0 ? void 0 : dc.context) === null || _a === void 0 ? void 0 : _a.activity) === null || _b === void 0 ? void 0 : _b.value) === null || _c === void 0 ? void 0 : _c.type) === null || _d === void 0 ? void 0 : _d.toLowerCase().includes("cancel")) {
                await dc.cancelAllDialogs();
            }
        }
    }
    async runDialog(dc, dialogId) {
        var _a, _b;
        let dialogResult;
        // logs needed for testing
        console.log(`MainDialog runDialog: Dialog Id: ${dialogId}`);
        console.log(`MainDialog runDialog: Dialog Context: ${util_1.inspect(dc)}`);
        if (!dialogId) {
            throw new Error("Exception: MainDialog.runDialog(), parameter dialogId null");
        }
        if (((_a = dc === null || dc === void 0 ? void 0 : dc.activeDialog) === null || _a === void 0 ? void 0 : _a.id) === dialogId) {
            console.log("Continue active dialog: ", dc.activeDialog.id);
            dialogResult = await dc.continueDialog();
        }
        else {
            console.log("No active ", dialogId, " dialog. Starting dialog ", dialogId);
            dialogResult = await dc.beginDialog(dialogId);
        }
        // once the child dialog of MainDialog ends, it's possible that other dialogs are on stack and
        // the returned status from continueDialog or beginDialog might be DialogTurnStatus.waiting; in
        // this case, we need to return "completed" if dialogId is not on the stack anymore
        if (dialogResult.status === botbuilder_dialogs_1.DialogTurnStatus.waiting && ((_b = dc === null || dc === void 0 ? void 0 : dc.activeDialog) === null || _b === void 0 ? void 0 : _b.id) !== dialogId) {
            let dialogIdOnStack = dc.stack.find((dialogInstance) => {
                return dialogInstance.id === dialogId;
            });
            if (dialogIdOnStack) {
                // No dialogId found on stack, need to return "completed"
                dialogResult = { status: botbuilder_dialogs_1.DialogTurnStatus.complete, result: dialogResult.result };
            }
        }
        return dialogResult;
    }
    async complete(dc, result) {
        // The active dialog's stack ended with a complete status
        // User can receive a message like "What else can I help with"
        // For SkillDialog, there's no need to end the dialog since this is done automatically in forwardToSkill called from continue/begin
        // logs needed for testing
        console.log(`MainDialog complete: Dialog Context: ${util_1.inspect(dc)}`);
        // Clean skillStatus object's nlpResult value
        const sc = await this.skillContextAccessor.get(dc.context, new botbuilder_skills_1.SkillContext());
        const skillContext = Object.assign(new botbuilder_skills_1.SkillContext(), sc);
        skillContext.setObj("nlpResult", null);
        skillContext.setObj("variables", null);
        skillContext.setObj("currentUser", null);
        this.skillInstanceId = "";
        await this.skillContextAccessor.set(dc.context, skillContext);
    }
    async onInterruptDialog(dc) {
        var _a, _b, _c, _d, _e, _f, _g;
        // logs needed for testing
        console.log(`MainDialog onInterruptDialog: Dialog context: ${util_1.inspect(dc)}`);
        // Check dispatch result
        if ((_b = (_a = dc === null || dc === void 0 ? void 0 : dc.context) === null || _a === void 0 ? void 0 : _a.activity) === null || _b === void 0 ? void 0 : _b.text) {
            let responseObj = await this.skillRecognizer.recognize(dc.context);
            const intent = responseObj.intent;
            console.log(`MainDialog:onInterruptDialog Detected intent is ${intent} for ${dc.context.activity.text} `);
            // Mircea: Is this even needed? We can have va with no v4 skills
            if (!this.settings.skills) {
                throw new Error("There is no skills in settings value");
            }
            let identifiedSkill = botbuilder_skills_1.SkillRouter.isSkill(this.settings.skills, intent);
            if (intent === constants_1.ControlSkillEnum.INTENT && ((_e = (_d = (_c = responseObj === null || responseObj === void 0 ? void 0 : responseObj.metadata) === null || _c === void 0 ? void 0 : _c.luisResponse) === null || _d === void 0 ? void 0 : _d.entities) === null || _e === void 0 ? void 0 : _e.length) > 0) {
                let normalizedEntities = responseObj.metadata.luisResponse.entities
                    .map((el) => {
                    var _a, _b;
                    if (((_b = (_a = el.resolution) === null || _a === void 0 ? void 0 : _a.values) === null || _b === void 0 ? void 0 : _b.length) > 0 && el.type) {
                        return { value: el.resolution.values[0], type: el.type };
                    }
                    return undefined;
                })
                    .filter((el) => el);
                let appName = (_f = normalizedEntities.find((el) => el.type === constants_1.ControlSkillEnum.APP_NAME_TAG)) === null || _f === void 0 ? void 0 : _f.value;
                let controlType = (_g = normalizedEntities.find((el) => el.type === constants_1.ControlSkillEnum.COMMAND_ENTITY_TAG)) === null || _g === void 0 ? void 0 : _g.value;
                if (controlType === constants_1.ControlSkillEnum.LAUNCH) {
                    if (appName) {
                        identifiedSkill = botbuilder_skills_1.SkillRouter.isSkill(this.settings.skills, appName);
                        await dc.cancelAllDialogs();
                    }
                }
                else if (controlType === constants_1.ControlSkillEnum.STOP) {
                    identifiedSkill = undefined;
                    await dc.cancelAllDialogs();
                }
            }
            // This if executes only if there is an identifiedSkill and it can happen only at the begininng of a new skill
            if (identifiedSkill && dc.stack.length < 1) {
                // We have identified a skill so initialize the skill connection with the target skill
                const sc = await this.skillContextAccessor.get(dc.context, new botbuilder_skills_1.SkillContext());
                const skillContext = Object.assign(new botbuilder_skills_1.SkillContext(), sc);
                const skillInstanceIdTemp = short_uuid_1.default.generate();
                this.skillInstanceId = skillInstanceIdTemp;
                skillContext.setObj("nlpResult", responseObj);
                skillContext.setObj("currentUser", dc.context.turnState.get("currentUser").userId);
                //SkillInstanceID is need for multi-turn skills.
                skillContext.setObj("skillInstanceId", this.skillInstanceId);
                await this.skillContextAccessor.set(dc.context, skillContext);
                //Set skillInstanceId and v4Skill on the turnState as it will be used in the analyticsMiddleware
                //for capturing botOuput.
                dc.context.turnState.set("skillInstanceId", this.skillInstanceId);
                dc.context.turnState.set("v4Skill", identifiedSkill.id);
                await this.analyticsSvc.saveUserInput(new analyticsService_1.UserInputMsg(dc.context));
                await this.analyticsSvc.saveNlpData(new analyticsService_1.NlpDataMsg(dc.context, responseObj, new Date()));
                const result = await this.runDialog(dc, identifiedSkill.id);
                if (result.status === botbuilder_dialogs_1.DialogTurnStatus.complete) {
                    await this.complete(dc);
                }
                return botbuilder_solutions_1.InterruptionAction.StartedDialog;
            }
            else if (intent === constants_1.SmartIntentSkillEnum.INTENT && dc.stack.length < 1) {
                const skillInstanceIdTemp = short_uuid_1.default.generate();
                this.skillInstanceId = skillInstanceIdTemp;
                dc.context.turnState.set("skillInstanceId", this.skillInstanceId);
                dc.context.turnState.set("v4Skill", intent);
                await this.analyticsSvc.saveUserInput(new analyticsService_1.UserInputMsg(dc.context));
                await this.analyticsSvc.saveNlpData(new analyticsService_1.NlpDataMsg(dc.context, responseObj, new Date()));
                console.log("*********** skillInstanceId for SmartIntentDialog ***********" + this.skillInstanceId);
                await dc.beginDialog(smartIntentDialog_1.SmartIntentDialog.SmartIntentDialog_DIALOG_ID, {
                    data: responseObj,
                    skillInstanceId: this.skillInstanceId
                });
                return botbuilder_solutions_1.InterruptionAction.StartedDialog;
            }
            else {
                await this.captureAnalyticsIfV4Skill(dc, this.settings.skills, responseObj);
            }
        }
        return botbuilder_solutions_1.InterruptionAction.NoAction;
    }
    //Checks if the current activeDialog is a v4 registered skill
    // and if true then captures analytics data for UserInput and NLP data
    async captureAnalyticsIfV4Skill(dc, skills, nlpResult) {
        // logs needed for testing
        console.log(`MainDialog captureAnalyticsIfV4Skill: Dialog context: ${util_1.inspect(dc)}`);
        if (dc.activeDialog !== undefined) {
            const currentActiveDialog = dc.activeDialog.id;
            const currentActiveSkill = skills.some((manifest) => {
                return manifest.id === currentActiveDialog;
            });
            if (currentActiveSkill) {
                dc.context.turnState.set("v4Skill", true);
                await this.analyticsSvc.saveUserInput(new analyticsService_1.UserInputMsg(dc.context));
                await this.analyticsSvc.saveNlpData(new analyticsService_1.NlpDataMsg(dc.context, nlpResult, new Date()));
            }
        }
    }
    //Helper method to get skillIstanceId from shared state (skillContextAccessor).
    async getSkillInstanceId(dc) {
        const sc = await this.skillContextAccessor.get(dc.context, new botbuilder_skills_1.SkillContext());
        const skillContext = Object.assign(new botbuilder_skills_1.SkillContext(), sc);
        const skillInstanceId = skillContext.getObj("skillInstanceId");
        return skillInstanceId;
    }
    async onContinueDialog(innerDc) {
        if (!innerDc.context.turnState.get("skillInstanceId")) {
            innerDc.context.turnState.set("skillInstanceId", this.skillInstanceId);
        }
        return super.onContinueDialog(innerDc);
    }
}
exports.MainDialog = MainDialog;
//# sourceMappingURL=mainDialog.js.map
import { Application, AI, PredictedSayCommand } from '@microsoft/teams-ai';
import { ActivityTypes, TurnContext } from 'botbuilder';
import { ApplicationTurnState } from './prepareApp';


interface PauseParameters {
    time: number;
}
/**
 *
 * @param {Application} app Application to add the response formatter to.
 */
export function preparePromtTesteActions(app: Application<ApplicationTurnState>): void {
    // Register action handlers
    app.ai.action('LightsOn', async (context: TurnContext, state: ApplicationTurnState) => {
        state.conversation.lightsOn = true;
        await context.sendActivity(`[lights on]`);
        return `the lights are now on`;
    });
    
    app.ai.action('LightsOff', async (context: TurnContext, state: ApplicationTurnState) => {
        state.conversation.lightsOn = false;
        await context.sendActivity(`[lights off]`);
        return `the lights are now off`;
    });
    
    app.ai.action('Pause', async (context: TurnContext, state: ApplicationTurnState, parameters: PauseParameters) => {
        await context.sendActivity(`[pausing for ${parameters.time / 1000} seconds]`);
        await new Promise((resolve) => setTimeout(resolve, parameters.time));
        return `done pausing`;
    });
}

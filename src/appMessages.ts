import { Application, AI, PredictedSayCommand, TurnState } from '@microsoft/teams-ai';
import { ActivityTypes, MessagingExtensionAttachment, TurnContext } from 'botbuilder';
import { ApplicationTurnState } from './prepareApp';
import axios from 'axios';



interface HtmlDataParameters {
    data:any;
}

/**
 *
 * @param {Application} app Application to add the response formatter to.
 */
export function addAppMessages(app: Application<ApplicationTurnState>): void {
    
    app.message('/sql', async (context: TurnContext, state: ApplicationTurnState) => {
        await context.sendActivities([
            { type: ActivityTypes.Typing },
            { type: 'delay', value: 3000 },
            { type: ActivityTypes.Message, text: 'Finished typing' }
        ]);
        await context.sendActivity(`sql command running....`);
    });
    
    // Listen for search actions
    app.messageExtensions.query('searchCmd', async (context: TurnContext, state: TurnState, query) => {
        console.log("searchcm - here" + query);
        const searchQuery = query.parameters.searchQuery?? '';
        const count = query.count ?? 10;
        const response = await axios.get(
            `http://registry.npmjs.com/-/v1/search?${new URLSearchParams({
                size: count.toString(),
                text: searchQuery
            }).toString()}`
        );
    
        // Format search results
        const results: MessagingExtensionAttachment[] = [];
        response?.data?.objects?.forEach((obj: any) => results.push(obj.package));
    
        // Return results as a list
        return {
            attachmentLayout: 'list',
            attachments: results,
            type: 'result'
        };
    });
    
    // Listen for item tap
    app.messageExtensions.selectItem(async (context: TurnContext, state: TurnState, item) => {
        // Generate detailed result
        //const card = createNpmPackageCard(item);
    
        // Return results
        return {
            attachmentLayout: 'list',
            attachments: [item.description],
            type: 'result'
        };
    });
}

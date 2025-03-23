import { Application, AI, PredictedSayCommand } from '@microsoft/teams-ai';
import { ActivityTypes, TurnContext } from 'botbuilder';
import { ApplicationTurnState } from './prepareApp';
import { AzureOpenAIClient, OpenAIClient } from '@microsoft/teams-ai/lib/internals';

interface SqlScriptParameters {
    sqlServerName: string;
    databaseName: string;
    query:string;
}

/**
 *
 * @param {Application} app Application to add the response formatter to.
 */
export function prepareSqlActions(app: Application<ApplicationTurnState>): void {
    
    app.ai.action('sqlscript', async (context: TurnContext, state: ApplicationTurnState, parameters: SqlScriptParameters) => {
        if(parameters.sqlServerName == "")
        {
            await context.sendActivity("PLease write sql server name");
        }
    
        if(parameters.databaseName == "")
            {
                await context.sendActivity("PLease write database name");
            }
        //await context.sendActivity(`Sql query ${parameters.sqlServerName} and database: ${parameters.databaseName}`);
        //await new Promise((resolve) => setTimeout(resolve, parameters.time));
        state.setValue("sqlScript", "executed");
    
        await context.sendActivity({
            type: ActivityTypes.Message,
            text: `Sql Execution: ${parameters.sqlServerName} and database: ${parameters.databaseName}`,
            entities: [
              {
               type: "https://schema.org/Message",
               "@type": "Message",
               "@context": "https://schema.org",
              }
            ]
          });
    
          await context.sendActivity({
            type: ActivityTypes.Message,
            text: `Query ${parameters.query}`,
            entities: [
              {
               type: "https://schema.org/Message",
               "@type": "Message",
               "@context": "https://schema.org",
              }
            ]
          });
    
          const data = [
               { id:1, name:"a1", email:'e1'},
               { id:1, name:"a1", email:'e1'},
               { id:1, name:"a1", email:'e1'},
               { id:1, name:"a1", email:'e1'},
               { id:1, name:"a1", email:'e1'},
               { id:1, name:"a1", email:'e1'},
               { id:1, name:"a1", email:'e1'},
               { id:1, name:"a1", email:'e1'}
          ]
    
          state.conversation.data = data;
          await context.sendActivity("data:"+ JSON.stringify(data));
          //const prompt = "tranfer below data in html format:"+ JSON.stringify(data);
          //const response = await planner.completePrompt(context, state, "query");
          //await context.sendActivity("response:"+ response);
    
          context.turnState.set("data", data);

          /*
          const client = new OpenAIClient({
            endpoint:process.env.AzureOpenAIEndpoint,
            apiKey:process.env.AzureOpenAIApiKey ?? ""
          }
        );

        let message = "Please format below data to required format:<br/>"+ JSON.stringify(data);
        
        //state.temp.input = message;
        //let output= await app.ai.planner.beginTask(context, state, app.ai);
*/
          return "";
    });
   
}

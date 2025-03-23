import { Application, AI, PredictedSayCommand } from '@microsoft/teams-ai';
import { ActivityTypes, TurnContext } from 'botbuilder';
import { ApplicationTurnState } from './prepareApp';




interface HtmlDataParameters {
    data:any;
    email:any;
    emailSubject:any;
}

/**
 *
 * @param {Application} app Application to add the response formatter to.
 */
export function prepareTransformationActions(app: Application<ApplicationTurnState>): void {
    // Register action handlers

    const prepareDataInHtmlTable =  async(context: TurnContext, state: ApplicationTurnState, data:any[]) =>
    {
        await context.sendActivity("Data processing in html format.....");
    
        if (data == null) {
            await context.sendActivity('Invalid data');
            return;
        }
    
        if( !Array.isArray(data))
        {
            data = state.conversation.data;
        }
    
        if( data == undefined)
        {
            await context.sendActivity("No data, please request again for a query");
            return "";
        }
    
        let htmlTable = '<table border="1"><tr>';
    
        // Add table headers
        Object.keys(data[0]).forEach(key => {
            htmlTable += `<th>${key}</th>`;
        });
        htmlTable += '</tr>';
    
        // Add table rows
        data.forEach(row => {
            htmlTable += '<tr>';
            Object.values(row).forEach(value => {
                htmlTable += `<td>${value}</td>`;
            });
            htmlTable += '</tr>';
        });
    
        htmlTable += '</table>';
    
        await context.sendActivity(htmlTable);
    
        return "";
    }
    
    app.ai.action("prepareDataInHtmlTable", async (context: TurnContext, state: ApplicationTurnState, parameters: HtmlDataParameters) =>
    {
          await prepareDataInHtmlTable(context,state,parameters.data);
          return "";
    });

    app.ai.action("formatResultXml", async (context: TurnContext, state: ApplicationTurnState, parameters: HtmlDataParameters) =>
        {
              await context.sendActivity("<xml>...</xml>");
              return "";
        });

        app.ai.action("formatResultPdf", async (context: TurnContext, state: ApplicationTurnState, parameters: HtmlDataParameters) =>
            {
                  await context.sendActivity("Pdf for data was created on <a href='https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf'>generated_document.pdf</a>");
                  return "";
            });

            app.ai.action("sendDataByEmail", async (context: TurnContext, state: ApplicationTurnState, parameters: HtmlDataParameters) =>
                {
                    if( parameters.email == undefined || parameters.email == "<recipient-email-address>")
                    {
                        await context.sendActivity("Missing email, what is your email address");
                        return "Missing email address";
                    }


                      await context.sendActivity("Sending data to email:"+ parameters.email +" and data:"+ parameters.data);
                      return "";
                });
}

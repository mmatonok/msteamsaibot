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

    const prepareDataInHtmlTable =  async(context: TurnContext, state: ApplicationTurnState, parameters: any) =>
    {
        await context.sendActivity("Data processing in html format.....");
    /*
        let data = parameters.data;
        if (data == null) {
            data = parameters;
            if( data == null )
            {
                await context.sendActivity('Invalid data');
                return;
            }
        }
    
        let jsonArray = [];
        try{
        jsonArray = JSON.parse(data) as any[];
        if( !Array.isArray(jsonArray))
        {
            data = state.conversation.data;
        }
    }catch(e)
    {
        data = state.conversation.data;
        jsonArray = JSON.parse(data) as any[];
    }
    
        if( jsonArray == undefined)
        {
            await context.sendActivity("No data, please request again for a query:", parameters);
            return "";
        }
    */
        const jsonArray = [
            { id:1, case:"a1", email:'e1', title:'case1', descption:'case company1', 'path': 'https://a.com'},
            { id:2, case:"a11", email:'e1', title:'case11', descption:'case company1', 'path': 'https://a.com'},
            { id:3, case:"a111", email:'e2', title:'case12', descption:'case company1', 'path': 'https://a.com'},
            { id:4, case:"a1111", email:'e2', title:'case122', descption:'case company1', 'path': 'https://a.com'},
            { id:5, case:"a12", email:'e3', title:'case13', descption:'case company1', 'path': 'https://a.com'},
            { id:6, case:"a122", email:'e3', title:'case13', descption:'case company1', 'path': 'https://a.com'},
       ]


        let htmlTable = '<h1 style="color:orange">Summary:</h1><br/><table border="1"><tr>';
    
        // Add table headers
        Object.keys(jsonArray[0]).forEach(key => {
            htmlTable += `<th>${key}</th>`;
        });
        htmlTable += '</tr>';
    
        // Add table rows
        jsonArray.forEach(row => {
            htmlTable += '<tr>';
            Object.values(row).forEach(value => {
                htmlTable += `<td>${value}</td>`;
            });
            htmlTable += '</tr>';
        });
    
        htmlTable += '</table>';
    
        await context.sendActivity(htmlTable);

        await context.sendActivity("Anything else?");
    
        return AI.StopCommandName;
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

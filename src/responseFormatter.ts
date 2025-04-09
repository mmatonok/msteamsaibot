import { Application, AI, PredictedSayCommand, Citation } from '@microsoft/teams-ai';
import { ActivityTypes } from 'botbuilder';


/**
 *
 * @param {Application} app Application to add the response formatter to.
 */
export function addResponseFormatter(app: Application): void {

    app.ai.action<PredictedSayCommand>(AI.SayCommandActionName, async (context, state, data:PredictedSayCommand) => {
        // Replace markdown code blocks with <pre> tags
        let addTag = false;
        let inCodeBlock = false;
        const output: string[] = [];
        const response = data.response.content!.split('\n');

        for (const line of response) {
            if (line.startsWith('```')) {
                if (!inCodeBlock) {
                    // Add tag to start of next line
                    addTag = true;
                    inCodeBlock = true;
                } else {
                    // Add tag to end of previous line
                    output[output.length - 1] += '</pre>';
                    addTag = false;
                    inCodeBlock = false;
                }
            } else if (addTag) {
                output.push(`<pre>${line}`);
                addTag = false;
            } else {
                output.push(line);
            }
        }
        // Send response
        const formattedResponse = output.join('\n');

        //console.log("asnwer:", formattedResponse);
        

        var ret = state.getValue("sqlScript");

        if(ret != "executed")
        {
         await context.sendActivity({
                    type: ActivityTypes.Message,
                    text: formattedResponse,
                    /*
                    entities: [
                        {
                         type: "https://schema.org/Message",
                         "@type": "Message",
                         "@context": "https://schema.org",
                         additionalType: ["AIGeneratedContent"], // Enables AI label
                        }
                      ],       
                      */         
                    entities: [
                        {
                          type: "https://schema.org/Message",
                          "@type": "Message",
                          "@context": "https://schema.org",
                          additionalType: ["AIGeneratedContent"],
                          citation: [
                          {
                            "@type": "Claim",
                            position: 1, // Required. Must match the [1] in the text above
                            appearance: {
                              "@type": "DigitalDocument",
                              name: "AI bot", // Title
                              url: "https://example.com/claim-1", // Hyperlink on the title
                              abstract: "Excerpt description", // Appears in the citation pop-up window
                              text: "{\"type\":\"AdaptiveCard\",\"$schema\":\"http://adaptivecards.io/schemas/adaptive-card.json\",\"version\":\"1.6\",\"body\":[{\"type\":\"TextBlock\",\"text\":\"Adaptive Card text\"}]}", // Appears as a stringified Adaptive Card
                              keywords: ["keyword 1", "keyword 2", "keyword 3"], // Appears in the citation pop-up window
                              encodingFormat: "application/vnd.microsoft.card.adaptive",
                              image: {
                                "@type": "ImageObject",
                                name: "Microsoft Word"
                              },
                             },
                          },
                        ]
                    },
                ],
                  });

            state.setValue("sqlScript", "");
        }
        else{
            console.log("No AI response");
        }
        
        return "";
    });
}

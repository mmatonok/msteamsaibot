// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Import required packages
import * as path from 'path';
import * as restify from 'restify';
import axios from 'axios';

//const fs = require('fs');
// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { ActivityTypes, ConfigurationServiceClientCredentialFactory, MemoryStorage, MessagingExtensionAttachment, TurnContext } from 'botbuilder';

import {
    ActionPlanner,
    AI,
    Application,
    DefaultConversationState,
    Memory,
    OpenAIModel,
    PredictedSayCommand,
    PromptManager,
    TeamsAdapter,
    TurnState
} from '@microsoft/teams-ai';

import { addResponseFormatter } from './responseFormatter';
import { VectraDataSource } from './VectraDataSource';

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
const adapter = new TeamsAdapter(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppType: process.env.BOT_TYPE,
        MicrosoftAppId: process.env.BOT_ID,
        MicrosoftAppPassword: process.env.BOT_PASSWORD,
        MicrosoftAppTenantId: process.env.BOT_TENANT_ID
    })
);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: any) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error}`);
    console.log(error);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
    console.log('\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator');
    console.log('\nTo test your bot in Teams, sideload the app manifest.json within Teams Apps.');
});

// eslint-disable-next-line @typescript-eslint/no-empty-interface
interface ConversationState extends DefaultConversationState {
    lightsOn: boolean;
    ai?:AI;
}

type ApplicationTurnState = TurnState<ConversationState>;

if (!process.env.OPENAI_KEY && !process.env.AZURE_OPENAI_KEY) {
    throw new Error('Missing environment variables - please check that OPENAI_KEY or AZURE_OPENAI_KEY is set.');
}

// Create AI components
const model = new OpenAIModel({
    // OpenAI Support
    apiKey: process.env.OPENAI_KEY!,
    defaultModel: 'gpt-4o',
    project: process.env.OPENAI_PROJECT_KEY!,
    clientOptions: { apiKey: process.env.OPENAI_KEY! },

    // Azure OpenAI Support
    azureApiKey: process.env.AZURE_OPENAI_KEY!,
    azureDefaultDeployment: 'gpt-4o',
    azureEndpoint: process.env.AZURE_OPENAI_ENDPOINT!,
    azureApiVersion: '2023-03-15-preview',

    // Request logging
    logRequests: true
});

const prompts = new PromptManager({
    promptsFolder: path.join(__dirname, '../src/prompts')
});

const planner = new ActionPlanner<ApplicationTurnState>({
    model,
    prompts,
    defaultPrompt: 'default'
});

// Define a prompt function for getting the current status of the lights
planner.prompts.addFunction('getLightStatus', async (context: TurnContext, memory: Memory) => {
    return memory.getValue('conversation.lightsOn') ? 'on' : 'off';
});

// Define storage and application
const storage = new MemoryStorage();
const app = new Application<ApplicationTurnState>({
    
    storage,
    ai: {
        planner
    }
});

/*
withAuthentication(adapter, {
    settings: { },
    autoSignIn: (context: TurnContext) => {
        const signOutActivity = context.activity?.value.commandId === 'signOutCommand';
        if (signOutActivity) {
            return Promise.resolve(false);
        }

        return Promise.resolve(true);
    }
})
*/

// Register your data source with planner
planner.prompts.addDataSource(
    new VectraDataSource({
        name: 'customdata',
        apiKey: process.env.OPENAI_KEY!,
        azureApiKey: process.env.AZURE_OPENAI_KEY!,
        azureEndpoint: process.env.AZURE_OPENAI_ENDPOINT!,
        indexFolder: path.join(__dirname, '../index')
    })
);

// Add a custom response formatter to convert markdown code blocks to <pre> tags
addResponseFormatter(app);


// Listen for new members to join the conversation
app.conversationUpdate('membersAdded', async (context: TurnContext, state: ApplicationTurnState) => {
    
        await context.sendActivity("ollaaaaa....");
    
});

// Register other AI actions
app.ai.action(
    AI.FlaggedInputActionName,
    async (context: TurnContext, state: ApplicationTurnState, data: Record<string, any>) => {
        await context.sendActivity(`I'm sorry your message was flagged: ${JSON.stringify(data)}`);
        return AI.StopCommandName;
    }
);

app.ai.action(AI.FlaggedOutputActionName, async (context: TurnContext, state: ApplicationTurnState, data: any) => {
    await context.sendActivity(`I'm not allowed to talk about such things.`);
    return AI.StopCommandName;
});

/*
//action are executed before text response is shown

app.ai.action("restartService", async (context: TurnContext, state: ApplicationTurnState) => {
    console.log('restartService');
    await context.sendActivity("Service restarted");
    return "";
    });

    //get data from database
    interface reportParameters {
        reportName: string
    }
    
    app.ai.action('get data from database', async (context: TurnContext, state: ApplicationTurnState, parameters: reportParameters) => {
        await context.sendActivity(`Preparing report ${parameters.reportName}...`);
        await new Promise((resolve) => setTimeout(resolve, 5000));
        let activity = await context.sendActivity(`Report ${parameters.reportName} ready:`);
        
        await sendFileCard(context);
        
        return `done`;
    });

    async function sendFileCard(context:TurnContext) {
        const filename = 'teams-logo.png';
        //const stats = fs.statSync(path.join('files', filename));
        //const stats = fs.
        const fileSize = 100;//stats.size;

        const consentContext = { filename: filename };
        const fileCard = {
            description: 'This is a report',
            sizeInBytes: fileSize,
            acceptContext: consentContext,
            declineContext: consentContext
        };

        const asAttachment = {
            content: fileCard,
            contentType: 'application/vnd.microsoft.teams.card.file.consent',
            name: filename
        };
        await context.sendActivity({ attachments: [asAttachment] });
    }
*/
    

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

interface PauseParameters {
    time: number;
}

app.ai.action('Pause', async (context: TurnContext, state: ApplicationTurnState, parameters: PauseParameters) => {
    await context.sendActivity(`[pausing for ${parameters.time / 1000} seconds]`);
    await new Promise((resolve) => setTimeout(resolve, parameters.time));
    return `done pausing`;
});

interface SqlScriptParameters {
    sqlServerName: string;
    databaseName: string;
}

app.ai.action('sqlscript', async (context: TurnContext, state: ApplicationTurnState, parameters: SqlScriptParameters) => {
    if(parameters.sqlServerName == "")
    {
        await context.sendActivity("PLease write sql server name");
    }

    if(parameters.databaseName == "")
        {
            await context.sendActivity("PLease write database name");
        }
    await context.sendActivity(`Sql query ${parameters.sqlServerName} and database: ${parameters.databaseName}`);
    //await new Promise((resolve) => setTimeout(resolve, parameters.time));
    return "execute sql script....";
});


// Listen for user to say '/reset' and then delete conversation state
app.message('/reset', async (context: TurnContext, state: ApplicationTurnState) => {
    state.deleteConversationState();
    await context.sendActivity(`Ok I've deleted the current conversation state.`);
});


// Listen for user to say '/reset' and then delete conversation state
app.message('/login', async (context: TurnContext, state: ApplicationTurnState) => {
    const token = await app.getTokenOrStartSignIn(context, state, "graph");
    if (!token) {
    await context.sendActivity("You have to be signed in to fulfill this request. Starting sign in flow...");
    }
});


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

// Listen for incoming server requests.
server.post('/api/messages', async (req, res) => {
    // Route received a request to adapter for processing
    await adapter.process(req, res as any, async (context) => {
        // Dispatch to application for routing
        await app.run(context);
    });
});

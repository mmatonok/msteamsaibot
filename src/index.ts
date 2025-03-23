// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Import required packages
import * as path from 'path';
import * as restify from 'restify';

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
    PromptCompletionModelResponseReceivedEvent,
    PromptManager,
    TeamsAdapter,
    TurnState
} from '@microsoft/teams-ai';

import { addResponseFormatter } from './responseFormatter';
import { VectraDataSource } from './VectraDataSource';
import { ApplicationTurnState } from './prepareApp';
import { prepareSqlActions } from './prepreSqlActins';
import { prepareTransformationActions } from './prepreTranformationy';
import { addAppMessages } from './appMessages';

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
    logRequests: true,
    stream: false
});

const endStreamHandler: PromptCompletionModelResponseReceivedEvent = (ctx, memory, response, streamer) => {
    // ... Setup attachments
    //streamer.setAttachments([...cards]);                      // Set attachments
};


const prompts = new PromptManager({
    promptsFolder: path.join(__dirname, '../src/prompts')
});

const planner = new ActionPlanner<ApplicationTurnState>({
    model,
    prompts,
    defaultPrompt: 'default',
    startStreamingMessage: 'Loading stream results...', // Set informative message
    endStreamHandler: endStreamHandler                  // Set final chunk handler
});

// Define a prompt function for getting the current status of the lights
planner.prompts.addFunction('getLightStatus', async (context: TurnContext, memory: Memory) => {
    return memory.getValue('conversation.lightsOn') ? 'on' : 'off';
});

planner.prompts.addFunction('getData', async (context: TurnContext, memory: Memory) => {
    return memory.getValue('conversation.data');
});

planner.prompts.addFunction('getFormat', async (context: TurnContext, memory: Memory) => {
    return "Html";
});

// Define storage and application
const storage = new MemoryStorage();
const app = new Application<ApplicationTurnState>({
    storage,
    ai: {
        planner
    }
});


// welcome init mesasge... Listen for new members to join the conversation
app.conversationUpdate('membersAdded', async (context: TurnContext, state: ApplicationTurnState) => {
    
        await context.sendActivity("Welcome in support bot....");
    
});

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

app.ai.defaultAction("unknow", async (context: TurnContext, state: ApplicationTurnState, data: any) => {
    await context.sendActivity(`I have not detect any custom action, I am going to ask AI:`);
    return "";
});

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

// Add a custom response formatter to convert markdown code blocks to <pre> tags
addResponseFormatter(app);

prepareSqlActions(app);

prepareTransformationActions(app);

addAppMessages(app);

// Listen for incoming server requests.
server.post('/api/messages', async (req, res) => {
    // Route received a request to adapter for processing
    await adapter.process(req, res as any, async (context) => {
        // Dispatch to application for routing
        await app.run(context);
    });
});

import { DefaultConversationState, DefaultTempState, DefaultUserState, TurnState } from "@microsoft/teams-ai";

// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface ConversationState extends DefaultConversationState {
    lightsOn: boolean;
    data: any;
    sqlServerName: string;
    database: string;
    query:string;
}

export type ApplicationTurnState = TurnState<ConversationState, DefaultUserState, DefaultTempState>;


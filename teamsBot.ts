import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
} from "botbuilder";

import axios from 'axios';

import { 

  ConfidentialClientApplication, 
  OnBehalfOfRequest,
  SilentFlowRequest
} from "@azure/msal-node";

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      
      // Initialize MSAL
      const msalConfig = {
        auth: {
          clientId: "04e40a47-8f9e-4a24-8fc6-8f5f3a5a9328",
          clientSecret: "t4F8Q~4CLFV1~2zCtvpFTwd~4xL85WQFbIz5Qcpo",
          authority: "https://login.microsoftonline.com/b51d8ca7-faa4-413d-8717-8394e361608f",
        },
      };

      const msalClient = new ConfidentialClientApplication(msalConfig);
      
      // Get the current user's Teams JWT
      const silentRequest: SilentFlowRequest = {
        scopes: ["User.Read"],
        account: {
          homeAccountId: context.activity.from.aadObjectId,
          environment: "login.microsoftonline.com",
          tenantId: "b51d8ca7-faa4-413d-8717-8394e361608f",
          username: context.activity.from.id,
          localAccountId: context.activity.from.aadObjectId,
        }
      };

      // Get users JWT
      const response = await msalClient.acquireTokenSilent(silentRequest);

      // Create a new JWT with a new scope from User's Token (On Behalf of flow)
      const onBehalfOfRequest: OnBehalfOfRequest = {
        oboAssertion: response.accessToken,
        scopes: ["YOUR_NEW_SCOPE"],
      };
      const newTokenResponse = await msalClient.acquireTokenOnBehalfOf(onBehalfOfRequest);

      // Make the API call with new token for API
      const apiResponse = await axios.post('YOUR_API_URL', {
        query: txt
      }, {
        headers: {
          Authorization: `Bearer ${newTokenResponse.accessToken}`
        }
      });

      // return the response to the channel (Might have to parse the response for the channel)
      console.log(apiResponse.data);
      await context.sendActivity(`Echo: ${apiResponse.data}`);

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(
            `Hi there! I'm a Teams bot that will echo what you said to me.`
          );
          break;
        }
      }
      await next();
    });
  }
}

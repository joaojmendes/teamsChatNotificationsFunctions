import * as request from 'request-promise';
import * as moment from 'moment';
import { ISubscriptionResult } from './ISubscriptionResult';
import { IAADToken } from './IAADToken';
import { isContext } from 'vm';
import { getEnviromentVariable } from '../Common/Utils';

let accessToken: string = null;


const clientId: string = getEnviromentVariable('AppClientId');
const clientSecret: string = getEnviromentVariable('AppClientSecret');
const ChatMessageNotificationURL = getEnviromentVariable('ChatMessageNotificationURL');
const encryptioncertificate = getEnviromentVariable('encryptioncertificate');
/**
   *  Get Access Token to MSGraph
   *
   * @returns {Promise<string>}
   */
  export async function getAccessToken(): Promise<string> {
    let options = {
      method: 'POST',
      uri: 'https://login.microsoftonline.com/a0cb7b70-2a99-4bf1-b92a-ee32ca12fb3d/oauth2/v2.0/token',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      form: {
        grant_type: 'client_credentials',
        client_id: `${clientId}`,
        client_secret: `${clientSecret}`,
        scope: 'https://graph.microsoft.com/.default'
      }
    };
    const results = await request(options);
    const aadToken: IAADToken = JSON.parse(results);
    return aadToken.access_token;
  }

  /**
   * Create WebHook for Chat Message
   *
   * @param {string} chatId
   * @returns {Promise<ISubscriptionAddResult>}
   */
  export async function addSubscription(chatId: string): Promise<ISubscriptionResult> {
    let accessToken = await getAccessToken();
    
    try {
      let options = {
        method: 'POST',
        uri: 'https://graph.microsoft.com/beta/subscriptions',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${accessToken}`
        },
        
        body: JSON.stringify({
          changeType: 'created,updated',
          notificationUrl: ChatMessageNotificationURL,
          resource: `/chats/${chatId}/messages`,
          expirationDateTime: moment()
            .add(30, 'minutes')
            .toISOString(),
          clientState: `ID_${chatId}`,
          encryptionCertificate: encryptioncertificate,
          encryptionCertificateId: 'teamsChat',
          includeResourceData: true
        })
      };
      console.log(options);
      const results = await request(options);
      return JSON.parse(results);

    } catch (error) {
      throw new Error(error.message);
    }
  }

  export async function updateSubscription(subscriptionId: string): Promise<ISubscriptionResult> {
    let accessToken = await getAccessToken();
    
    try {
      let options = {
        method: 'PATCH',
        uri: `https://graph.microsoft.com/beta/subscriptions/${subscriptionId}`,
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${accessToken}`
        },
        
        body: JSON.stringify({     
          expirationDateTime: moment()
            .add(30, 'minutes')
            .toISOString()       
        })
      };
      
      const results = await request(options);
      return JSON.parse(results);

    } catch (error) {
      throw new Error(error.message);
    }
  }

 
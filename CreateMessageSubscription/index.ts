import { AzureFunction, Context, HttpRequest } from '@azure/functions';
import * as moment from 'moment';
import DataService from '../DataServices/AzureStorageTableService';
import { ISubscriptionResult } from '../DataServices/ISubscriptionResult';
import { IEntity } from '../DataServices/IEntity';
import { addSubscription, updateSubscription } from '../DataServices/MSGraphServices';
import { IReturnEntity } from '../DataServices/IReturnEntity';
import { getEnviromentVariable } from '../Common/Utils';

const AZURE_STORAGE_URI = getEnviromentVariable('Azure_Storage_URI');
const AZURE_STORAGE_SAS = getEnviromentVariable('Azure_Storage_SAS');
const TABLE_NAME = getEnviromentVariable('TableName');

const dataService = new DataService(TABLE_NAME, AZURE_STORAGE_URI, AZURE_STORAGE_SAS);

const httpTrigger: AzureFunction = async function(context: Context, req: HttpRequest): Promise<void> {
  context.log('HTTP trigger function processed a request.');
  const chatId = req.query.chatId || (req.body && req.body.chatId);

  let _subscriptionId: string = '';
  let _subsResult: ISubscriptionResult = {} as ISubscriptionResult;
  let _entity: IEntity = {} as IEntity;
  if (chatId) {
    try {
      // Try to Get a Saved ChatId from Table Storage
      const _returnSubsEntity: IReturnEntity = await dataService.getEntity('TeamsChats', chatId);
      if (!_returnSubsEntity) { 

        // New CHatId add subscriptions and add to Table Storage
          _subsResult  = await addSubscription(chatId);
      } else {
        // subs exist in Table , check expeiration Date 
        _subscriptionId = _returnSubsEntity.SubscriptionId;
        const _expired:boolean = moment(_returnSubsEntity.ExpirationDateTime).isBefore(moment());
        // exprired ? Add new subscriptions
        if (_expired){
          _subsResult  = await addSubscription(chatId);
        }else{
          // update subscriptions expiration date
          _subsResult  = await updateSubscription(_subscriptionId);
           
        }
      }
     
     // update Table with new data
        _entity  = {
        PartitionKey: 'TeamsChats',
        RowKey: chatId,
        SubscriptionId: _subsResult.id,
        ExpirationDateTime: _subsResult.expirationDateTime
      };
      console.log('entity', _entity);
      await dataService.insertOrUpdateEntity(_entity);

      context.res = {
        // status: 200, /* Defaults to 200 */
        body:  { subscriptionId : _subsResult.id }
      };
    } catch (error) {
      context.res = {
        status: 400,
        body: JSON.stringify(error)
      };
    }
  } else {
    context.res = {
      status: 400,
      body: 'ChatId is Missing'
    };
  }
};

export default httpTrigger;

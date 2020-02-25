import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import DataService from '../DataServices/AzureStorageTableService';
import {getEnviromentVariable } from '../Common/Utils';

const AZURE_STORAGE_URI = getEnviromentVariable('Azure_Storage_URI');
const AZURE_STORAGE_SAS = getEnviromentVariable('Azure_Storage_SAS');
const TABLE_NAME = getEnviromentVariable('TableName');

const dataService = new DataService('TeamsChatSubsScriptions', AZURE_STORAGE_URI, AZURE_STORAGE_SAS);

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    let name = (req.query.name || (req.body && req.body.name));
    let v1 ='';
    name ='a';
    if (name) {
       
        const t = await  dataService.createTableIfNotExists();
        
        
        const returnSubsEntity:any = await dataService.getEntity('TeamsChats', '19:8af6cc7eff4f48eb97cab001ff19e112@thread.v2');
         v1 = returnSubsEntity;

         const entity: any = {
            PartitionKey: 'TeamsChats',
            RowKey:  '19:8af6cc7eff4f48eb97cab001ff19e112@thread.v2',
            SubscriptionId: 'aaa',
            ExpirationDateTime: '2020-02-17'
          };

        const res = await dataService.insertOrUpdateEntity(entity);
        v1 = v1 + "-" + res;
        context.res = {
            // status: 200, /* Defaults to 200 */
           
            body: "Hello " + v1
        };
    }
    else {
        context.res = {
            status: 400,
            body: "Please pass a name on the query string or in the request body"
        };
    }
};

export default httpTrigger;

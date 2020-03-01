import * as request from 'request-promise';
import { getEnviromentVariable } from '../Common/Utils';
import { IReturnEntity } from './IReturnEntity';
import { IEntity } from './IEntity';

class DataService {
  constructor(private TABLE_NAME: string, private AZURE_STORAGE_URI: string, private AZURE_STORAGE_SAS: string) {}

  /**
   *  Create Table
   *
   * @returns {Promise<string>}
   * @memberof DataService
   */
  public async createTableIfNotExists(): Promise<string> {
    try {
      const tableExists: boolean = await this.checkIfTableExists();
      if (tableExists) {
        return this.TABLE_NAME;
      } else {
        const URI: string = `${this.AZURE_STORAGE_URI}/tables?${this.AZURE_STORAGE_SAS}`;
        let options = {
          method: 'POST',
          uri: URI,
          headers: {
            'Content-Type': 'application/json;odata=nometadata'
          },
          body: {
            TableName: this.TABLE_NAME
          },
          json: true,
          resolveWithFullResponse: true
        };
        const results = await request(options);
        const tablesResult = JSON.parse(results);
        return tablesResult.TABLE_NAME;
      }
    } catch (error) {
      throw new Error(error);
    }
  }

  /**
   *  Check if Table Exists
   *
   * @returns {Promise<boolean>}
   * @memberof DataService
   */
  public async checkIfTableExists(): Promise<boolean> {
    try {
      const URI: string = `${this.AZURE_STORAGE_URI}/tables?$filter=TableName eq '${this.TABLE_NAME}'&${this.AZURE_STORAGE_SAS}`;
      let options = {
        method: 'GET',
        uri: URI,
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json'
        }
      };
      const results = await request(options);
      const tablesResult = JSON.parse(results);
      return (tablesResult.value && tablesResult.value.length) > 0 ? true : false;
    } catch (error) {
      throw new Error(error);
    }
  }

  /**
   * Insert Entity
   *
   * @param {IEntity} entity
   * @returns {Promise<IReturnEntity>}
   * @memberof DataService
   */
  public async insertEntity(entity: IEntity): Promise<IReturnEntity> {
    try {
      await this.createTableIfNotExists();

      const URI: string = `${this.AZURE_STORAGE_URI}/${this.TABLE_NAME}?${this.AZURE_STORAGE_SAS}`;
      let options = {
        method: 'POST',
        uri: URI,
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json'
        },
        body: {
          PartitionKey: entity.PartitionKey,
          RowKey: entity.RowKey,
          SubscriptionId: entity.SubscriptionId,
          ExpirationDateTime: entity.ExpirationDateTime
        },
        json: true,
        resolveWithFullResponse: true
      };
      const results = await request(options);
      const entityAddedResult: IReturnEntity = JSON.parse(results);
      return entityAddedResult;
    } catch (error) {
      throw new Error(error);
    }
  }

  /**
   * Insert or Upadate Entity
   *
   * @param {IEntity} entity
   * @returns {Promise<void>}
   * @memberof DataService
   */
  public async insertOrUpdateEntity(entity: IEntity): Promise<void> {
    try {
      await this.createTableIfNotExists();

      const URI: string = `${this.AZURE_STORAGE_URI}/${this.TABLE_NAME}(PartitionKey='${entity.PartitionKey}', RowKey='${entity.RowKey}')?${this.AZURE_STORAGE_SAS}`;
      let options = {
        method: 'PUT',
        uri: URI,
        headers: {
          'Content-Type': 'Application/json',
          Accept: 'Aplication/json'
        },
        body: {
          PartitionKey: entity.PartitionKey,
          RowKey: entity.RowKey,
          SubscriptionId: entity.SubscriptionId,
          ExpirationDateTime: entity.ExpirationDateTime
        },
        json: true,
        resolveWithFullResponse: true
      };
      const results = await request(options);
    } catch (error) {
      throw new Error(error);
    }
  }

  /**
   *  Update Entity
   *
   * @param {IEntity} entity
   * @returns {Promise<void>}
   * @memberof DataService
   */
  public async updateEntity(entity: IEntity): Promise<void> {
    try {
      await this.createTableIfNotExists();

      const URI: string = `${this.AZURE_STORAGE_URI}/${this.TABLE_NAME}(PartitionKey='${entity.PartitionKey}', RowKey='${entity.RowKey}')?${this.AZURE_STORAGE_SAS}`;
      let options = {
        method: 'PUT',
        uri: URI,
        headers: {
          'Content-Type': 'Application/json',
          Accept: 'Application/json',
          'if-match': '*'
        },
        body: {
          PartitionKey: entity.PartitionKey,
          RowKey: entity.RowKey,
          SubscriptionId: entity.SubscriptionId,
          ExpirationDateTime: entity.ExpirationDateTime
        },
        json: true,
        resolveWithFullResponse: true
      };
      const results = await request(options);
    } catch (error) {
      throw new Error(error);
    }
  }

  /**
   * Remove Entity
   *
   * @param {IEntity} entity
   * @returns {Promise<void>}
   * @memberof DataService
   */
  public async removeEntity(entity: IEntity): Promise<void> {
    try {
      await this.createTableIfNotExists();

      const URI: string = `${this.AZURE_STORAGE_URI}/${this.TABLE_NAME}(PartitionKey='${entity.PartitionKey}', RowKey='${entity.RowKey}')?${this.AZURE_STORAGE_SAS}`;
      let options = {
        method: 'DELETE',
        uri: URI,
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json'
        }
      };
      const results = await request(options);
    } catch (error) {
      throw new Error(error);
    }
  }

  /**
   *  Get Entity by Partition and rowKey
   *
   * @param {string} partitionKey
   * @param {string} rowKey
   * @returns {Promise<IReturnEntity>}
   * @memberof DataService
   */
  public async getEntity(partitionKey: string, rowKey: string): Promise<IReturnEntity> {
    try {
      await this.createTableIfNotExists();

      const URI: string = `${this.AZURE_STORAGE_URI}/${this.TABLE_NAME}(PartitionKey='${partitionKey}', RowKey='${rowKey}')?${this.AZURE_STORAGE_SAS}`;
      console.log('URI', URI);
      let options = {
        method: 'GET',
        uri: URI,
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json'
        }
        
      };
      var results = await request(options);

      var entityResult: IReturnEntity = JSON.parse(results);
      return entityResult;
    } catch (error) {
        if (error.statusCode == 404){ // not found
            return entityResult = undefined;
        }
      throw new Error(error);
    }
  }

  /**
   *
   *  Get Entities from Table base on Query 
   *   accepts ODATA $filter and $select
   * @param {string} query
   * @returns {Promise<IReturnEntity[]>}
   * @memberof DataService
   */
  public async listEntities(query?: string): Promise<IReturnEntity[]> {
    try {
    let _query:string = '';
      await this.createTableIfNotExists();
      if (query){
        _query=`&$filter=${query}`;
      }
      const URI: string = `${this.AZURE_STORAGE_URI}/${this.TABLE_NAME}()?${this.AZURE_STORAGE_SAS}${_query}`;
      let options = {
        method: 'GET',
        uri: URI,
        headers: {
          'Content-Type': 'application/json;odata=nometadata',
          Accept: 'application/json'
        }
      };
      const results = JSON.parse(await request(options));
      const entitiesListResults: IReturnEntity[] = results.value;
      return entitiesListResults;
    } catch (error) {
      throw new Error(error);
    }
  }
}

export default DataService;

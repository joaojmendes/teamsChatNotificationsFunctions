## Teams Notifications Functions

This functions are use by SPFx Teams Notifications Application Customizer 



### Azure Services need to be configured

. Azure SignalR
. Azure Storage - we use Table Storage to save data - the Functions use Azure Storage Table REST API to read and write data, the authentication that is used is SAS - Shared Access Assignature Connection String that must be generated. the Table is created automaticaly by the function.
. Azure Functions App 
    - In the CORS options you have to enable Enable Access-Control-Allow-Credentials  that is need by SignalR
    - In the Plataform Features Tab select Identity in System Assigned Tab on Status Click Enable 
- Register a new App Principal on AAD and add the API Permissions SCOPES - Chat.Read.All and User.Read
- Azure Key Vault 
    - Create Self-Signed Certificate
    - Export Certificate CER

### Configuration

The local.settings.json file has all environment variables uses and have to be configured, after Azure Services configured.
environment variable Sample| comments
--------------------|----|--------|----------
    "AzureSignalRConnectionString"| "Endpoint=https|//xxxxxx.service.signalr.net;AccessKey=99999VFafXK/UjpRH0YY5g8ivxcN8IE0Mn9gCIjuBDwQ4=;Version=1.0;", |  SignalR Connection String 
    "Azure_Storage_URI"| "https|//txxxxxxxx.table.core.windows.net", | Azure Table Storage URI
    "Azure_Storage_SAS"| "sv=2019-02-02&ss=bfqt&srt=sco&sp=rwdlacup&se=2050-02-16T01|48|23Z&st=2020-02-15T17|48|23Z&spr=https&sig=%2BT7OE%2B5tGT%2Fuzb6QAA3X5bZff1f9f9TL29CCJeyo7pI%3D", | Table Storage SAS connection String
    "TableName"| "TeamsChatSubsScriptions" |  Name of Azure Storage Table that will be Created 
    "TenantId"| "xxxxxxxxxxxxxxxxxx" |  Tenant Id
    "KeyVaultUrl"| "https|//xxxxxxxx.vault.azure.net/" | Azure Key Vault URL
    "KeyVault_ClientId_Secret_Name"|"TeamsChatNotificationsClientId" |  Name of Secret Name of Client Id
    "KeyVault_ClientSecret_Secret_Name"| "TeamsChatNotificationsClientSecret" | Name of Secret Name of Client Secret
    "KeyVault_EncriptionCertificate_Secret_Name"| "TeamsChatNotificationEncryptionCertificate" | Name of Client Secret with Exported Certificate with Public Key (Content of CER file)
    "ChatMessageNotificationURL"|"https|//xxxxxxxxxx.azurewebsites.net/api/TeamsChatWebhook" |  URL of TeamsChatWebhook , define only the URL of Azure Function App
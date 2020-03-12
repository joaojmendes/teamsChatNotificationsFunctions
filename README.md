# SPFx Teams Notifications Application Customizer  - Azure Functions

## Summary
This Project is part of SPFx Teams Notifications Application Customizer 
and has details of what services must have to be defined in Azure and the Azure Functions to deploy.

 
## Azure Services need to be configured

- Azure SignalR

- Azure Storage - we use Table Storage to save data - the Functions use Azure Storage Table REST API to read and write data, the authentication that is used is SAS - Shared Access Assignature Connection String that must be generated. the Table is created automaticaly by the function.

- Azure Functions App 
    - In the CORS options you have to enable Enable Access-Control-Allow-Credentials  that is need by SignalR
    - In the Plataform Features Tab select Identity in System Assigned Tab on Status Click Enable 

- Register a new App Principal on AAD and add the API Permissions SCOPES - Chat.Read.All and User.Read

- Azure Key Vault 
    - Create Self-Signed Certificate
    - Export Certificate CER

## Configuration local.setting.json 

Environment Variable | Required Value
--------------------|------------------------------------ 
Azure_Storage_URI | < Azure Table Storage URI >
Azure_Storage_SAS | < Table Storage SAS connection String >
TableName|  TeamsChatSubsScriptions 
TenantId| < Tenant Id > 
KeyVaultUrl | < https://xxxxxxxx.vault.azure.net/ >
KeyVault_ClientId_Secret_Name | TeamsChatNotificationsClientId 
KeyVault_ClientSecret_Secret_Name | TeamsChatNotificationsClientSecret 
KeyVault_EncriptionCertificate_Secret_Name |  TeamsChatNotificationEncryptionCertificate 
ChatMessageNotificationURL| < https://xxxxxxxxxx.azurewebsites.net/api/TeamsChatWebhook >


## Solution

Solution|Author(s)
--------|---------
SPFx - Teams Notifications Application Customizer

## Version history

Version|Date|Comments
-------|----|--------
1.0.0|Mars 12, 2020 | Initial Version

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---




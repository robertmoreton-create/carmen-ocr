param location string = resourceGroup().location
param environment string = 'production'
param appName string = 'carmen-ocr'

// Shared App Service Plan - Basic tier for multiple apps
resource appServicePlan 'Microsoft.Web/serverfarms@2022-03-01' = {
  name: 'shared-plan-${environment}'
  location: location
  sku: {
    name: 'B1'
    tier: 'Basic'
    size: 'B1'
    family: 'B'
    capacity: 1
  }
  kind: 'linux'
  properties: {
    reserved: true
  }
}

// App Service for Carmen OCR
resource appService 'Microsoft.Web/sites@2022-03-01' = {
  name: '${appName}-app-${environment}'
  location: location
  kind: 'app,linux'
  properties: {
    serverFarmId: appServicePlan.id
    siteConfig: {
      pythonVersion: '3.11'
      appCommandLine: 'gunicorn --bind 0.0.0.0:8000 app:app'
      alwaysOn: true
    }
    httpsOnly: true
  }
}

// Access Restriction - Allow only specific IP
resource accessRestriction 'Microsoft.Web/sites/config@2022-03-01' = {
  parent: appService
  name: 'web'
  properties: {
    ipSecurityRestrictions: [
      {
        ipAddress: '84.247.40.92/32'
        action: 'Allow'
        priority: 100
        name: 'Office-IP'
        description: 'Allow access from office IP'
      }
      {
        ipAddress: 'AzureFrontDoor.Backend'
        action: 'Allow'
        priority: 200
        name: 'AzureFrontDoor'
        description: 'Allow Azure Front Door'
        tag: 'ServiceTag'
      }
    ]
    scmIpSecurityRestrictions: [
      {
        ipAddress: '84.247.40.92/32'
        action: 'Allow'
        priority: 100
        name: 'Office-IP-SCM'
        description: 'Allow SCM access from office IP'
      }
    ]
    scmIpSecurityRestrictionsUseMain: false
  }
}

// Document Intelligence
resource documentIntelligence 'Microsoft.CognitiveServices/accounts@2023-05-01' = {
  name: '${appName}-di-${environment}'
  location: location
  kind: 'FormRecognizer'
  sku: {
    name: 'S0'
  }
  properties: {
    publicNetworkAccess: 'Enabled'
    customSubDomainName: '${appName}-di-${environment}'
  }
}

// Storage Account for file uploads
resource storageAccount 'Microsoft.Storage/storageAccounts@2022-09-01' = {
  name: 'carmenstorage${environment}'
  location: location
  sku: {
    name: 'Standard_LRS'
  }
  kind: 'StorageV2'
  properties: {
    accessTier: 'Hot'
    allowBlobPublicAccess: false
  }
}

// Blob container for uploads
resource blobContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2022-09-01' = {
  name: '${storageAccount.name}/default/uploads'
  properties: {
    publicAccess: 'None'
  }
}

// App settings
resource appSettings 'Microsoft.Web/sites/config@2022-03-01' = {
  parent: appService
  name: 'appsettings'
  properties: {
    AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT: documentIntelligence.properties.endpoint
    AZURE_DOCUMENT_INTELLIGENCE_KEY: documentIntelligence.listKeys().key1
    AZURE_STORAGE_CONNECTION_STRING: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=core.windows.net'
    SCM_DO_BUILD_DURING_DEPLOYMENT: 'true'
    WEBSITES_PORT: '8000'
  }
}

// Outputs
output appServiceName string = appService.name
output appServiceUrl string = 'https://${appService.properties.defaultHostName}'
output documentIntelligenceEndpoint string = documentIntelligence.properties.endpoint
output appServicePlanName string = appServicePlan.name

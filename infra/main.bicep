// 효도택시 인프라: Azure Static Web Apps (정적 화면 + 관리형 Functions API)
// 배포: az deployment group create -g <리소스그룹> -f infra/main.bicep -p appName=<이름>

@description('앱 이름 (영문 소문자·숫자·하이픈)')
param appName string = 'hyodo-taxi'

@description('Static Web Apps 지역 (eastasia, eastus2, westeurope, westus2, centralus 중 하나)')
@allowed([
  'eastasia'
  'eastus2'
  'westeurope'
  'westus2'
  'centralus'
])
param location string = 'eastasia'

@description('Azure OpenAI 엔드포인트 (예: https://xxx.openai.azure.com). 비우면 규칙 기반으로 동작')
param openAiEndpoint string = ''

@description('Azure OpenAI 배포 이름 (예: gpt-4o)')
param openAiDeployment string = ''

@secure()
@description('Azure OpenAI API 키')
param openAiKey string = ''

resource site 'Microsoft.Web/staticSites@2023-12-01' = {
  name: appName
  location: location
  sku: {
    name: 'Free'
    tier: 'Free'
  }
  properties: {
    stagingEnvironmentPolicy: 'Enabled'
    allowConfigFileUpdates: true
  }
  tags: {
    project: 'hyodo-taxi'
    event: 'cheongdo-vibecoding-hackathon-2026'
  }
}

resource settings 'Microsoft.Web/staticSites/config@2023-12-01' = {
  parent: site
  name: 'appsettings'
  properties: {
    AZURE_OPENAI_ENDPOINT: openAiEndpoint
    AZURE_OPENAI_DEPLOYMENT: openAiDeployment
    AZURE_OPENAI_API_KEY: openAiKey
    AZURE_OPENAI_API_VERSION: '2024-10-21'
  }
}

output siteName string = site.name
output url string = 'https://${site.properties.defaultHostname}'

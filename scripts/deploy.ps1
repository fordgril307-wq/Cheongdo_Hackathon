# 효도택시 원클릭 배포 (Windows PowerShell)
# 사용법: .\scripts\deploy.ps1 -ResourceGroup <리소스그룹> [-AppName <앱이름>]
#   - 리소스그룹이 없으면 만들고, 있으면(운영진 제공) 그대로 씁니다.
#   - Azure OpenAI 값은 환경변수로 넘기면 함께 설정됩니다:
#     $env:AZURE_OPENAI_ENDPOINT, $env:AZURE_OPENAI_DEPLOYMENT, $env:AZURE_OPENAI_API_KEY
param(
  [Parameter(Mandatory = $true)][string]$ResourceGroup,
  [string]$AppName = "hyodo-taxi-$(Get-Random -Maximum 99999)",
  [string]$Location = "eastasia"
)
$ErrorActionPreference = "Stop"
Set-Location (Join-Path $PSScriptRoot "..")

Write-Host "1/5 Azure 로그인 확인"
az account show 1>$null 2>$null
if ($LASTEXITCODE -ne 0) { az login --use-device-code }

Write-Host "2/5 리소스 그룹 준비: $ResourceGroup"
if ((az group exists -n $ResourceGroup) -eq "false") {
  az group create -n $ResourceGroup -l koreacentral -o none
}

Write-Host "3/5 Static Web Apps 생성 (Bicep): $AppName"
$url = az deployment group create -g $ResourceGroup -f infra/main.bicep `
  -p appName=$AppName location=$Location `
     openAiEndpoint="$env:AZURE_OPENAI_ENDPOINT" `
     openAiDeployment="$env:AZURE_OPENAI_DEPLOYMENT" `
     openAiKey="$env:AZURE_OPENAI_API_KEY" `
  --query properties.outputs.url.value -o tsv
if ($LASTEXITCODE -ne 0) { throw "Bicep 배포 실패" }

Write-Host "4/5 코드 업로드"
$token = az staticwebapp secrets list -n $AppName -g $ResourceGroup --query properties.apiKey -o tsv
npm install --prefix api --silent
npx --yes @azure/static-web-apps-cli@2 deploy ./app `
  --api-location ./api --api-language node --api-version 20 `
  --deployment-token $token --env production
if ($LASTEXITCODE -ne 0) { throw "코드 업로드 실패" }

Write-Host "5/5 GitHub 자동 배포 연결 (gh CLI가 있으면)"
if (Get-Command gh -ErrorAction SilentlyContinue) {
  gh secret set AZURE_STATIC_WEB_APPS_API_TOKEN --body $token
  Write-Host "   → 이제 main에 push하면 자동 배포됩니다."
} else {
  Write-Host "   → gh CLI가 없어 건너뜀. README의 'GitHub 자동 배포 연결'을 따라 secret을 등록하세요."
}

Write-Host ""
Write-Host "배포 완료: $url"
Write-Host "상태 확인: $url/api/health"

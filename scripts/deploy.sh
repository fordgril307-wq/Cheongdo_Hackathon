#!/usr/bin/env bash
# 효도택시 원클릭 배포 (macOS·Linux·Codespaces)
# 사용법: ./scripts/deploy.sh <리소스그룹> [앱이름]
#   - 리소스그룹이 없으면 만들고, 있으면(운영진 제공) 그대로 씁니다.
#   - Azure OpenAI 값은 환경변수로 넘기면 함께 설정됩니다:
#     AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_DEPLOYMENT, AZURE_OPENAI_API_KEY
set -euo pipefail

RG="${1:?리소스 그룹 이름을 입력하세요. 예) ./scripts/deploy.sh rg-hyodo}"
APP="${2:-hyodo-taxi-$RANDOM}"
LOCATION="${LOCATION:-eastasia}"
cd "$(dirname "$0")/.."

echo "1/5 Azure 로그인 확인"
az account show >/dev/null 2>&1 || az login --use-device-code

echo "2/5 리소스 그룹 준비: $RG"
if [ "$(az group exists -n "$RG")" = "false" ]; then
  az group create -n "$RG" -l koreacentral -o none
fi

echo "3/5 Static Web Apps 생성 (Bicep): $APP"
URL=$(az deployment group create -g "$RG" -f infra/main.bicep \
  -p appName="$APP" location="$LOCATION" \
     openAiEndpoint="${AZURE_OPENAI_ENDPOINT:-}" \
     openAiDeployment="${AZURE_OPENAI_DEPLOYMENT:-}" \
     openAiKey="${AZURE_OPENAI_API_KEY:-}" \
  --query properties.outputs.url.value -o tsv)

echo "4/5 코드 업로드"
TOKEN=$(az staticwebapp secrets list -n "$APP" -g "$RG" --query properties.apiKey -o tsv)
npm install --prefix api --silent
npx --yes @azure/static-web-apps-cli@2 deploy ./app \
  --api-location ./api --api-language node --api-version 20 \
  --deployment-token "$TOKEN" --env production

echo "5/5 GitHub 자동 배포 연결 (gh CLI가 있으면)"
if command -v gh >/dev/null 2>&1 && gh auth status >/dev/null 2>&1; then
  gh secret set AZURE_STATIC_WEB_APPS_API_TOKEN --body "$TOKEN" && echo "   → 이제 main에 push하면 자동 배포됩니다."
else
  echo "   → gh CLI가 없어 건너뜀. README의 'GitHub 자동 배포 연결'을 따라 secret을 등록하세요."
fi

echo ""
echo "배포 완료: $URL"
echo "상태 확인: $URL/api/health"

# `{{ 팀 이름 }}` - `청도효도택시`

해커그라운드 해커톤에 참여하는 `{{ 팀 이름 }}` 팀의 `청도효도택시`입니다.
청도 어르신이 사투리로 말하면 AI가 예약을 정리하고, 같은 날·같은 병원 가는 분들을 한 차로 묶어 드립니다.

- 배포 주소: {{ https://xxxx.azurestaticapps.net }}
- 상태 확인: {{ 배포 주소 }}/api/health

## 제품/서비스 소개

<!-- 아래 링크는 지우지 마세요 -->
[제품/서비스 소개 보기](TOPIC.md)
<!-- 위 링크는 지우지 마세요 -->

## 오픈 소스 라이센스

<!-- 아래 링크는 지우지 마세요 -->
[오픈소스 라이센스 보기](./LICENSE)
<!-- 위 링크는 지우지 마세요 -->

## 설치 방법

> 아래 순서를 위에서부터 그대로 따라 하면 누구나 Microsoft Azure에 배포할 수 있습니다. 명령어는 복사해서 붙여 넣으세요.
> 가장 쉬운 방법은 **방법 A (GitHub Codespaces)** 입니다. 내 컴퓨터에 아무것도 설치하지 않아도 됩니다.

### 사전 준비 사항

| 필요한 것 | 방법 A (Codespaces) | 방법 B (내 Windows PC) |
|---|---|---|
| GitHub 계정 | 필요 | 필요 |
| Azure 구독 (리소스 그룹 만들기 권한) | 필요 | 필요 |
| Node.js 20 이상 | 자동 설치됨 | [nodejs.org](https://nodejs.org/ko)에서 LTS 설치 |
| Azure CLI | 자동 설치됨 | [설치 안내](https://learn.microsoft.com/ko-kr/cli/azure/install-azure-cli-windows) (MSI 다운로드 후 실행) |
| Git | 자동 설치됨 | [git-scm.com](https://git-scm.com/download/win)에서 64-bit 설치 (옵션은 기본값) |
| (선택) Azure OpenAI 엔드포인트·키·배포 이름 | 있으면 AI 해석, 없으면 규칙 기반으로 동작 | 동일 |

설치 확인 (방법 B만): PowerShell을 새로 열고 아래 세 줄이 모두 버전 숫자를 출력하면 준비 끝입니다.

```powershell
node --version
az --version
git --version
```

### 방법 A — GitHub Codespaces로 배포 (추천)

1. 이 저장소 페이지에서 초록색 **Code** 버튼 → **Codespaces** 탭 → **Create codespace on main**을 누릅니다.
2. 화면 아래 터미널이 열리고 `npm install`이 끝날 때까지 1~2분 기다립니다.
3. (선택) Azure OpenAI를 쓰려면 터미널에 아래를 입력합니다. 값은 운영진이 준 것으로 바꿉니다.

   ```bash
   export AZURE_OPENAI_ENDPOINT="https://<리소스이름>.openai.azure.com"
   export AZURE_OPENAI_DEPLOYMENT="<배포 이름, 예: gpt-4o>"
   export AZURE_OPENAI_API_KEY="<키>"
   ```

4. 배포 명령을 실행합니다. `<리소스그룹>`은 운영진이 준 리소스 그룹 이름(없으면 새 이름, 예: `rg-hyodo`)으로 바꿉니다.

   ```bash
   ./scripts/deploy.sh <리소스그룹>
   ```

5. 처음이면 `az login` 안내가 나옵니다. 화면에 나온 주소를 브라우저에서 열고, 화면에 나온 코드를 입력해 Azure 계정으로 로그인합니다.
6. 마지막 줄에 `배포 완료: https://...azurestaticapps.net` 이 나오면 성공입니다. 그 주소를 브라우저에서 엽니다.

### 방법 B — 내 Windows PC에서 배포

1. PowerShell을 엽니다 (시작 메뉴에서 "PowerShell" 검색).
2. 저장소를 내려받고 폴더로 이동합니다.

   ```powershell
   git clone https://github.com/<저장소 주소>.git
   cd <저장소 폴더 이름>
   ```

3. 스크립트 실행이 막혀 있으면 이 창에서만 허용합니다.

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```

4. (선택) Azure OpenAI 값을 넣습니다.

   ```powershell
   $env:AZURE_OPENAI_ENDPOINT = "https://<리소스이름>.openai.azure.com"
   $env:AZURE_OPENAI_DEPLOYMENT = "<배포 이름>"
   $env:AZURE_OPENAI_API_KEY = "<키>"
   ```

5. 배포합니다.

   ```powershell
   .\scripts\deploy.ps1 -ResourceGroup <리소스그룹>
   ```

6. 브라우저가 열리면 Azure 계정으로 로그인합니다. 마지막 줄의 `배포 완료:` 주소를 엽니다.

### 배포 스크립트가 하는 일 (5단계)

1. Azure 로그인 확인
2. 리소스 그룹 준비 (있으면 그대로 사용)
3. `infra/main.bicep`으로 Azure Static Web Apps 생성 + Azure OpenAI 설정값 등록
4. 웹 화면(`app/`)과 API(`api/`) 업로드
5. GitHub 자동 배포용 토큰을 저장소 secret에 등록 (gh CLI가 로그인된 경우)

### GitHub 자동 배포(CI/CD) 연결

배포 스크립트 5단계가 건너뛰어졌다면 직접 연결합니다.

1. 토큰 확인: `az staticwebapp secrets list -n <앱이름> -g <리소스그룹> --query properties.apiKey -o tsv`
2. GitHub 저장소 → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**
3. Name: `AZURE_STATIC_WEB_APPS_API_TOKEN`, Secret: 1번에서 나온 값 → **Add secret**
4. 이후 `main` 브랜치에 push하면 `.github/workflows/azure-static-web-apps.yml`이 테스트를 돌리고 자동 배포합니다. 진행 상황은 저장소 **Actions** 탭에서 볼 수 있습니다.

### 배포 확인

| 확인 | 기대 결과 |
|---|---|
| `<배포 주소>/api/health` | `{"ok":true,"service":"hyodo-taxi","ai":true}` (AI 키가 없으면 `"ai":false`) |
| `<배포 주소>` → **어르신** → **대구 효성병원** → **내일**·**오전 10시** → **다음** → **예약할게요** | "예약됐어요" 화면에 탑승 시각·내 자리·함께 타는 분 수 |
| **어르신** → **글로 적기** → 예시 문장 → **AI로 정리하기** | 날짜·시간까지 채워져 바로 확인 화면으로 이동 |
| **어르신** → **이장님·경로당이 대신** → 어르신 2분 체크 → **청도대남병원** → 날짜·시간 → **2분 예약할게요** | 두 분이 같은 차·같은 탑승 시각으로 묶인 명단 |
| **배차 담당** 탭 | '지금 방식(한 분씩 운행)'과 '청도효도택시(통합 배차)'의 운행·군 정산 건수 비교, 지도에 차량별 동선 |

### 운영 설정 (선택)

`app/config.js`
- `callCenter`: 전화 예약 번호. 넣으면 첫 화면에 "전화로 예약하기" 버튼이 나타나고, 비워 두면 숨겨집니다.
- `happyTaxiFare`: 행복택시 본인 부담금(기본 100원, 청도군 확인 필요). API 쪽은 Static Web Apps 환경 변수 `HAPPY_TAXI_FARE`로 바꿉니다.

### 내 컴퓨터에서 미리 보기 (Azure 없이)

```bash
npm install --prefix api
node local/dev-server.js
```

브라우저에서 http://localhost:4280 을 엽니다. 테스트는 `node --test tests/core.test.js` 입니다.

### 문제 해결

| 증상 | 해결 |
|---|---|
| `az: command not found` / `'az'은(는) ... 아닙니다` | Azure CLI 설치 후 터미널을 **새로** 엽니다 |
| `The subscription is not registered to use namespace 'Microsoft.Web'` | `az provider register --namespace Microsoft.Web` 실행 후 1~2분 뒤 다시 배포 |
| `Location 'koreacentral' is not available for resource type 'Microsoft.Web/staticSites'` | 정상입니다. Static Web Apps는 `eastasia`에 만들어지고, 리소스 그룹만 한국 중부에 있습니다 |
| 이름 중복 오류 | `deploy.sh <리소스그룹> <다른앱이름>` 처럼 앱 이름을 바꿔 다시 실행 |
| 화면은 뜨는데 "API에 연결하지 못했어요" | 2~3분 뒤 새로고침 (API 배포가 조금 늦게 끝남). 계속되면 `/api/health` 확인 |
| 예약 카드에 "규칙 기반(AI 키 없음)" | Azure OpenAI 값이 없을 때의 정상 동작. 값을 넣고 다시 배포하거나 Azure Portal → Static Web App → **환경 변수**에서 추가 |

## 폴더 구조

```
app/        웹 화면 (index.html, app.js, app.css, 지도 라이브러리)
api/        Azure Functions API (사투리 해석, 합승 배차)
infra/      Bicep 인프라 코드
scripts/    원클릭 배포 스크립트 (deploy.sh, deploy.ps1)
tests/      핵심 로직 테스트
local/      Azure 없이 미리 보는 로컬 서버
images/     아키텍처 그림
decks/      발표 자료
```

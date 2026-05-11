# HWP Proposal Cloud

순수 정적 사이트로 배포되는 클라우드 버전입니다. 서버는 GPT 호출까지만 (사실은 브라우저가 OpenAI에 직접 호출) 처리하고, 한글 변환은 사용자의 본인 PC에서 실행합니다.

## 구조

- `index.html`, `app.js`, `styles.css` — 프론트엔드 (Vercel/Netlify/GitHub Pages 어디든 정적 배포)
- `pipeline/*.ps1` — 한글/Word 변환 스크립트. 브라우저가 fetch해서 런타임에 ZIP에 묶어 다운로드 제공
- `vercel.json` — `/pipeline/*.ps1`을 text/plain으로 서빙해 브라우저가 받을 수 있게 함

## 흐름

1. 사용자가 사이트 접속
2. 본인 OpenAI API key + HWPX 양식 + 작성 지시 입력
3. 브라우저가 HWPX 풀어서 텍스트 추출 → OpenAI API 직접 호출 → Markdown 받음
4. 브라우저가 `pipeline/*.ps1` + `generated.md` + `run.cmd` 묶어서 `hwp-runner-YYYYMMDDHHmmss.zip` 다운로드
5. 사용자가 본인 Windows PC에서 ZIP 풀고 `run.cmd` 더블클릭 → Word + 한글 COM이 final.hwpx 생성

## 보안

- OpenAI API 키는 브라우저에 입력되고 OpenAI에만 전송됩니다 (`api.openai.com`).
- 이 사이트 자체에는 백엔드가 없으므로 키가 서버에 저장되지 않습니다.
- 사용자가 원하면 localStorage에 저장되며, 브라우저 데이터 지우기로 즉시 삭제됩니다.

## 배포

Vercel:

1. Vercel에서 이 GitHub 저장소를 import
2. Framework: Other (정적 사이트)
3. Build/Output 설정 그대로 (정적 파일)
4. Deploy

GitHub Pages:

1. Settings → Pages → Source: Deploy from a branch → main / (root)
2. 저장 후 1~2분이면 `https://grantiq-spark.github.io/hwp-proposal-cloud/` 에서 접속 가능

## 관련 저장소

- 변환 코어 단독: https://github.com/grantiq-spark/hwp-converter-core
- 풀스택 로컬 버전(Word COM + 한글 COM 직접): https://github.com/grantiq-spark/hwp-proposal-automation

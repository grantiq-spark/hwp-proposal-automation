# HWP DOCX → HWPX Runner Builder

순수 정적 사이트. 백엔드도 AI API도 없습니다.

## 흐름

1. 사용자가 사이트에서 DOCX 업로드
2. 브라우저가 DOCX + `pipeline/*.ps1` + `run.cmd` + `run.ps1` 을 한 ZIP으로 묶어 다운로드 제공
3. 사용자가 본인 Windows PC에서 ZIP 풀고 `run.cmd` 더블클릭
4. Word + 한컴 한글 COM 자동화로 `final.hwpx` 생성

## 비용

- 0원. API 키 필요 없음. 서버 호출 없음. OpenAI 호출 없음.

## 구조

- `index.html`, `app.js`, `styles.css` — 정적 프론트엔드
- `pipeline/*.ps1` — 변환 스크립트. 브라우저가 fetch해서 ZIP에 묶음
- `vercel.json` — 정적 호스팅 설정

## 배포

Vercel:

1. Vercel에서 이 GitHub 저장소 import
2. Root Directory: `cloud`
3. Framework: Other
4. Deploy

main 브랜치에 push만 하면 자동 재배포.

## 관련 저장소

- 변환 코어: https://github.com/grantiq-spark/hwp-converter-core
- 풀스택 로컬 버전 (이 폴더의 모체): https://github.com/grantiq-spark/hwp-proposal-automation

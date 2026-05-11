# Windows Server Setup

## Required Software

- Windows 10/11 or Windows Server with an interactive desktop session
- Microsoft Word
- Hancom Hangul with COM automation available
- ImageMagick
- Ghostscript
- Node.js 18 or newer
- OpenAI API key for GPT draft mode

## Start Locally

```powershell
cd hwp-saas
.\scripts\run-dev.ps1
```

Open `http://localhost:8787`.

## Environment Variables

Copy `.env.example` into your hosting environment or set variables manually:

- `PORT`: HTTP port. Default `8787`.
- `PIPELINE_ROOT`: folder containing the PowerShell conversion scripts. Default `..`.
- `PIPELINE_FACTOR`: layout stretch factor. Default `2.02`, matching the best current result.
- `MAX_UPLOAD_MB`: upload limit. Default `80`.
- `OPENAI_API_KEY`: required for GPT draft mode.
- `OPENAI_MODEL`: model used for draft writing. Default `gpt-5.2`.
- `OPENAI_MAX_OUTPUT_TOKENS`: draft length cap. Default `12000`.

## Important Hosting Note

Do not run this as a headless Windows service at first. Word and Hangul automation are much more reliable in an active signed-in desktop session. For production, use a locked-down Windows VM that stays signed in, then put a reverse proxy or private network gateway in front of it.

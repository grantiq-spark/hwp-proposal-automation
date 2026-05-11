# HWP Proposal Automation

HTTP app for drafting a Korean proposal from an uploaded HWPX template, generating a DOCX draft with GPT, converting it into HWPX with Word/Hancom Hangul automation, and producing visual similarity reports.

This is the higher-level product app. It includes the converter scripts under `pipeline/`.

## Workflow

1. Upload the HWPX template for the document you want to write.
2. Enter writing instructions and optional reference material.
3. GPT generates a Markdown proposal draft.
4. Word COM creates `source.docx`.
5. The converter pipeline creates `final.hwpx`.
6. The app returns the HWPX, generated DOCX, rendered PDFs, and similarity reports.

## Requirements

- Windows 10/11 or Windows Server with an interactive desktop session
- Microsoft Word
- Hancom Hangul with COM automation
- ImageMagick
- Ghostscript
- Node.js 18+
- OpenAI API key for GPT draft mode

## Run

```powershell
copy .env.example .env
notepad .env
node server.js
```

Open `http://localhost:8787`.

## Configuration

- `PORT`: default `8787`
- `PIPELINE_ROOT`: default `./pipeline`
- `PIPELINE_FACTOR`: default `2.02`
- `MAX_UPLOAD_MB`: default `80`
- `OPENAI_API_KEY`: required for GPT draft mode
- `OPENAI_MODEL`: default `gpt-5.2`
- `OPENAI_MAX_OUTPUT_TOKENS`: default `12000`

## Deployment

Use GitHub for source hosting, but run the app on a Windows VM or workstation. Word and Hancom Hangul COM automation is not reliable in a headless Linux container or serverless hosting platform.

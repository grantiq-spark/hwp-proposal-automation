# Architecture

This app is an HTTP wrapper around a Windows desktop document pipeline. It is intentionally small: no database, no npm dependencies, and no background queue service.

## Request Flow

1. Browser uploads a `.hwpx` template plus writing instructions to `POST /api/jobs`.
2. The server extracts text from the HWPX template.
3. The server calls the OpenAI Responses API to generate a Markdown proposal draft.
4. Word COM converts the Markdown-like draft text into `source.docx`.
5. The server runs the existing PowerShell pipeline:
   - `paste_word_to_new_hangul.ps1`
   - `stretch_hwpx_layout.ps1`
   - `reflow_with_hangul_engine.ps1`
   - `compare_rendered_pages.ps1`
6. The browser polls `GET /api/jobs/<jobId>`.
7. Finished jobs expose `source.docx`, `final.hwpx`, rendered PDFs, and similarity reports.

## Why Windows Is Required

The conversion depends on Microsoft Word COM automation and Hancom Hangul COM automation. Those APIs require a Windows desktop session with the applications installed. This is not suitable for GitHub Pages, Vercel, Netlify, or a normal Linux container.

## Current Limits

- Job status is kept in memory. Restarting the server loses the visible status, although files remain in `data/jobs`.
- There is no authentication.
- Jobs run on the same host process and are best handled one at a time.
- Generated files are retained until manually deleted.

## Production Direction

For a real SaaS, keep this web app as the control plane and run conversion workers on dedicated Windows machines. Add authentication, a persistent job database, storage cleanup, per-user file access control, and a worker queue.

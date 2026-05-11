const http = require("http");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const { spawn } = require("child_process");
const crypto = require("crypto");

loadEnvFile(path.join(__dirname, ".env"));

const PORT = Number(process.env.PORT || 8787);
const ROOT = path.resolve(process.env.PIPELINE_ROOT || path.join(__dirname, "pipeline"));
const DATA = path.join(__dirname, "data");
const JOBS = path.join(DATA, "jobs");
const PUBLIC = path.join(__dirname, "public");
const FACTOR = process.env.PIPELINE_FACTOR || "2.02";
const MAX_UPLOAD_MB = Number(process.env.MAX_UPLOAD_MB || 80);
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-5.2";
const OPENAI_MAX_OUTPUT_TOKENS = Number(process.env.OPENAI_MAX_OUTPUT_TOKENS || 12000);

const jobs = new Map();

const mime = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".hwpx": "application/octet-stream",
  ".pdf": "application/pdf"
};

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

async function main() {
  await fsp.mkdir(JOBS, { recursive: true });
  const server = http.createServer(handle);
  server.listen(PORT, () => {
    console.log(`HWP Conversion SaaS listening on http://localhost:${PORT}`);
    console.log(`Pipeline root: ${ROOT}`);
  });
}

async function handle(req, res) {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);

    if (req.method === "GET" && url.pathname === "/api/health") {
      return sendJson(res, 200, {
        ok: true,
        pipelineRoot: ROOT,
        factor: FACTOR,
        maxUploadMb: MAX_UPLOAD_MB,
        openaiConfigured: Boolean(process.env.OPENAI_API_KEY),
        openaiModel: OPENAI_MODEL
      });
    }

    if (req.method === "POST" && url.pathname === "/api/jobs") {
      return createJob(req, res);
    }

    const match = url.pathname.match(/^\/api\/jobs\/([^/]+)$/);
    if (req.method === "GET" && match) {
      return getJob(match[1], res);
    }

    const fileMatch = url.pathname.match(/^\/api\/jobs\/([^/]+)\/files\/(.+)$/);
    if (req.method === "GET" && fileMatch) {
      return sendJobFile(fileMatch[1], decodeURIComponent(fileMatch[2]), res);
    }

    if (req.method === "GET") {
      return sendStatic(url.pathname, res);
    }

    sendJson(res, 404, { error: "Not found" });
  } catch (error) {
    sendJson(res, 500, { error: error.message });
  }
}

async function createJob(req, res) {
  const contentType = req.headers["content-type"] || "";
  const boundary = contentType.match(/boundary=(.+)$/)?.[1];
  if (!boundary) return sendJson(res, 400, { error: "multipart/form-data required" });

  const body = await readRequest(req, MAX_UPLOAD_MB * 1024 * 1024);
  const parts = parseMultipart(body, boundary);
  const directDocx = parts.find((part) => part.name === "docx" && part.filename);
  const templateHwpx = parts.find((part) => part.name === "template" && part.filename);
  const brief = fieldText(parts, "brief");
  const reference = fieldText(parts, "reference");

  if (directDocx && !directDocx.filename.toLowerCase().endsWith(".docx")) {
    return sendJson(res, 400, { error: "DOCX file required for direct conversion." });
  }
  if (!directDocx && (!templateHwpx || !templateHwpx.filename.toLowerCase().endsWith(".hwpx"))) {
    return sendJson(res, 400, { error: "HWPX template is required when GPT draft mode is used." });
  }
  if (!directDocx && brief.trim().length < 10) {
    return sendJson(res, 400, { error: "Writing brief is required when GPT draft mode is used." });
  }

  const jobId = newJobId();
  const dir = path.join(JOBS, jobId);
  await fsp.mkdir(dir, { recursive: true });

  const job = {
    id: jobId,
    dir,
    source: path.join(dir, "source.docx"),
    template: templateHwpx ? path.join(dir, "template.hwpx") : null,
    mode: directDocx ? "direct-docx" : "gpt-template",
    brief,
    reference,
    status: "queued",
    log: "",
    averageSimilarity: null,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };

  if (directDocx) {
    await fsp.writeFile(job.source, directDocx.data);
  }
  if (templateHwpx) {
    await fsp.writeFile(job.template, templateHwpx.data);
  }
  if (brief) await fsp.writeFile(path.join(dir, "brief.txt"), brief, "utf8");
  if (reference) await fsp.writeFile(path.join(dir, "reference.txt"), reference, "utf8");

  jobs.set(jobId, job);
  runPipeline(job).catch((error) => {
    job.status = "failed";
    job.log += `\n${error.stack || error.message}`;
    touch(job);
  });

  sendJson(res, 202, { jobId });
}

function getJob(jobId, res) {
  const job = jobs.get(jobId);
  if (!job) return sendJson(res, 404, { error: "Job not found" });
  sendJson(res, 200, serializeJob(job));
}

async function sendJobFile(jobId, fileName, res) {
  const job = jobs.get(jobId);
  if (!job) return sendJson(res, 404, { error: "Job not found" });

  const safeName = path.basename(fileName);
  const file = path.join(job.dir, safeName);
  if (!file.startsWith(job.dir) || !fs.existsSync(file)) {
    return sendJson(res, 404, { error: "File not found" });
  }
  sendFile(file, res);
}

async function sendStatic(requestPath, res) {
  const clean = requestPath === "/" ? "/index.html" : requestPath;
  const target = path.normalize(path.join(PUBLIC, clean));
  if (!target.startsWith(PUBLIC) || !fs.existsSync(target)) {
    return sendJson(res, 404, { error: "Not found" });
  }
  sendFile(target, res);
}

function sendFile(file, res) {
  const ext = path.extname(file).toLowerCase();
  res.writeHead(200, { "Content-Type": mime[ext] || "application/octet-stream" });
  fs.createReadStream(file).pipe(res);
}

function sendJson(res, status, data) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(data));
}

function serializeJob(job) {
  const files = [];
  for (const [label, name] of [
    ["Generated DOCX", "source.docx"],
    ["Generated draft text", "generated.md"],
    ["Template text", "template_text.txt"],
    ["Final HWPX", "final.hwpx"],
    ["Similarity CSV", "page_similarity.csv"],
    ["Similarity JSON", "page_similarity.json"],
    ["Source PDF", "source.docx.pdf"],
    ["Candidate PDF", "candidate.hwpx.pdf"]
  ]) {
    if (fs.existsSync(path.join(job.dir, name))) {
      files.push({ label, href: `/api/jobs/${job.id}/files/${encodeURIComponent(name)}` });
    }
  }
  return {
    id: job.id,
    mode: job.mode,
    status: job.status,
    log: job.log,
    averageSimilarity: job.averageSimilarity,
    createdAt: job.createdAt,
    updatedAt: job.updatedAt,
    files
  };
}

async function runPipeline(job) {
  job.status = "running";
  touch(job);

  if (job.mode === "gpt-template") {
    await prepareGptDraft(job);
  }

  const rich = path.join(job.dir, "word_rich_paste.hwpx");
  const stretched = path.join(job.dir, "stretched.hwpx");
  const final = path.join(job.dir, "final.hwpx");

  await runPowerShell(job, path.join(ROOT, "paste_word_to_new_hangul.ps1"), [
    "-SourceDocx", job.source,
    "-OutputHwpx", rich
  ]);

  await runPowerShell(job, path.join(ROOT, "stretch_hwpx_layout.ps1"), [
    "-InputHwpx", rich,
    "-OutputHwpx", stretched,
    "-Factor", FACTOR
  ], { sta: false });

  await runPowerShell(job, path.join(ROOT, "reflow_with_hangul_engine.ps1"), [
    "-InputHwpx", stretched,
    "-OutputHwpx", final
  ]);

  await runPowerShell(job, path.join(ROOT, "compare_rendered_pages.ps1"), [
    "-SourceDocx", job.source,
    "-CandidateHwpx", final,
    "-WorkDir", path.join(job.dir, "compare")
  ], { allowFailure: true });

  await copyIfExists(path.join(job.dir, "compare", "page_similarity.csv"), path.join(job.dir, "page_similarity.csv"));
  await copyIfExists(path.join(job.dir, "compare", "page_similarity.json"), path.join(job.dir, "page_similarity.json"));
  await copyIfExists(path.join(job.dir, "compare", "source.docx.pdf"), path.join(job.dir, "source.docx.pdf"));
  await copyIfExists(path.join(job.dir, "compare", "candidate.hwpx.pdf"), path.join(job.dir, "candidate.hwpx.pdf"));

  job.averageSimilarity = await readAverageSimilarity(path.join(job.dir, "page_similarity.json"));
  job.status = "done";
  touch(job);
}

async function prepareGptDraft(job) {
  job.log += "\nPreparing GPT draft from uploaded HWPX template.\n";
  touch(job);

  const templateText = path.join(job.dir, "template_text.txt");
  await runPowerShell(job, path.join(ROOT, "extract_hwpx_text.ps1"), [
    "-InputHwpx", job.template,
    "-OutputText", templateText
  ], { sta: false, allowFailure: true });

  const extracted = fs.existsSync(templateText)
    ? await fsp.readFile(templateText, "utf8")
    : "";

  const generated = await generateDraftWithOpenAI(job, extracted);
  const generatedPath = path.join(job.dir, "generated.md");
  await fsp.writeFile(generatedPath, generated, "utf8");

  await runPowerShell(job, path.join(ROOT, "create_docx_from_text.ps1"), [
    "-InputText", generatedPath,
    "-OutputDocx", job.source
  ]);
}

async function generateDraftWithOpenAI(job, templateText) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error("OPENAI_API_KEY is not set. Set it before using GPT draft mode.");
  }

  job.log += `\nCalling OpenAI Responses API with model ${OPENAI_MODEL}.\n`;
  touch(job);

  const input = [
    "작성할 문서의 한글 양식에서 추출한 텍스트:",
    truncate(templateText, 18000) || "(양식 텍스트 추출 실패 또는 비어 있음)",
    "",
    "사용자 작성 지시:",
    job.brief,
    "",
    "추가 참고자료:",
    job.reference || "(없음)"
  ].join("\n");

  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: OPENAI_MODEL,
      instructions: [
        "You draft Korean government proposal documents.",
        "Use the uploaded Hangul template text as the structural guide.",
        "Return only clean Markdown suitable for conversion into a Word document.",
        "Use #, ##, and ### headings. Do not wrap the answer in code fences.",
        "Write in formal Korean business style. Preserve section order when it is clear."
      ].join(" "),
      input,
      max_output_tokens: OPENAI_MAX_OUTPUT_TOKENS
    })
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(`OpenAI request failed: ${payload.error?.message || response.statusText}`);
  }

  const text = payload.output_text || extractOutputText(payload);
  if (!text || text.trim().length < 20) {
    throw new Error("OpenAI response did not contain enough draft text.");
  }
  return text.trim();
}

async function runPowerShell(job, script, args, options = {}) {
  if (!fs.existsSync(script)) throw new Error(`Missing script: ${script}`);
  const psArgs = [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    ...(options.sta === false ? [] : ["-STA"]),
    "-File",
    script,
    ...args
  ];
  job.log += `\n> powershell ${psArgs.map(quote).join(" ")}\n`;
  touch(job);

  await new Promise((resolve, reject) => {
    const child = spawn("powershell.exe", psArgs, { cwd: ROOT, windowsHide: false });
    child.stdout.on("data", (data) => {
      job.log += data.toString();
      touch(job);
    });
    child.stderr.on("data", (data) => {
      job.log += data.toString();
      touch(job);
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0 && !options.allowFailure) {
        reject(new Error(`${path.basename(script)} exited with ${code}`));
      } else {
        resolve();
      }
    });
  });
}

async function copyIfExists(from, to) {
  if (fs.existsSync(from)) await fsp.copyFile(from, to);
}

async function readAverageSimilarity(jsonPath) {
  if (!fs.existsSync(jsonPath)) return null;
  const rows = JSON.parse(await fsp.readFile(jsonPath, "utf8"));
  if (!Array.isArray(rows) || rows.length === 0) return null;
  return rows.reduce((sum, row) => sum + Number(row.Similarity || row.similarity || 0), 0) / rows.length;
}

function readRequest(req, limit) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > limit) {
        reject(new Error("Upload too large"));
        req.destroy();
      } else {
        chunks.push(chunk);
      }
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function parseMultipart(buffer, boundary) {
  const delimiter = Buffer.from(`--${boundary}`);
  const parts = [];
  let start = buffer.indexOf(delimiter);
  while (start !== -1) {
    start += delimiter.length;
    if (buffer[start] === 45 && buffer[start + 1] === 45) break;
    if (buffer[start] === 13 && buffer[start + 1] === 10) start += 2;
    const next = buffer.indexOf(delimiter, start);
    if (next === -1) break;
    const part = buffer.slice(start, next - 2);
    const headerEnd = part.indexOf(Buffer.from("\r\n\r\n"));
    if (headerEnd !== -1) {
      const headers = part.slice(0, headerEnd).toString("utf8");
      const data = part.slice(headerEnd + 4);
      const disposition = headers.match(/content-disposition:[^\r\n]+/i)?.[0] || "";
      const name = disposition.match(/name="([^"]+)"/)?.[1];
      const filename = disposition.match(/filename="([^"]*)"/)?.[1] || "";
      parts.push({ name, filename, data });
    }
    start = next;
  }
  return parts;
}

function fieldText(parts, name) {
  const part = parts.find((item) => item.name === name && !item.filename);
  return part ? part.data.toString("utf8").trim() : "";
}

function quote(value) {
  return /\s/.test(String(value)) ? `"${String(value).replaceAll('"', '\\"')}"` : String(value);
}

function newJobId() {
  return `${new Date().toISOString().replace(/[-:.TZ]/g, "")}-${crypto.randomBytes(4).toString("hex")}`;
}

function touch(job) {
  job.updatedAt = new Date().toISOString();
}

function truncate(value, maxLength) {
  if (!value || value.length <= maxLength) return value || "";
  return `${value.slice(0, maxLength)}\n\n[...truncated...]`;
}

function extractOutputText(payload) {
  const chunks = [];
  for (const item of payload.output || []) {
    for (const content of item.content || []) {
      if (content.type === "output_text" && content.text) chunks.push(content.text);
    }
  }
  return chunks.join("\n");
}

function loadEnvFile(file) {
  if (!fs.existsSync(file)) return;
  const raw = fs.readFileSync(file, "utf8");
  for (const line of raw.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const index = trimmed.indexOf("=");
    if (index === -1) continue;
    const key = trimmed.slice(0, index).trim();
    const value = trimmed.slice(index + 1).trim().replace(/^["']|["']$/g, "");
    if (key && process.env[key] == null) process.env[key] = value;
  }
}

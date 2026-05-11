const form = document.querySelector("#uploadForm");
const modeInputs = document.querySelectorAll("input[name='mode']");
const gptFields = document.querySelector("#gptFields");
const docxFields = document.querySelector("#docxFields");
const templateInput = document.querySelector("#templateInput");
const templateName = document.querySelector("#templateName");
const docxInput = document.querySelector("#docxInput");
const docxName = document.querySelector("#docxName");
const briefInput = document.querySelector("#briefInput");
const referenceInput = document.querySelector("#referenceInput");
const statusCard = document.querySelector("#statusCard");
const jobStatus = document.querySelector("#jobStatus");
const jobScore = document.querySelector("#jobScore");
const jobLog = document.querySelector("#jobLog");
const downloads = document.querySelector("#downloads");

let pollTimer = null;

templateInput.addEventListener("change", () => {
  templateName.textContent = templateInput.files[0]?.name || "선택된 파일 없음";
});

docxInput.addEventListener("change", () => {
  docxName.textContent = docxInput.files[0]?.name || "선택된 파일 없음";
});

for (const input of modeInputs) {
  input.addEventListener("change", syncMode);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const mode = selectedMode();
  if (mode === "gpt" && !templateInput.files[0]) {
    alert("작성할 한글 HWPX 양식을 먼저 선택해 주세요.");
    return;
  }
  if (mode === "gpt" && briefInput.value.trim().length < 10) {
    alert("작성 지시를 조금 더 구체적으로 입력해 주세요.");
    return;
  }
  if (mode === "docx" && !docxInput.files[0]) {
    alert("DOCX 파일을 선택해 주세요.");
    return;
  }

  const submitButton = form.querySelector("button");
  submitButton.disabled = true;
  statusCard.classList.remove("hidden");
  jobStatus.textContent = "업로드 중";
  jobScore.textContent = "-";
  jobLog.textContent = "";
  downloads.innerHTML = "";

  const body = new FormData();
  if (mode === "gpt") {
    body.append("template", templateInput.files[0]);
    body.append("brief", briefInput.value);
    body.append("reference", referenceInput.value);
  } else {
    body.append("docx", docxInput.files[0]);
  }

  try {
    const response = await fetch("/api/jobs", { method: "POST", body });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "업로드 실패");
    }
    poll(data.jobId, submitButton);
  } catch (error) {
    submitButton.disabled = false;
    jobStatus.textContent = "실패";
    jobLog.textContent = error.message;
  }
});

function syncMode() {
  const mode = selectedMode();
  gptFields.classList.toggle("hidden", mode !== "gpt");
  docxFields.classList.toggle("hidden", mode !== "docx");
}

function selectedMode() {
  return document.querySelector("input[name='mode']:checked").value;
}

async function poll(jobId, submitButton) {
  clearInterval(pollTimer);
  const tick = async () => {
    const response = await fetch(`/api/jobs/${jobId}`);
    const job = await response.json();
    jobStatus.textContent = job.status;
    jobScore.textContent = job.averageSimilarity == null ? "-" : `${Math.round(job.averageSimilarity * 1000) / 10}%`;
    jobLog.textContent = job.log || "";
    downloads.innerHTML = "";

    if (job.files) {
      for (const file of job.files) {
        const link = document.createElement("a");
        link.href = file.href;
        link.textContent = file.label;
        downloads.appendChild(link);
      }
    }

    if (job.status === "done" || job.status === "failed") {
      clearInterval(pollTimer);
      submitButton.disabled = false;
    }
  };

  await tick();
  pollTimer = setInterval(tick, 3000);
}

syncMode();

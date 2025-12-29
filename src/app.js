// server.js
const express = require("express");
const cors = require("cors");
const { exec } = require("child_process");
const fs = require("fs");
const path = require("path");
const os = require("os");

const app = express();
const PORT = 4310;

const isPkg = typeof process.pkg !== "undefined";

const baseDir = isPkg ? path.dirname(process.execPath) : __dirname;

const SUMATRA_PATH = path.join(baseDir, "SumatraPDF.exe");
app.use(express.json({ limit: "50mb" }));

app.use(
  cors({
    origin: [
      "https://anniecong.o-r.kr", 
      "http://localhost:5173", 
      "http://localhost:3000",
      "https://anniecong.o-r.kr/api",
      "https://223.130.143.87:4000",
      "http://13.219.181.99:3000",
    ],
    methods: ["GET", "POST", "OPTIONS"],
  })
);

app.get("/health", (req, res) => {
  res.json({ ok: true, message: "Local printer agent is running" });
});

let printerCache = {
  data: [],
  fetchedAt: 0,
};
const PRINTER_CACHE_TTL = 10 * 1000; 

app.get("/printers", (req, res) => {
  const now = Date.now();
  const age = now - printerCache.fetchedAt;

  if (printerCache.data.length > 0 && age < PRINTER_CACHE_TTL) {
    return res.status(200).json({
      ok: true,
      message: `프린터 목록 조회 성공 (${printerCache.data.length}개)`,
      data: printerCache.data,
    });
  }

  console.log("프린터 목록 조회 요청 수신 (PowerShell 실행)");

  const psCommand =
    "Get-Printer | Select-Object Name,DriverName,Default,PrinterStatus | ConvertTo-Json";

  exec(
    `powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "${psCommand}"`,
    (error, stdout, stderr) => {
      if (error) {
        console.error("PowerShell 실행 오류:", error);
        return res.status(500).json({
          ok: false,
          message: "PowerShell을 통한 프린터 조회 중 오류가 발생했습니다.",
          error: error.message,
          data: [],
        });
      }

      if (stderr && stderr.trim().length > 0) {
        console.error("PowerShell stderr:", stderr);
      }

      try {
        if (!stdout || stdout.trim().length === 0) {
          printerCache = { data: [], fetchedAt: Date.now() };
          return res.status(200).json({
            ok: true,
            message: "프린터가 발견되지 않았습니다.",
            data: [],
          });
        }

        let printers = JSON.parse(stdout);

        if (!Array.isArray(printers)) {
          printers = [printers];
        }

        const mapped = printers.map((p) => ({
          name: String(p.Name ?? "").trim(),
          driver: String(p.DriverName ?? "").trim(),
          isDefault: Boolean(p.Default ?? false),
          status: String(p.PrinterStatus ?? "").trim(),
        }));

        printerCache = {
          data: mapped,
          fetchedAt: Date.now(),
        };

        return res.status(200).json({
          ok: true,
          message: `프린터 목록 조회 성공 (${mapped.length}개)`,
          data: mapped,
        });
      } catch (jsonError) {
        console.error("JSON 파싱 오류:", jsonError, "stdout:", stdout);
        return res.status(500).json({
          ok: false,
          message: "프린터 정보 JSON 파싱에 실패했습니다.",
          error: jsonError.message,
          data: [],
        });
      }
    }
  );
});

let lastPrintJob = {
  id: null,
  printerName: null,
  copies: 0,
  status: "idle", 
  message: "",
  error: null,
  startedAt: null,
  finishedAt: null,
};
let jobCounter = 0;

app.get("/print-status", (req, res) => {
  res.json({
    ok: true,
    data: lastPrintJob,
  });
});

app.post("/print", (req, res) => {
  console.log("=== /print 라벨 프린트 요청 수신 ===");
  console.log("요청 바디:", req.body);

  const { pdfBase64, printerName, printCount = 1 } = req.body || {};

  if (typeof pdfBase64 === "undefined" || pdfBase64 === null) {
    return res.status(400).json({
      ok: false,
      message: "pdfBase64 값이 비어있습니다.",
    });
  }

  if (typeof printerName !== "string" || printerName.trim().length === 0) {
    return res.status(400).json({
      ok: false,
      message: "printerName 값이 필요합니다.",
    });
  }

  const copiesRaw = Number(printCount);
  const copies =
    Number.isFinite(copiesRaw) && copiesRaw > 0 ? Math.floor(copiesRaw) : 1;

  const jobId = ++jobCounter;
  lastPrintJob = {
    id: jobId,
    printerName,
    copies,
    status: "queued",
    message: "프린트 요청이 정상적으로 들어왔습니다.",
    error: null,
    startedAt: new Date().toISOString(),
    finishedAt: null,
  };

  console.log("- 프린터:", printerName);
  console.log("- 매수:", copies);

  if (!fs.existsSync(SUMATRA_PATH)) {
    console.error("❌ SUMATRA_PATH 위치에 실행 파일이 없습니다:", SUMATRA_PATH);
    lastPrintJob.status = "error";
    lastPrintJob.message = "SumatraPDF 실행 파일을 찾을 수 없습니다.";
    lastPrintJob.error = `경로 확인: ${SUMATRA_PATH}`;
    lastPrintJob.finishedAt = new Date().toISOString();

    return res.status(500).json({
      ok: false,
      message: `SumatraPDF 실행 파일을 찾을 수 없습니다. 경로를 확인해주세요: ${SUMATRA_PATH}`,
    });
  }

  let pdfBuffer;

  try {
    if (typeof pdfBase64 === "string") {
      const trimmed = pdfBase64.trim();
      const looksLikeCsv = /^[0-9]+(,[0-9]+)*$/.test(trimmed);
      if (looksLikeCsv) {
        const bytes = trimmed.split(",").map((n) => Number(n));
        pdfBuffer = Buffer.from(bytes);
      } else {
        pdfBuffer = Buffer.from(trimmed, "base64");
      }
    } else if (Array.isArray(pdfBase64)) {
      pdfBuffer = Buffer.from(pdfBase64.map((n) => Number(n)));
    } else if (
      typeof pdfBase64 === "object" &&
      pdfBase64 !== null &&
      pdfBase64.type === "Buffer" &&
      Array.isArray(pdfBase64.data)
    ) {
      pdfBuffer = Buffer.from(pdfBase64.data.map((n) => Number(n)));
    } else {
      lastPrintJob.status = "error";
      lastPrintJob.message = "지원하지 않는 pdfBase64 형식입니다.";
      lastPrintJob.error = "지원하지 않는 pdfBase64 형식";
      lastPrintJob.finishedAt = new Date().toISOString();

      return res.status(400).json({
        ok: false,
        message: "지원하지 않는 pdfBase64 형식입니다.",
      });
    }
  } catch (e) {
    console.error("pdfBase64 → Buffer 변환 오류:", e);
    lastPrintJob.status = "error";
    lastPrintJob.message = "pdfBase64 변환 중 오류";
    lastPrintJob.error = e.message;
    lastPrintJob.finishedAt = new Date().toISOString();

    return res.status(500).json({
      ok: false,
      message: "pdfBase64를 PDF 버퍼로 변환하는 중 오류가 발생했습니다.",
      error: e.message,
    });
  }

  if (!Buffer.isBuffer(pdfBuffer) || pdfBuffer.length === 0) {
    lastPrintJob.status = "error";
    lastPrintJob.message = "PDF 버퍼가 비어있거나 형식 오류";
    lastPrintJob.error = "빈 버퍼";
    lastPrintJob.finishedAt = new Date().toISOString();

    return res.status(500).json({
      ok: false,
      message: "PDF 버퍼가 비어있거나 형식이 올바르지 않습니다.",
    });
  }

  try {
    const tempDir = path.join(os.tmpdir(), "anniecong-labels");

    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    const fileName = `label-${Date.now()}.pdf`;
    const filePath = path.join(tempDir, fileName);

    fs.writeFileSync(filePath, pdfBuffer);
    console.log("📄 PDF 저장 완료:", filePath);

    lastPrintJob.status = "printing";
    lastPrintJob.message = "프린트를 시작합니다.";

    const command = `"${SUMATRA_PATH}" -print-to "${printerName}" -print-settings "copies=${copies}" -silent "${filePath}"`;

    console.log("실행할 SumatraPDF 명령:", command);

    exec(command, { windowsHide: true }, (error, stdout, stderr) => {
      if (error) {
        console.error("❌ Sumatra 프린트 오류:", error);
        if (stderr && stderr.trim().length > 0) {
          console.error("stderr:", stderr);
        }

        lastPrintJob.status = "error";
        lastPrintJob.message = "프린트 실행 중 오류가 발생했습니다.";
        lastPrintJob.error = error.message;
        lastPrintJob.finishedAt = new Date().toISOString();

        return res.status(500).json({
          ok: false,
          message: "SumatraPDF 프린트 실행 중 오류가 발생했습니다.",
          error: error.message,
        });
      }

      if (stderr && stderr.trim().length > 0) {
        console.error("Sumatra stderr:", stderr);
      }

      console.log("Sumatra stdout:", stdout);

      const successMsg = `프린트 ${copies}장이 요청 완료 되었습니다.`;

      lastPrintJob.status = "success";
      lastPrintJob.message = successMsg;
      lastPrintJob.error = null;
      lastPrintJob.finishedAt = new Date().toISOString();

      return res.json({
        ok: true,
        message: successMsg,
        printerName,
        copies,
        filePath,
      });
    });
  } catch (e) {
    console.error("PDF 처리/프린트 준비 중 오류:", e);

    lastPrintJob.status = "error";
    lastPrintJob.message = "PDF 처리 중 오류가 발생했습니다.";
    lastPrintJob.error = e.message;
    lastPrintJob.finishedAt = new Date().toISOString();

    return res.status(500).json({
      ok: false,
      message: "PDF 처리 중 오류가 발생했습니다.",
      error: e.message,
    });
  }
});

const UI_HTML = `
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title>Anniecong Local Printer Agent</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      margin: 0;
      padding: 16px;
      background: #0f172a;
      color: #e5e7eb;
    }
    h1 {
      font-size: 20px;
      margin-bottom: 4px;
    }
    .subtitle {
      font-size: 12px;
      color: #9ca3af;
      margin-bottom: 16px;
    }
    .grid {
      display: grid;
      grid-template-columns: 1.2fr 1fr;
      gap: 16px;
      align-items: flex-start;
    }
    .card {
      background: #111827;
      border-radius: 12px;
      padding: 14px 16px;
      border: 1px solid #1f2937;
      box-shadow: 0 10px 25px rgba(0,0,0,0.4);
    }
    .card h2 {
      font-size: 15px;
      margin: 0 0 8px 0;
    }
    .card small {
      color: #6b7280;
    }
    button {
      border: none;
      border-radius: 999px;
      padding: 6px 12px;
      font-size: 12px;
      cursor: pointer;
      background: #2563eb;
      color: white;
      display: inline-flex;
      align-items: center;
      gap: 4px;
    }
    button.secondary {
      background: #374151;
    }
    button:disabled {
      opacity: 0.5;
      cursor: default;
    }
    .list {
      margin-top: 8px;
      max-height: 260px;
      overflow-y: auto;
      font-size: 12px;
      border-top: 1px solid #1f2937;
      padding-top: 6px;
    }
    .printer-item {
      padding: 4px 2px;
      border-bottom: 1px solid #111827;
      display: flex;
      justify-content: space-between;
      gap: 8px;
    }
    .printer-name {
      font-weight: 500;
      color: #e5e7eb;
    }
    .printer-meta {
      color: #9ca3af;
      font-size: 11px;
    }
    .tag {
      font-size: 10px;
      padding: 2px 6px;
      border-radius: 999px;
      background: #1f2937;
      color: #9ca3af;
      margin-left: 4px;
    }
    .tag.default {
      background: #22c55e22;
      color: #4ade80;
      border: 1px solid #22c55e66;
    }
    .status-row {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-bottom: 6px;
      font-size: 13px;
    }
    .dot {
      width: 10px;
      height: 10px;
      border-radius: 999px;
    }
    .dot.idle { background: #4b5563; }
    .dot.queued { background: #fbbf24; }
    .dot.printing { background: #22c55e; }
    .dot.success { background: #22c55e; }
    .dot.error { background: #ef4444; }
    .status-label {
      font-weight: 500;
    }
    .status-message {
      font-size: 12px;
      color: #e5e7eb;
      margin-top: 4px;
    }
    .status-meta {
      font-size: 11px;
      color: #9ca3af;
      margin-top: 4px;
    }
    .pill {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 2px 8px;
      border-radius: 999px;
      font-size: 11px;
      background: #111827;
      border: 1px solid #1f2937;
      color: #9ca3af;
    }
    .row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 8px;
    }
    .toast-area {
      margin-top: 8px;
      font-size: 11px;
      color: #9ca3af;
      min-height: 30px;
      white-space: pre-line;
    }
  </style>
</head>
<body>
  <h1>Anniecong Local Printer Agent</h1>
  <div class="subtitle">이 창은 exe 에이전트 상태를 보여주는 미니 대시보드입니다.</div>

  <div class="grid">
    <div class="card">
      <div class="row">
        <div>
          <h2>프린터 목록</h2>
          <small>로컬 PC에 설치된 프린터를 조회합니다.</small>
        </div>
        <div>
          <button id="refreshPrintersBtn">목록 새로고침</button>
        </div>
      </div>
      <div class="list" id="printersList">
        불러오는 중...
      </div>
    </div>

    <div class="card">
      <div class="row">
        <div>
          <h2>마지막 프린트 상태</h2>
          <small>/print 요청이 들어오면 여기에서 상태를 확인할 수 있습니다.</small>
        </div>
        <div>
          <button class="secondary" id="forceStatusBtn">상태 새로고침</button>
        </div>
      </div>
      <div id="statusArea">
        <div class="status-row">
          <div class="dot idle"></div>
          <div class="status-label">대기 중</div>
        </div>
        <div class="status-message">아직 프린트 요청이 없습니다.</div>
      </div>
      <div class="status-meta" id="statusMeta"></div>
      <div class="toast-area" id="toastArea"></div>
    </div>
  </div>

  <script>
    const printersListEl = document.getElementById("printersList");
    const refreshPrintersBtn = document.getElementById("refreshPrintersBtn");
    const statusAreaEl = document.getElementById("statusArea");
    const statusMetaEl = document.getElementById("statusMeta");
    const toastAreaEl = document.getElementById("toastArea");
    const forceStatusBtn = document.getElementById("forceStatusBtn");

    function setToast(message) {
      const now = new Date();
      const ts =
        now.getHours().toString().padStart(2, "0") +
        ":" +
        now.getMinutes().toString().padStart(2, "0") +
        ":" +
        now.getSeconds().toString().padStart(2, "0");
      toastAreaEl.textContent = "[" + ts + "] " + message;
    }

    async function loadPrinters() {
      printersListEl.textContent = "불러오는 중...";
      refreshPrintersBtn.disabled = true;
      try {
        const res = await fetch("/printers");
        const json = await res.json();
        if (!json.ok) {
          printersListEl.textContent = "프린터 목록 조회 중 오류가 발생했습니다.";
          setToast(json.message || "프린터 조회 실패");
          return;
        }
        const printers = json.data || [];
        if (printers.length === 0) {
          printersListEl.textContent = "발견된 프린터가 없습니다.";
          return;
        }

        printersListEl.innerHTML = "";
        printers.forEach(function (p) {
          const item = document.createElement("div");
          item.className = "printer-item";
          const left = document.createElement("div");
          const right = document.createElement("div");
          right.style.textAlign = "right";

          const nameEl = document.createElement("div");
          nameEl.className = "printer-name";
          nameEl.textContent = p.name || "(이름 없음)";

          const metaEl = document.createElement("div");
          metaEl.className = "printer-meta";
          metaEl.textContent = (p.driver || "") + (p.status ? " · 상태: " + p.status : "");

          left.appendChild(nameEl);
          left.appendChild(metaEl);

          if (p.isDefault) {
            const tag = document.createElement("span");
            tag.className = "tag default";
            tag.textContent = "기본 프린터";
            right.appendChild(tag);
          }

          item.appendChild(left);
          item.appendChild(right);
          printersListEl.appendChild(item);
        });

        setToast("프린터 목록을 " + printers.length + "개 불러왔습니다.");
      } catch (err) {
        console.error(err);
        printersListEl.textContent = "프린터 목록 조회 중 오류가 발생했습니다.";
        setToast("프린터 조회 중 에러 발생");
      } finally {
        refreshPrintersBtn.disabled = false;
      }
    }

    function renderStatus(job) {
      var status = job && job.status ? job.status : "idle";
      var printerName = job && job.printerName ? job.printerName : "-";
      var copies = job && job.copies ? job.copies : 0;
      var message = job && job.message ? job.message : "아직 프린트 요청이 없습니다.";

      var labelMap = {
        idle: "대기 중",
        queued: "요청 접수",
        printing: "프린트 중",
        success: "완료",
        error: "오류 발생"
      };

      var label = labelMap[status] || "대기 중";

      statusAreaEl.innerHTML =
        '<div class="status-row">' +
        '<div class="dot ' +
        status +
        '"></div>' +
        '<div class="status-label">' +
        label +
        "</div>" +
        "</div>" +
        '<div class="status-message">' +
        message +
        "</div>";

      var meta = [];
      if (printerName && printerName !== "-") {
        meta.push("프린터: " + printerName);
      }
      if (copies > 0) {
        meta.push("매수: " + copies + "장");
      }
      if (job && job.startedAt) {
        meta.push("시작: " + job.startedAt);
      }
      if (job && job.finishedAt) {
        meta.push("종료: " + job.finishedAt);
      }
      statusMetaEl.textContent = meta.join(" · ");
    }

    async function loadStatus() {
      try {
        const res = await fetch("/print-status");
        const json = await res.json();
        if (!json.ok) {
          return;
        }
        renderStatus(json.data);
      } catch (err) {
        console.error(err);
      }
    }

    refreshPrintersBtn.addEventListener("click", function () {
      loadPrinters();
    });

    forceStatusBtn.addEventListener("click", function () {
      loadStatus();
      setToast("프린트 상태를 수동으로 새로고침했습니다.");
    });

    window.addEventListener("load", function () {
      loadPrinters();
      loadStatus();
      setInterval(loadStatus, 1000);
    });
  </script>
</body>
</html>
`;

app.get("/ui", (req, res) => {
  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.send(UI_HTML);
});

app.listen(PORT, () => {
  console.log(`✅ Local printer agent listening on http://localhost:${PORT}`);

  const url = `http://localhost:${PORT}/ui`;
  try {
    exec(`start "" "${url}"`, { windowsHide: true });
  } catch (e) {
    console.error("브라우저 자동 오픈 실패:", e.message);
  }
});

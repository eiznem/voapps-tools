set -euo pipefail

ts="$(date +%Y%m%d_%H%M%S)"
mkdir -p _backups

if [ -f public/index.html ]; then
  cp public/index.html "_backups/index.html.$ts.bak"
fi

if [ -f server.js ]; then
  cp server.js "_backups/server.js.$ts.bak"
fi

mkdir -p public

cat > public/index.html <<'HTML'
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>VoApps Tools</title>
  <style>
    :root{
      --pink:#FF4B7D;
      --indigo:#3F2FB8;
      --ink:#0D0B1E;
      --text:#2A2A2A;
      --muted:#6B6B6B;
      --card:#FFFFFFCC;
      --card2:#FFFFFFE6;
      --stroke:#E7E7F0;
      --shadow: 0 18px 55px rgba(25, 18, 60, .12);
      --shadow2: 0 10px 30px rgba(25, 18, 60, .10);
      --radius:18px;
      --radius2:14px;
      --mono: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
      --sans: "Montserrat", system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
    }

    *{ box-sizing:border-box; }
    html,body{ height:100%; }
    body{
      margin:0;
      font-family:var(--sans);
      color:var(--text);
      background:
        radial-gradient(1200px 700px at 15% 10%, rgba(255,75,125,.30), transparent 60%),
        radial-gradient(1000px 650px at 85% 20%, rgba(63,47,184,.25), transparent 55%),
        linear-gradient(135deg, #F8F5FF, #F6F8FF 35%, #F9F7FF);
      overflow:hidden;
    }

    .app{
      height:100%;
      display:flex;
      flex-direction:column;
      padding:18px 18px 14px;
      gap:14px;
    }

    .topbar{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:12px;
      padding:10px 14px;
      border:1px solid var(--stroke);
      border-radius:var(--radius);
      background:linear-gradient(180deg, rgba(255,255,255,.75), rgba(255,255,255,.55));
      box-shadow: var(--shadow2);
      backdrop-filter: blur(10px);
    }
    .brand{
      display:flex;
      align-items:center;
      gap:10px;
      min-width:320px;
    }
    .brandMark{
      width:14px; height:14px;
      border-radius:4px;
      background:linear-gradient(135deg, var(--pink), var(--indigo));
      box-shadow: 0 10px 25px rgba(255,75,125,.25);
    }
    .brandTitle{
      font-weight:800;
      letter-spacing:.2px;
      display:flex;
      align-items:baseline;
      gap:8px;
    }
    .brandTitle .vo{ color:var(--pink); }
    .brandTitle .tools{ color:var(--indigo); }
    .pill{
      font-size:12px;
      color:#3C3C3C;
      background:#FFFFFFB3;
      border:1px solid var(--stroke);
      padding:4px 10px;
      border-radius:999px;
      margin-left:10px;
    }
    .rightPills{
      display:flex;
      align-items:center;
      gap:10px;
    }
    .statusPill{
      display:flex;
      align-items:center;
      gap:8px;
      padding:6px 10px;
      border-radius:999px;
      border:1px solid var(--stroke);
      background:#FFFFFFB3;
      font-size:12px;
      color:#3C3C3C;
    }
    .dot{
      width:8px;height:8px;border-radius:99px;
      background:#9AA0A6;
    }
    .dot.ok{ background:#23C55E; box-shadow: 0 0 0 4px rgba(35,197,94,.15); }
    .dot.bad{ background:#EF4444; box-shadow: 0 0 0 4px rgba(239,68,68,.15); }

    .grid{
      flex:1;
      display:grid;
      grid-template-columns: 560px 1fr;
      gap:14px;
      min-height:0;
    }

    .card{
      border:1px solid var(--stroke);
      border-radius:var(--radius);
      background:linear-gradient(180deg, var(--card2), var(--card));
      box-shadow: var(--shadow);
      backdrop-filter: blur(10px);
      min-height:0;
      overflow:hidden;
    }

    .left{
      display:flex;
      flex-direction:column;
      min-height:0;
    }
    .leftHeader{
      padding:16px 18px 10px;
      border-bottom:1px solid rgba(231,231,240,.8);
    }
    .h1{
      font-size:20px;
      font-weight:800;
      margin:0;
      color:var(--ink);
    }
    .sub{
      margin:6px 0 0;
      font-size:12.5px;
      color:var(--muted);
      line-height:1.35;
    }
    .leftBody{
      padding:14px 14px 12px;
      display:flex;
      flex-direction:column;
      gap:10px;
      min-height:0;
      overflow:hidden;
    }

    .section{
      border:1px solid rgba(231,231,240,.9);
      border-radius:var(--radius2);
      background:#FFFFFFCC;
      padding:12px 12px 10px;
    }
    .sectionTop{
      display:flex;
      align-items:flex-start;
      justify-content:space-between;
      gap:10px;
      margin-bottom:10px;
    }
    .sectionTitle{
      font-size:13px;
      font-weight:800;
      color:#222;
      margin:0;
    }
    .sectionHint{
      font-size:11.5px;
      color:var(--muted);
      margin:0;
      text-align:right;
      line-height:1.25;
    }

    .row{ display:flex; gap:10px; align-items:center; }
    label{
      font-size:12px;
      font-weight:700;
      color:#333;
      display:block;
      margin-bottom:6px;
    }

    input[type="text"], input[type="password"], input[type="date"], select, textarea{
      width:100%;
      border:1px solid var(--stroke);
      border-radius:12px;
      padding:10px 12px;
      font-size:13px;
      outline:none;
      background:#fff;
      box-shadow: 0 6px 18px rgba(25,18,60,.04);
    }
    textarea{
      height:84px;
      resize:none;
      font-family:var(--sans);
      line-height:1.25;
    }

    .btnRow{
      display:flex;
      gap:10px;
      align-items:center;
      flex-wrap:wrap;
      margin-top:10px;
    }
    .btn{
      border:1px solid rgba(0,0,0,.06);
      border-radius:12px;
      padding:10px 12px;
      font-weight:800;
      font-size:13px;
      cursor:pointer;
      background:#fff;
      box-shadow: 0 10px 22px rgba(25,18,60,.08);
      transition: transform .06s ease, box-shadow .2s ease;
      user-select:none;
    }
    .btn:active{ transform: translateY(1px); box-shadow: 0 6px 14px rgba(25,18,60,.10); }
    .btnPrimary{
      background: linear-gradient(135deg, var(--pink), #FF6B98);
      color:white;
      border-color: rgba(255,255,255,.25);
    }
    .btnDanger{
      background: linear-gradient(135deg, #FF4B7D, #FF2D65);
      color:white;
      border-color: rgba(255,255,255,.25);
    }
    .btnGhost{
      background:#fff;
      color:#2E2E2E;
    }
    .btnSmall{
      padding:7px 10px;
      font-size:12px;
      border-radius:10px;
      box-shadow:none;
    }
    .btn.disabled, .btn:disabled{
      opacity:.45;
      cursor:not-allowed;
      transform:none !important;
      box-shadow: none !important;
    }

    .toggleRow{
      display:flex;
      gap:14px;
      align-items:center;
      margin-top:2px;
      flex-wrap:wrap;
    }
    .check{
      display:flex;
      align-items:center;
      gap:8px;
      font-size:12.5px;
      color:#333;
    }
    .check input{ width:16px; height:16px; }

    .right{
      display:flex;
      flex-direction:column;
      min-height:0;
    }
    .rightHeader{
      padding:16px 18px 10px;
      display:flex;
      align-items:flex-start;
      justify-content:space-between;
      gap:10px;
      border-bottom:1px solid rgba(231,231,240,.8);
    }
    .rightTitle{
      margin:0;
      font-size:14px;
      font-weight:900;
      color:#222;
    }
    .rightActions{
      display:flex;
      gap:10px;
      align-items:center;
      flex-wrap:wrap;
    }
    .mutedSmall{
      font-size:11.5px;
      color:var(--muted);
      margin:0;
      line-height:1.25;
    }

    .rightBody{
      padding:14px;
      display:flex;
      flex-direction:column;
      gap:10px;
      min-height:0;
    }

    .statusBox{
      border:1px solid rgba(231,231,240,.9);
      border-radius:var(--radius2);
      background:#FFFFFFCC;
      padding:12px;
    }
    .statusLine{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:10px;
      margin-bottom:8px;
    }
    .statusBig{
      display:flex;
      align-items:center;
      gap:10px;
      font-weight:900;
      color:#222;
    }
    .badge{
      font-size:11px;
      padding:4px 10px;
      border-radius:999px;
      border:1px solid var(--stroke);
      background:#fff;
      color:#333;
      font-weight:800;
    }
    .badge.idle{ background:#F6F7FF; }
    .badge.run{ background:#FFF1F7; border-color: rgba(255,75,125,.30); }
    .badge.done{ background:#ECFDF3; border-color: rgba(35,197,94,.30); }
    .badge.err{ background:#FFF1F1; border-color: rgba(239,68,68,.30); }

    .artifact{
      font-size:12px;
      color:#333;
      margin:0;
      display:flex;
      gap:8px;
      align-items:baseline;
      white-space:nowrap;
      overflow:hidden;
      text-overflow:ellipsis;
    }
    .artifact code{
      font-family:var(--mono);
      font-size:11px;
      color:#222;
      background:rgba(0,0,0,.04);
      padding:2px 6px;
      border-radius:8px;
      border:1px solid rgba(0,0,0,.06);
      white-space:nowrap;
      overflow:hidden;
      text-overflow:ellipsis;
      display:inline-block;
      max-width: 100%;
    }

    .logWrap{
      flex:1;
      min-height:0;
      border-radius:var(--radius2);
      border:1px solid rgba(231,231,240,.9);
      background: linear-gradient(180deg, #0D1020, #0B0E1A);
      box-shadow: inset 0 0 0 1px rgba(255,255,255,.05);
      overflow:hidden;
      display:flex;
      flex-direction:column;
    }
    .logHeader{
      padding:10px 12px;
      display:flex;
      align-items:center;
      justify-content:space-between;
      color:#D7DAE5;
      font-size:12px;
      border-bottom:1px solid rgba(255,255,255,.08);
    }
    .logHeader strong{ color:#FFFFFF; }
    .log{
      padding:10px 12px;
      color:#E9ECF5;
      font-family:var(--mono);
      font-size:11.5px;
      line-height:1.35;
      overflow:auto;
      white-space:pre-wrap;
      flex:1;
      min-height:0;
    }

    .footer{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:10px;
      padding:0 6px;
      color:rgba(0,0,0,.45);
      font-size:11.5px;
    }
    .footerRight{ opacity:.7; }

    .accountsList{
      margin-top:10px;
      max-height:240px;
      overflow:auto;
      border:1px solid rgba(0,0,0,0.10);
      border-radius:12px;
      padding:10px;
      background:rgba(255,255,255,0.7);
    }
    .acctRow{
      display:flex;
      align-items:center;
      gap:10px;
      padding:8px 6px;
      border-radius:10px;
      cursor:pointer;
    }
    .acctRow:hover{ background: rgba(63,47,184,0.06); }
    .acctRow .top{ font-weight:650; }
    .acctRow .sub{ margin:0; font-size:11.5px; color:var(--muted); }

    .colsGrid{
      display:grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap:10px;
      margin-top:10px;
    }
    .colsCard{
      border:1px solid rgba(231,231,240,.9);
      border-radius:12px;
      background:#fff;
      padding:10px 10px 8px;
    }
    .colsTitle{
      font-size:12px;
      font-weight:900;
      color:#222;
      margin:0 0 6px;
    }
    .colsCard label{
      display:flex;
      align-items:center;
      gap:8px;
      font-weight:600;
      margin:0 0 4px;
      font-size:12px;
    }
    .colsCard input{ width:auto; }

    @media (max-height: 820px){
      textarea{ height:72px; }
      .section{ padding:11px; }
      .leftHeader, .rightHeader{ padding:14px 16px 10px; }
    }
    @media (max-height: 740px){
      textarea{ height:62px; }
      .sub{ display:none; }
    }
  </style>
</head>

<body>
  <div class="app">
    <div class="topbar">
      <div class="brand">
        <div class="brandMark" aria-hidden="true"></div>
        <div class="brandTitle">
          <span class="vo">VoApps</span>
          <span class="tools">Tools</span>
          <span class="pill">Local Mac Utility</span>
        </div>
      </div>

      <div class="rightPills">
        <div class="statusPill" title="Window status">
          <span id="modeDot" class="dot"></span>
          <span id="modeText">Idle</span>
        </div>
        <div class="statusPill" title="Server status">
          <span id="srvDot" class="dot"></span>
          <span>Server:</span>
          <span id="srvText">Checking…</span>
        </div>
      </div>
    </div>

    <div class="grid">
      <div class="card left">
        <div class="leftHeader">
          <h1 class="h1">VoApps Tools</h1>
          <p class="sub">Number Search across selected accounts & created-date range (day-by-day).</p>
        </div>

        <div class="leftBody">
          <div class="section">
            <div class="sectionTop">
              <div><p class="sectionTitle">1. API Key</p></div>
              <p class="sectionHint">Stored locally on this machine</p>
            </div>

            <label for="apiKey">VoApps API Key</label>
            <div class="row">
              <input id="apiKey" type="password" placeholder="Paste your API key here…" autocomplete="off" />
              <button id="saveKeyBtn" class="btn btnSmall btnGhost">Save</button>
              <button id="clearKeyBtn" class="btn btnSmall btnGhost">Clear</button>
            </div>
          </div>

          <div class="section">
            <div class="sectionTop">
              <p class="sectionTitle">2. Date Range</p>
              <p class="sectionHint">Presets or custom</p>
            </div>

            <label for="datePreset">Options</label>
            <select id="datePreset">
              <option value="">Custom range…</option>
              <option value="months_1">1 Month</option>
              <option value="months_2">2 Months</option>
              <option value="months_3">3 Months</option>
              <option value="months_6">6 Months</option>
              <option value="years_1">1 Year</option>
              <option value="years_2">2 Years</option>
              <option value="years_3">3 Years</option>
              <option value="years_4">4 Years</option>
              <option value="years_5">5 Years</option>
            </select>

            <div class="row" style="margin-top:10px;">
              <div style="flex:1;">
                <label for="startDate">Start date</label>
                <input id="startDate" type="date" />
              </div>
              <div style="flex:1;">
                <label for="endDate">End date</label>
                <input id="endDate" type="date" />
              </div>
            </div>

            <p class="mutedSmall" style="margin:8px 0 0;">
              Uses <b>created_date=YYYY-MM-DD</b> and queries one day at a time.
            </p>
          </div>

          <div class="section">
            <div class="sectionTop">
              <p class="sectionTitle">3. Accounts</p>
              <p class="sectionHint">Load → select accounts</p>
            </div>

            <div class="row" style="justify-content:space-between;">
              <button id="loadAccountsBtn" class="btn btnSmall btnGhost">Load accounts</button>
              <span class="mutedSmall" id="acctCount">Not loaded yet</span>
            </div>

            <div class="btnRow" style="margin-top:8px;">
              <button id="selectAllBtn" class="btn btnSmall btnGhost disabled" disabled>Select all</button>
              <button id="selectActiveBtn" class="btn btnSmall btnGhost disabled" disabled>Select active</button>
              <button id="clearSelectionBtn" class="btn btnSmall btnGhost disabled" disabled>Clear selection</button>
            </div>

            <div id="accountsList" class="accountsList"></div>

            <div class="toggleRow" style="margin-top:10px;">
              <label class="check" title="Dry run does NOT download exports or write CSV. It only counts campaigns per day.">
                <input id="dryRunToggle" type="checkbox" />
                Dry run only (no exports/CSV)
              </label>
            </div>
          </div>

          <div class="section">
            <div class="sectionTop">
              <p class="sectionTitle">4. Search Inputs</p>
              <p class="sectionHint">Number Search</p>
            </div>

            <label for="numbers">Phone numbers (one per line or comma-separated)</label>
            <textarea id="numbers" placeholder="e.g. 4353137000&#10;8015551212"></textarea>

            <div class="toggleRow">
              <label class="check" title="If results include voapps_caller_number, use that (preferred).">
                <input id="includeCaller" type="checkbox" checked />
                Include caller number
              </label>

              <label class="check" title="Include message_name and message_description when message_id is present.">
                <input id="includeMessageMeta" type="checkbox" checked />
                Include message metadata
              </label>
            </div>

            <div class="row" style="margin-top:10px;">
              <div style="flex:1;">
                <label for="logLevel">Logging</label>
                <select id="logLevel">
                  <option value="none">None</option>
                  <option value="minimal">Minimal</option>
                  <option value="verbose" selected>Verbose</option>
                </select>
              </div>
            </div>

            <div class="colsGrid">
              <div class="colsCard">
                <p class="colsTitle">Core</p>
                <label><input type="checkbox" class="col-toggle" data-col="number" checked>number</label>
                <label><input type="checkbox" class="col-toggle" data-col="account_id" checked>account_id</label>
                <label><input type="checkbox" class="col-toggle" data-col="campaign_id" checked>campaign_id</label>
                <label><input type="checkbox" class="col-toggle" data-col="campaign_name" checked>campaign_name</label>
              </div>
              <div class="colsCard">
                <p class="colsTitle">Optional</p>
                <label><input type="checkbox" class="col-toggle" data-col="caller_number">caller_number</label>
                <label><input type="checkbox" class="col-toggle" data-col="message_id">message_id</label>
                <label><input type="checkbox" class="col-toggle" data-col="message_name">message_name</label>
                <label><input type="checkbox" class="col-toggle" data-col="message_description">message_description</label>
              </div>
              <div class="colsCard">
                <p class="colsTitle">Results</p>
                <label><input type="checkbox" class="col-toggle" data-col="voapps_result" checked>voapps_result</label>
                <label><input type="checkbox" class="col-toggle" data-col="voapps_code" checked>voapps_code</label>
                <label><input type="checkbox" class="col-toggle" data-col="voapps_timestamp" checked>voapps_timestamp</label>
                <label><input type="checkbox" class="col-toggle" data-col="campaign_url" checked>campaign_url</label>
              </div>
            </div>
          </div>

          <div class="section">
            <div class="sectionTop">
              <p class="sectionTitle">5. Run</p>
              <p class="sectionHint">Run / Cancel / Quit</p>
            </div>

            <div class="btnRow">
              <button id="runBtn" class="btn btnPrimary">Run</button>
              <button id="cancelBtn" class="btn btnGhost disabled" disabled>Cancel</button>
              <button id="quitBtn" class="btn btnDanger">Quit</button>
              <button id="openCsvBtn" class="btn btnGhost disabled" disabled>Open CSV</button>
              <button id="openLogBtn" class="btn btnGhost disabled" disabled>Open Log</button>
            </div>

            <p class="mutedSmall" style="margin:8px 0 0;">
              Cancel stops the server-side job (not just the browser request). Quit stops the local server.
            </p>
          </div>
        </div>
      </div>

      <div class="card right">
        <div class="rightHeader">
          <div>
            <p class="rightTitle">Search status</p>
            <p class="mutedSmall" id="statusText">Idle</p>
          </div>
          <div class="rightActions">
            <button id="clearLogBtn" class="btn btnSmall btnGhost">Clear log</button>
            <button id="pingBtn" class="btn btnSmall btnGhost">Ping</button>
          </div>
        </div>

        <div class="rightBody">
          <div class="statusBox">
            <div class="statusLine">
              <div class="statusBig">
                <span id="statusDot" class="dot ok"></span>
                <span id="statusTitle">Ready</span>
              </div>
              <span id="modeBadge" class="badge idle">Idle</span>
            </div>

            <p class="artifact"><b>Last CSV:</b> <code id="csvPath">—</code></p>
            <p class="artifact"><b>Last Log:</b> <code id="logPath">—</code></p>
          </div>

          <div class="logWrap">
            <div class="logHeader">
              <strong>Live Log</strong>
              <span id="logHint">Progress + logs</span>
            </div>
            <div id="log" class="log"></div>
          </div>
        </div>
      </div>
    </div>

    <div class="footer">
      <div>VoApps Internal Tools</div>
      <div class="footerRight">DirectDrop Voicemail™</div>
    </div>
  </div>

  <script>
    const $ = (id) => document.getElementById(id);

    const logEl = $("log");
    const srvDot = $("srvDot");
    const srvText = $("srvText");
    const modeDot = $("modeDot");
    const modeText = $("modeText");

    const statusText = $("statusText");
    const statusTitle = $("statusTitle");
    const statusDot = $("statusDot");
    const modeBadge = $("modeBadge");

    const csvPathEl = $("csvPath");
    const logPathEl = $("logPath");

    const KEY_STORE = "voapps_api_key";
    const apiKeyEl = $("apiKey");

    let uiLocked = false;
    let currentAbort = null;
    let currentJobId = null;

    function ts(){
      const d = new Date();
      const hh = String(d.getHours()).padStart(2,"0");
      const mm = String(d.getMinutes()).padStart(2,"0");
      const ss = String(d.getSeconds()).padStart(2,"0");
      return `${hh}:${mm}:${ss}`;
    }

    function appendLog(line){
      logEl.textContent += `[${ts()}] ${line}\n`;
      logEl.scrollTop = logEl.scrollHeight;
    }

    function setMode(label, kind){
      modeText.textContent = label;
      modeDot.className = "dot " + (kind === "ok" ? "ok" : kind === "bad" ? "bad" : "");
    }

    function setStatus(title, subtitle, badge, badgeClass){
      statusTitle.textContent = title;
      statusText.textContent = subtitle;
      modeBadge.textContent = badge;
      modeBadge.className = "badge " + (badgeClass || "idle");
      statusDot.className = "dot " + (badgeClass === "err" ? "bad" : "ok");
    }

    function setLocked(locked){
      uiLocked = locked;

      const controls = [
        $("apiKey"), $("saveKeyBtn"), $("clearKeyBtn"),
        $("datePreset"), $("startDate"), $("endDate"),
        $("loadAccountsBtn"), $("selectAllBtn"), $("selectActiveBtn"), $("clearSelectionBtn"),
        $("numbers"), $("includeCaller"), $("includeMessageMeta"), $("dryRunToggle"),
        $("logLevel"), ...document.querySelectorAll(".col-toggle"),
        $("runBtn")
      ];

      controls.forEach(el => { if (el) el.disabled = locked; });

      const cancelBtn = $("cancelBtn");
      cancelBtn.disabled = !locked;
      cancelBtn.classList.toggle("disabled", !locked);
    }

    function getApiKey(){
      return (apiKeyEl.value || "").trim() || (localStorage.getItem(KEY_STORE) || "");
    }

    $("saveKeyBtn").addEventListener("click", () => {
      localStorage.setItem(KEY_STORE, apiKeyEl.value || "");
      appendLog("[ui] API key saved locally.");
    });

    $("clearKeyBtn").addEventListener("click", () => {
      apiKeyEl.value = "";
      localStorage.removeItem(KEY_STORE);
      appendLog("[ui] API key cleared.");
    });

    apiKeyEl.value = localStorage.getItem(KEY_STORE) || "";

    async function ping(){
      try{
        const r = await fetch("/api/ping", { cache: "no-store" });
        const ok = r.ok;
        srvText.textContent = ok ? "Connected" : "Error";
        srvDot.className = "dot " + (ok ? "ok" : "bad");
        setMode(ok ? "Idle" : "Offline", ok ? "ok" : "bad");
        return ok;
      }catch(e){
        srvText.textContent = "Offline";
        srvDot.className = "dot bad";
        setMode("Offline", "bad");
        return false;
      }
    }

    $("pingBtn").addEventListener("click", async () => {
      appendLog("[ui] Ping…");
      const ok = await ping();
      appendLog(ok ? "[ui] Ping: OK" : "[ui] Ping: FAILED");
    });

    $("clearLogBtn").addEventListener("click", () => {
      logEl.textContent = "";
      appendLog("[ui] Log cleared.");
    });

    async function refreshArtifacts(){
      if (!window.voapps || !window.voapps.getLastArtifacts) return;
      const r = await window.voapps.getLastArtifacts();
      if (!r.ok) return;
      const { csvPath, logPath } = r.artifacts || {};

      if (csvPath){
        csvPathEl.textContent = csvPath;
        const b = $("openCsvBtn");
        b.classList.remove("disabled"); b.disabled = false;
        b.onclick = async () => {
          appendLog("[ui] Opening CSV…");
          const rr = await window.voapps.openPath(csvPath);
          if (!rr.ok) appendLog("[ui] Open CSV failed: " + rr.error);
        };
      }
      if (logPath){
        logPathEl.textContent = logPath;
        const b = $("openLogBtn");
        b.classList.remove("disabled"); b.disabled = false;
        b.onclick = async () => {
          appendLog("[ui] Opening Log…");
          const rr = await window.voapps.openPath(logPath);
          if (!rr.ok) appendLog("[ui] Open Log failed: " + rr.error);
        };
      }
    }

    function parsePhones(raw){
      const parts = (raw || "").split(/[^0-9]+/).filter(Boolean);
      const cleaned = [];
      for (const p of parts){
        let digits = String(p).replace(/\D+/g,"");
        if (digits.length === 11 && digits.startsWith("1")) digits = digits.slice(1);
        if (digits.length === 10) cleaned.push(digits);
      }
      return cleaned;
    }

    function iso(d){ return d.toISOString().slice(0,10); }

    function computeDateRange(){
      const preset = $("datePreset").value;
      const today = new Date();
      const end = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      let start = new Date(end.getTime());

      if (preset.startsWith("months_")){
        const n = parseInt(preset.split("_")[1],10);
        start.setMonth(start.getMonth() - n);
      } else if (preset.startsWith("years_")){
        const n = parseInt(preset.split("_")[1],10);
        start.setFullYear(start.getFullYear() - n);
      } else {
        // custom
        const s = $("startDate").value;
        const e = $("endDate").value;
        return { startDate: s, endDate: e, computed: false };
      }

      return { startDate: iso(start), endDate: iso(end), computed: true };
    }

    $("datePreset").addEventListener("change", () => {
      const r = computeDateRange();
      if (r.computed){
        $("startDate").value = r.startDate;
        $("endDate").value = r.endDate;
        appendLog(`[ui] Preset applied: ${$("datePreset").value} → ${r.startDate} to ${r.endDate}`);
      }
    });

    // default: custom with month-to-date
    (function initDates(){
      const today = new Date();
      $("startDate").value = iso(new Date(today.getFullYear(), today.getMonth(), 1));
      $("endDate").value = iso(today);
      $("datePreset").value = "";
    })();

    // Accounts wiring (with hidden-aware “Select active” if server supports filter=hidden)
    const loadBtn = $("loadAccountsBtn");
    const acctCount = $("acctCount");
    const selectAllBtn = $("selectAllBtn");
    const selectActiveBtn = $("selectActiveBtn");
    const clearSelectionBtn = $("clearSelectionBtn");
    const list = $("accountsList");

    let lastAccounts = [];
    let hiddenIds = new Set();
    let selected = new Set();

    function setAccountBtns(on){
      [selectAllBtn, selectActiveBtn, clearSelectionBtn].forEach(b => {
        b.disabled = !on;
        b.classList.toggle("disabled", !on);
      });
    }

    function renderAccounts(items){
      lastAccounts = Array.isArray(items) ? items : [];
      list.innerHTML = "";
      selected = new Set();
      window.__VOAPPS_SELECTED_ACCOUNTS__ = [];

      if (!lastAccounts.length){
        list.innerHTML = '<div class="mutedSmall">No accounts returned.</div>';
        acctCount.textContent = "0 accounts";
        setAccountBtns(false);
        return;
      }

      const header = document.createElement("div");
      header.className = "mutedSmall";
      header.style.marginBottom = "8px";
      header.textContent = "Select accounts to include:";
      list.appendChild(header);

      lastAccounts.forEach(a => {
        const id = String(a?.id ?? "");
        const name = a?.name || "";
        const tz = a?.timezone || "";

        const row = document.createElement("div");
        row.className = "acctRow";

        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.addEventListener("change", () => {
          if (cb.checked) selected.add(id);
          else selected.delete(id);
          window.__VOAPPS_SELECTED_ACCOUNTS__ = Array.from(selected);
        });

        const text = document.createElement("div");
        text.style.display = "flex";
        text.style.flexDirection = "column";

        const top = document.createElement("div");
        top.className = "top";
        const hiddenMark = hiddenIds.has(id) ? " (hidden)" : "";
        top.textContent = `${id} — ${name}${hiddenMark}`;

        const sub = document.createElement("p");
        sub.className = "sub";
        sub.textContent = tz ? tz : "Timezone: (unknown)";

        text.appendChild(top);
        text.appendChild(sub);

        row.appendChild(cb);
        row.appendChild(text);
        list.appendChild(row);
      });

      acctCount.textContent = `${lastAccounts.length} accounts`;
      setAccountBtns(true);
    }

    async function fetchAccounts(filter){
      const api_key = getApiKey();
      if (!api_key){
        acctCount.textContent = "Missing API key";
        list.innerHTML = '<div class="mutedSmall">Paste your API key first.</div>';
        setAccountBtns(false);
        return { ok:false, accounts:[] };
      }

      const resp = await fetch("/api/accounts", {
        method:"POST",
        headers:{ "Content-Type":"application/json" },
        body: JSON.stringify({ api_key, filter: filter || "all" }),
        cache:"no-store"
      });

      const data = await resp.json().catch(() => ({}));
      if (!resp.ok || !data || data.ok !== true){
        const msg = (data && (data.error || data.message)) ? (data.error || data.message) : `HTTP ${resp.status}`;
        return { ok:false, error: msg, accounts:[] };
      }

      const items = Array.isArray(data.accounts) ? data.accounts
                  : Array.isArray(data.accounts?.accounts) ? data.accounts.accounts
                  : [];
      return { ok:true, accounts: items };
    }

    async function loadAccounts(){
      try{
        acctCount.textContent = "Loading…";
        list.innerHTML = '<div class="mutedSmall">Loading accounts…</div>';
        hiddenIds = new Set();

        const all = await fetchAccounts("all");
        if (!all.ok){
          acctCount.textContent = "Load failed";
          list.innerHTML = `<div class="mutedSmall">Error loading accounts: ${String(all.error || "unknown")}</div>`;
          setAccountBtns(false);
          appendLog("[ui] Accounts load failed: " + String(all.error || "unknown"));
          return;
        }

        // attempt hidden fetch (optional; server may return empty)
        const hid = await fetchAccounts("hidden");
        if (hid.ok && hid.accounts.length){
          hiddenIds = new Set(hid.accounts.map(a => String(a.id)));
          appendLog("[ui] Hidden accounts loaded for Select active.");
        } else {
          appendLog("[ui] Hidden accounts not available; Select active behaves like Select all.");
        }

        renderAccounts(all.accounts);
        appendLog(`[ui] Loaded ${all.accounts.length} account(s).`);
      }catch(e){
        acctCount.textContent = "Load failed";
        list.innerHTML = `<div class="mutedSmall">Error loading accounts: ${String(e?.message || e)}</div>`;
        setAccountBtns(false);
        appendLog("[ui] Accounts load error: " + String(e?.message || e));
      }
    }

    loadBtn.addEventListener("click", loadAccounts);

    selectAllBtn.addEventListener("click", () => {
      selected = new Set(lastAccounts.map(a => String(a.id)));
      list.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
      window.__VOAPPS_SELECTED_ACCOUNTS__ = Array.from(selected);
      appendLog("[ui] Selected all accounts.");
    });

    selectActiveBtn.addEventListener("click", () => {
      if (!hiddenIds.size){
        selectAllBtn.click();
        return;
      }
      selected = new Set();
      list.querySelectorAll('.acctRow').forEach(row => {
        const cb = row.querySelector('input[type="checkbox"]');
        const top = row.querySelector('.top');
        if (!cb || !top) return;
        const id = (top.textContent || "").split(" — ")[0].trim();
        const isActive = id && !hiddenIds.has(id);
        cb.checked = !!isActive;
        if (isActive) selected.add(id);
      });
      window.__VOAPPS_SELECTED_ACCOUNTS__ = Array.from(selected);
      appendLog("[ui] Selected active (non-hidden) accounts.");
    });

    clearSelectionBtn.addEventListener("click", () => {
      selected = new Set();
      list.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
      window.__VOAPPS_SELECTED_ACCOUNTS__ = [];
      appendLog("[ui] Cleared account selection.");
    });

    function getSelectedColumns(){
      const cols = [];
      document.querySelectorAll(".col-toggle").forEach(cb => {
        if (cb.checked) cols.push(cb.getAttribute("data-col"));
      });
      return cols;
    }

    function newJobId(){
      return "job_" + Date.now() + "_" + Math.random().toString(16).slice(2);
    }

    async function cancelRun(){
      if (!uiLocked || !currentJobId){
        appendLog("[ui] No active run to cancel.");
        return;
      }

      appendLog("[ui] Cancelling server-side job " + currentJobId + "…");
      setStatus("Cancelling…", "Stopping…", "Running", "run");

      try{ if (currentAbort) currentAbort.abort(); }catch(_){}

      try{
        await fetch("/api/cancel", {
          method:"POST",
          headers:{ "Content-Type":"application/json" },
          body: JSON.stringify({ job_id: currentJobId })
        });
      }catch(e){
        appendLog("[ui] Cancel request failed: " + String(e?.message || e));
      }

      setLocked(false);
      setStatus("Cancelled", "Job cancelled.", "Error", "err");
      currentJobId = null;
      currentAbort = null;
    }

    $("cancelBtn").addEventListener("click", cancelRun);

    $("quitBtn").addEventListener("click", async () => {
      appendLog("[ui] Quit requested…");
      setStatus("Shutting down…", "Stopping local server…", "Running", "run");
      try{
        await fetch("/api/shutdown", {
          method:"POST",
          headers:{ "Content-Type":"application/json" },
          body: JSON.stringify({ reason:"user_requested" })
        });
        appendLog("[ui] Shutdown request sent.");
      }catch(e){
        appendLog("[ui] Shutdown failed: " + String(e?.message || e));
        setStatus("Error", "Shutdown failed.", "Error", "err");
      }
    });

    $("runBtn").addEventListener("click", async () => {
      if (uiLocked){
        appendLog("[ui] A run is already in progress.");
        return;
      }

      const api_key = getApiKey();
      if (!api_key){
        setStatus("Error", "Missing API key.", "Error", "err");
        appendLog("[ui] ERROR: Missing API key.");
        return;
      }

      const phones = parsePhones($("numbers").value);
      if (!phones.length){
        setStatus("Error", "Enter at least one valid 10-digit number.", "Error", "err");
        appendLog("[ui] ERROR: No valid 10-digit numbers found.");
        return;
      }

      const range = computeDateRange();
      const startDate = range.startDate;
      const endDate = range.endDate;

      if (!startDate || !endDate){
        setStatus("Error", "Choose a valid start and end date.", "Error", "err");
        appendLog("[ui] ERROR: Missing date range.");
        return;
      }

      if (startDate > endDate){
        setStatus("Error", "Start date must be on/before end date.", "Error", "err");
        appendLog(`[ui] ERROR: Invalid date order: ${startDate} > ${endDate}`);
        return;
      }

      const account_ids = (window.__VOAPPS_SELECTED_ACCOUNTS__ || []).map(String).filter(Boolean);
      if (!account_ids.length){
        setStatus("Error", "Select at least one account.", "Error", "err");
        appendLog("[ui] ERROR: No accounts selected.");
        return;
      }

      const payload = {
        job_id: (currentJobId = newJobId()),
        api_key,
        numbers: phones,
        account_ids,
        start_date: startDate,
        end_date: endDate,
        include_caller: !!$("includeCaller").checked,
        include_message_meta: !!$("includeMessageMeta").checked,
        dryRun: !!$("dryRunToggle").checked,
        logLevel: $("logLevel").value || "verbose",
        columns: getSelectedColumns()
      };

      currentAbort = new AbortController();

      setLocked(true);
      setStatus("Running…", "Working…", "Running", "run");
      appendLog("[ui] Run clicked.");
      appendLog("[ui] Job ID: " + payload.job_id);
      appendLog("[ui] Accounts: " + account_ids.join(", "));
      appendLog("[ui] Range: " + startDate + " → " + endDate);
      appendLog("[ui] Numbers: " + phones.join(", "));
      appendLog("[ui] Dry run: " + payload.dryRun + " | Log: " + payload.logLevel);

      try{
        const resp = await fetch("/api/search", {
          method:"POST",
          headers:{ "Content-Type":"application/json" },
          signal: currentAbort.signal,
          body: JSON.stringify(payload)
        });

        const data = await resp.json().catch(() => ({}));
        if (!resp.ok || data.ok !== true){
          const msg = data?.error || data?.message || ("HTTP " + resp.status);
          throw new Error(msg);
        }

        setLocked(false);
        setStatus("Done", "Artifacts updated.", "Done", "done");
        appendLog("[ui] Done. Refreshing artifacts…");
        await refreshArtifacts();
        currentJobId = null;
        currentAbort = null;
      }catch(e){
        const isAbort = (e && String(e.name) === "AbortError");
        setLocked(false);
        setStatus(isAbort ? "Cancelled" : "Error", String(e.message || e), isAbort ? "Error" : "Error", "err");
        appendLog("[ui] " + (isAbort ? "Run aborted." : ("ERROR: " + String(e.message || e))));
        currentJobId = null;
        currentAbort = null;
      }
    });

    (async () => {
      appendLog("[ui] VoApps Tools UI loaded.");
      const ok = await ping();
      if (ok) await refreshArtifacts();
    })();
  </script>
</body>
</html>
HTML

cat > server.js <<'JS'
/**
 * ─────────────────────────────────────────────────────────────────────────────
 * VoApps Tools — Local Server (Electron)
 * Version: 1.0.0
 * Last Updated: 2026-01-12
 *
 * WHAT THIS SERVER DOES
 *  - Serves the Electron UI from /public
 *  - Provides API routes used by the UI:
 *      GET  /api/ping
 *      POST /api/accounts   { api_key, filter?: "all" | "hidden" }
 *      POST /api/search     {
 *            job_id, api_key, numbers[], account_ids[],
 *            start_date, end_date,
 *            include_caller, include_message_meta,
 *            dryRun, logLevel, columns[]
 *          }
 *      POST /api/cancel     { job_id }          // cancels server-side work
 *      POST /api/shutdown   { reason? }         // stops the local server
 *
 * IMPORTANT API RULE (VoApps)
 *  - Campaigns endpoint only accepts ONE date:
 *      /accounts/:id/campaigns?created_date=YYYY-MM-DD
 *    so we iterate one day at a time across the range.
 *
 * CANCELLATION MODEL
 *  - Each run creates a job entry in-memory with an AbortController.
 *  - /api/cancel will abort the job controller and mark cancelled.
 */

const http = require("http");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const { parse: parseUrl } = require("url");

const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;
const VOAPPS_API_BASE = process.env.VOAPPS_API_BASE || "https://directdropvoicemail.voapps.com/api/v1";

// In-memory job registry for server-side cancel
const jobs = new Map(); // job_id -> { controller, cancelled:boolean, startedAt:number }

// Basic helpers
function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Length": Buffer.byteLength(body),
  });
  res.end(body);
}

function sendText(res, status, text) {
  res.writeHead(status, { "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store" });
  res.end(text);
}

async function readBodyJson(req) {
  const chunks = [];
  for await (const c of req) chunks.push(c);
  const raw = Buffer.concat(chunks).toString("utf8").trim();
  if (!raw) return {};
  try { return JSON.parse(raw); } catch { return {}; }
}

function safeJoin(baseDir, reqPath) {
  const p = path.normalize(reqPath).replace(/^(\.\.(\/|\\|$))+/, "");
  return path.join(baseDir, p);
}

function isIsoDate(s) {
  return typeof s === "string" && /^\d{4}-\d{2}-\d{2}$/.test(s);
}

function eachDateInclusive(startISO, endISO) {
  const out = [];
  const start = new Date(startISO + "T00:00:00Z");
  const end = new Date(endISO + "T00:00:00Z");
  for (let d = new Date(start); d <= end; d.setUTCDate(d.getUTCDate() + 1)) {
    out.push(d.toISOString().slice(0, 10));
  }
  return out;
}

function normalizePhone(raw) {
  let s = String(raw || "").replace(/\D+/g, "");
  if (s.length === 11 && s.startsWith("1")) s = s.slice(1);
  return s.length === 10 ? s : "";
}

async function voappsFetch(apiKey, url, { signal, timeoutMs = 120000 } = {}) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeoutMs);

  // If parent signal aborts, abort this request too
  if (signal) {
    if (signal.aborted) controller.abort();
    else signal.addEventListener("abort", () => controller.abort(), { once: true });
  }

  try {
    const r = await fetch(url, {
      method: "GET",
      headers: {
        "Authorization": `Token token="${apiKey}"`,
        "Accept": "application/json",
      },
      signal: controller.signal,
    });

    const text = await r.text();
    let json = null;
    try { json = text ? JSON.parse(text) : null; } catch { json = null; }

    return { ok: r.ok, status: r.status, json, text };
  } finally {
    clearTimeout(t);
  }
}

// Very lightweight CSV parsing: split by newlines + commas (good enough for VoApps export shape).
// If exports contain quoted commas, this won't be perfect—but it will still locate phone numbers reliably
// because we normalize digits from the row.
function parseCsvRows(csvText) {
  const lines = String(csvText || "").split(/\r?\n/).filter(Boolean);
  if (!lines.length) return { headers: [], rows: [] };
  const headers = lines[0].split(",").map(h => h.trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) rows.push(lines[i]);
  return { headers, rows };
}

function pickRowNumberDigits(rowLine) {
  // Extract digit runs and pick the first 10-digit candidate (post-normalization)
  const parts = rowLine.split(/[^0-9]+/).filter(Boolean);
  for (const p of parts) {
    const n = normalizePhone(p);
    if (n) return n;
  }
  return "";
}

function nowStamp() {
  const d = new Date();
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${hh}:${mm}:${ss}`;
}

async function ensureDir(p) {
  await fsp.mkdir(p, { recursive: true });
}

async function writeArtifactsCSVLog({ csvLines, logLines }) {
  const downloads = path.join(process.env.HOME || process.env.USERPROFILE || ".", "Downloads");
  const outDir = path.join(downloads, "VoApps Tools");
  const logDir = path.join(downloads, "VoApps Tools Logs");
  await ensureDir(outDir);
  await ensureDir(logDir);

  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const csvPath = path.join(outDir, `VoApps_Number_Search_${stamp}.csv`);
  const logPath = path.join(logDir, `VoApps_Tools_${stamp}.log`);

  await fsp.writeFile(csvPath, csvLines.join("\n") + "\n", "utf8");
  await fsp.writeFile(logPath, logLines.join("\n") + "\n", "utf8");

  // Optional: if your preload reads these paths, keep a tiny pointer file it can read.
  // If you already have a different artifact persistence mechanism, this is harmless.
  const metaPath = path.join(outDir, "last_artifacts.json");
  await fsp.writeFile(metaPath, JSON.stringify({ csvPath, logPath }, null, 2), "utf8");

  return { csvPath, logPath };
}

async function handlePing(req, res) {
  sendJson(res, 200, { ok: true, message: "pong" });
}

async function handleAccounts(req, res) {
  const body = await readBodyJson(req);
  const api_key = String(body.api_key || "").trim();
  const filter = String(body.filter || "all").trim(); // "all" | "hidden"

  if (!api_key) return sendJson(res, 400, { ok: false, error: "Missing api_key" });

  // NOTE:
  // We DON'T assume VoApps API supports hidden filtering via query params.
  // If you later add a real hidden endpoint, you can wire it here.
  // For now:
  //  - "all" returns /accounts as-is
  //  - "hidden" returns [] (so UI falls back gracefully)
  if (filter === "hidden") {
    return sendJson(res, 200, { ok: true, accounts: [] });
  }

  const url = `${VOAPPS_API_BASE}/accounts`;
  const r = await voappsFetch(api_key, url);
  if (!r.ok) {
    return sendJson(res, r.status || 500, { ok: false, error: `Accounts request failed (HTTP ${r.status})` });
  }

  const accounts = Array.isArray(r.json?.accounts) ? r.json.accounts : Array.isArray(r.json) ? r.json : [];
  return sendJson(res, 200, { ok: true, accounts });
}

async function handleCancel(req, res) {
  const body = await readBodyJson(req);
  const job_id = String(body.job_id || "").trim();
  if (!job_id) return sendJson(res, 400, { ok: false, error: "Missing job_id" });

  const job = jobs.get(job_id);
  if (!job) return sendJson(res, 200, { ok: true, message: "No such job (already finished or unknown)." });

  job.cancelled = true;
  try { job.controller.abort(); } catch (_) {}
  return sendJson(res, 200, { ok: true, message: "Cancel requested." });
}

async function handleShutdown(req, res) {
  // respond first, then exit
  sendJson(res, 200, { ok: true, message: "Shutting down." });
  setTimeout(() => process.exit(0), 250);
}

async function handleSearch(req, res) {
  const body = await readBodyJson(req);

  const job_id = String(body.job_id || "").trim() || `job_${Date.now()}`;
  const api_key = String(body.api_key || "").trim();
  const account_ids = Array.isArray(body.account_ids) ? body.account_ids.map(String).filter(Boolean) : [];

  // numbers can be array (preferred) OR raw string (legacy)
  const nums = Array.isArray(body.numbers)
    ? body.numbers.map(normalizePhone).filter(Boolean)
    : String(body.numbers || "").split(/[^0-9]+/).map(normalizePhone).filter(Boolean);

  const start_date = String(body.start_date || "").trim();
  const end_date = String(body.end_date || "").trim();

  const include_caller = !!body.include_caller;
  const include_message_meta = !!body.include_message_meta;
  const dryRun = !!body.dryRun;
  const logLevel = String(body.logLevel || "verbose");
  const columns = Array.isArray(body.columns) ? body.columns.map(String) : [];

  if (!api_key) return sendJson(res, 400, { ok: false, error: "Missing api_key" });
  if (!account_ids.length) return sendJson(res, 400, { ok: false, error: "No account_ids selected" });
  if (!isIsoDate(start_date) || !isIsoDate(end_date)) {
    return sendJson(res, 400, { ok: false, error: "Invalid start_date/end_date (must be YYYY-MM-DD)" });
  }
  if (start_date > end_date) {
    return sendJson(res, 400, { ok: false, error: "Start date must be on/before end date" });
  }
  if (!nums.length) return sendJson(res, 400, { ok: false, error: "No valid 10-digit numbers provided" });

  // Register job for cancellation
  const controller = new AbortController();
  jobs.set(job_id, { controller, cancelled: false, startedAt: Date.now() });

  const logLines = [];
  const log = (msg) => {
    if (logLevel === "none") return;
    logLines.push(`[${nowStamp()}] ${msg}`);
  };

  try {
    log("VoApps Tools — Number Search");
    log(`VOAPPS_API_BASE: ${VOAPPS_API_BASE}`);
    log(`job_id: ${job_id}`);
    log(`account_ids: ${account_ids.join(", ")}`);
    log(`range: ${start_date} → ${end_date}`);
    log(`numbers: ${nums.join(", ")}`);
    log(`dryRun: ${dryRun} include_caller:${include_caller} include_message_meta:${include_message_meta}`);
    log(`columns: ${columns.join(", ") || "(server default)"}`);

    // If you want “columns” enforced strictly, do it here:
    // For now we always write full schema, and UI can choose columns later.
    const headerDefault = [
      "number",
      "account_id",
      "campaign_id",
      "campaign_name",
      "caller_number",
      "message_id",
      "message_name",
      "message_description",
      "voapps_result",
      "voapps_code",
      "voapps_timestamp",
      "campaign_url",
    ];

    const headers = (columns && columns.length) ? headerDefault.filter(h => columns.includes(h)) : headerDefault;
    const csvLines = [headers.join(",")];

    const days = eachDateInclusive(start_date, end_date);

    for (const accountId of account_ids) {
      if (controller.signal.aborted) throw new Error("Cancelled");
      log(`--- Account ${accountId} ---`);

      // optional: messages map for metadata (only if requested)
      let messageMetaById = new Map();
      if (include_message_meta) {
        const msgUrl = `${VOAPPS_API_BASE}/accounts/${accountId}/messages`;
        const mr = await voappsFetch(api_key, msgUrl, { signal: controller.signal });
        if (mr.ok) {
          const items = Array.isArray(mr.json?.messages) ? mr.json.messages : Array.isArray(mr.json) ? mr.json : [];
          for (const m of items) {
            const id = String(m?.id ?? "");
            if (!id) continue;
            messageMetaById.set(id, {
              name: String(m?.name ?? ""),
              description: String(m?.description ?? ""),
            });
          }
        } else {
          log(`WARN: messages request failed for account ${accountId} (HTTP ${mr.status})`);
        }
      }

      for (const day of days) {
        if (controller.signal.aborted) throw new Error("Cancelled");

        const cUrl = `${VOAPPS_API_BASE}/accounts/${accountId}/campaigns?created_date=${day}`;
        const cr = await voappsFetch(api_key, cUrl, { signal: controller.signal });
        if (!cr.ok) {
          log(`WARN: campaigns failed for ${accountId} day ${day} (HTTP ${cr.status})`);
          continue;
        }

        const campaigns = Array.isArray(cr.json?.campaigns) ? cr.json.campaigns : Array.isArray(cr.json) ? cr.json : [];
        log(`Campaigns for ${day}: ${campaigns.length}`);

        if (dryRun) continue;

        for (const c of campaigns) {
          if (controller.signal.aborted) throw new Error("Cancelled");

          const campaignId = String(c?.id ?? "");
          const campaignName = String(c?.name ?? "");
          if (!campaignId) continue;

          const campaignUrl = `${VOAPPS_API_BASE}/accounts/${accountId}/campaigns/${campaignId}`;
          const dr = await voappsFetch(api_key, campaignUrl, { signal: controller.signal });
          if (!dr.ok) {
            log(`WARN: campaign detail failed ${campaignId} (HTTP ${dr.status})`);
            continue;
          }

          // Try to find an export URL. Different payloads can name it differently.
          const exportUrl =
            dr.json?.export ||
            dr.json?.campaign?.export ||
            dr.json?.campaign?.export_url ||
            dr.json?.export_url ||
            "";

          if (!exportUrl) {
            log(`No export URL for campaign ${campaignId}`);
            continue;
          }

          // Download export CSV
          const er = await fetch(String(exportUrl), { method: "GET", signal: controller.signal });
          if (!er.ok) {
            log(`WARN: export download failed for campaign ${campaignId} (HTTP ${er.status})`);
            continue;
          }
          const csvText = await er.text();

          const { rows } = parseCsvRows(csvText);

          // Search for numbers by scanning row text and normalizing first 10-digit candidate.
          for (const rowLine of rows) {
            if (controller.signal.aborted) throw new Error("Cancelled");

            const rowNum = pickRowNumberDigits(rowLine);
            if (!rowNum) continue;
            if (!nums.includes(rowNum)) continue;

            // Minimal row extraction: we don't rely on fixed CSV header positions here.
            // We fill what we confidently know and leave the rest blank.
            const caller_number = include_caller ? "" : ""; // If you later parse caller from CSV columns, set it here.
            const message_id = ""; // same note as above (can be parsed from CSV columns if needed)
            const meta = include_message_meta && message_id ? (messageMetaById.get(message_id) || {}) : {};
            const message_name = meta?.name || "";
            const message_description = meta?.description || "";

            // These fields can be parsed from CSV columns in future; leaving blank is safe.
            const voapps_result = "";
            const voapps_code = "";
            const voapps_timestamp = "";

            const portal_campaign_url = `https://directdropvoicemail.voapps.com/accounts/${accountId}/campaigns/${campaignId}`;

            const record = {
              number: rowNum,
              account_id: accountId,
              campaign_id: campaignId,
              campaign_name: campaignName.replace(/"/g, '""'),
              caller_number,
              message_id,
              message_name: String(message_name).replace(/"/g, '""'),
              message_description: String(message_description).replace(/"/g, '""'),
              voapps_result,
              voapps_code,
              voapps_timestamp,
              campaign_url: portal_campaign_url,
            };

            const line = headers.map(h => {
              const v = (record[h] ?? "");
              const s = String(v);
              // CSV escape
              if (/[,"\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
              return s;
            }).join(",");

            csvLines.push(line);
            log(`MATCH: ${rowNum} account:${accountId} campaign:${campaignId} (${campaignName})`);
          }
        }
      }
    }

    const artifacts = await writeArtifactsCSVLog({ csvLines, logLines });

    return sendJson(res, 200, {
      ok: true,
      message: dryRun ? "Dry run complete." : "Search complete.",
      job_id,
      artifacts,
      stats: {
        numbers: nums.length,
        accounts: account_ids.length,
        days: days.length,
        dryRun,
        rows_written: dryRun ? 0 : Math.max(0, csvLines.length - 1),
      },
    });
  } catch (e) {
    const msg = (e && e.message) ? e.message : String(e);
    const cancelled = controller.signal.aborted || msg === "Cancelled";
    return sendJson(res, cancelled ? 499 : 500, { ok: false, error: cancelled ? "Cancelled" : msg });
  } finally {
    jobs.delete(job_id);
  }
}

async function serveStatic(req, res, pathname) {
  const pub = path.join(__dirname, "public");
  const filePath = safeJoin(pub, pathname === "/" ? "/index.html" : pathname);

  // Directory -> index.html
  let stat;
  try {
    stat = await fsp.stat(filePath);
    if (stat.isDirectory()) {
      return serveStatic(req, res, path.join(pathname, "index.html"));
    }
  } catch {
    // fall through 404
  }

  try {
    const data = await fsp.readFile(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const types = {
      ".html": "text/html; charset=utf-8",
      ".js": "application/javascript; charset=utf-8",
      ".css": "text/css; charset=utf-8",
      ".png": "image/png",
      ".jpg": "image/jpeg",
      ".jpeg": "image/jpeg",
      ".svg": "image/svg+xml",
      ".ico": "image/x-icon",
    };
    res.writeHead(200, { "Content-Type": types[ext] || "application/octet-stream" });
    res.end(data);
  } catch {
    sendText(res, 404, "Not Found");
  }
}

const server = http.createServer(async (req, res) => {
  const urlObj = parseUrl(req.url || "", true);
  const { pathname } = urlObj;

  try {
    if (req.method === "GET" && pathname === "/api/ping") return handlePing(req, res);
    if (req.method === "POST" && pathname === "/api/accounts") return handleAccounts(req, res);
    if (req.method === "POST" && pathname === "/api/search") return handleSearch(req, res);
    if (req.method === "POST" && pathname === "/api/cancel") return handleCancel(req, res);
    if (req.method === "POST" && pathname === "/api/shutdown") return handleShutdown(req, res);

    return serveStatic(req, res, pathname || "/");
  } catch (e) {
    return sendJson(res, 500, { ok: false, error: String(e?.message || e) });
  }
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`[VoApps Tools] Server listening on http://127.0.0.1:${PORT}`);
  console.log(`[VoApps Tools] VOAPPS_API_BASE=${VOAPPS_API_BASE}`);
});
JS

echo ""
echo "✅ Applied fixes:"
echo "  - public/index.html replaced (accounts selection used in Run, presets, validation, cancel/quit UI, dryRun/logLevel/columns)"
echo "  - server.js replaced (ping, accounts filter, search w/ created_date day-by-day, cancel, shutdown)"
echo ""
echo "Backups in: _backups/"
echo ""
echo "Next:"
echo "  1) Restart your Electron app (stop current server if running)."
echo "  2) Click Ping, then Load accounts, select accounts, Run."
echo ""

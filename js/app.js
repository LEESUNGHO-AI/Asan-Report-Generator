/* ══════════════════════════════════════════════════════════════════════════
 *  app.js — 런타임 (데이터 연계 · 설정 · 문서 생성)
 * ══════════════════════════════════════════════════════════════════════════ */
(function () {
  "use strict";
  const CFG = window.GovConfig, CORE = window.GovCore;
  const $ = (id) => document.getElementById(id);
  const S = { raw: {}, proc: null, busy: false, G: null };
  const LSK = "asan-govdoc-docset-v5";
  const LSH = "asan-govdoc-history-v5";

  /* ── 로그 ── */
  function log(t, m) {
    const e = $("log"); e.classList.add("show");
    const d = document.createElement("div");
    d.className = { ok: "ok", warn: "warn", err: "err", info: "info" }[t] || "";
    d.textContent = `${new Date().toLocaleTimeString("ko-KR")}  ${m}`;
    e.appendChild(d); e.scrollTop = e.scrollHeight;
  }
  function srcState(id, st, txt) {
    const el = $("src-" + id); if (!el) return;
    el.className = "src " + st; el.querySelector(".st").textContent = txt;
  }

  /* ── 설정(DOCSET) ── */
  function loadSettings() {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(LSK) || "{}"); } catch (e) { saved = {}; }
    return Object.assign(JSON.parse(JSON.stringify(CFG.DOCSET)), saved);
  }
  let DS = loadSettings();

  function autoDocNo(d) {
    const y = (d || new Date()).getFullYear();
    return `${DS.docNoPrefix}-${y}-${String(DS.docNoSeq || 1).padStart(4, "0")}`;
  }
  const sign = (s) => `${s.title || ""} ${s.name || ""}`.trim();
  const parseSign = (v) => {
    const p = String(v || "").trim().split(/\s+/);
    return p.length > 1 ? { title: p[0], name: p.slice(1).join(" ") } : { title: p[0] || "", name: "" };
  };

  function fillForm() {
    $("fDocNo").value = DS.docNo || autoDocNo();
    $("fDate").value = (DS.issueDate || new Date().toISOString().slice(0, 10));
    $("fRecipient").value = DS.recipient;
    $("fRecipientDept").value = DS.recipientDept || CFG.PROJECT.ownerDept;
    $("fVia").value = DS.via || "";
    $("fSenderName").value = DS.senderName;
    $("fDrafter").value = sign(DS.drafter);
    $("fReviewer").value = sign(DS.reviewer);
    $("fApprover").value = sign(DS.approver);
    $("fDept").value = DS.handlerDept;
    $("fTel").value = `${DS.tel} / ${DS.fax}`;
    $("fOpen").value = DS.openLevel;
    $("fPeriodTo").value = (DS.periodTo || CFG.PROJECT.periodTo);
  }
  function readForm() {
    DS.docNo = $("fDocNo").value.trim() || autoDocNo();
    DS.issueDate = $("fDate").value;
    DS.recipient = $("fRecipient").value.trim();
    DS.recipientDept = $("fRecipientDept").value.trim();
    DS.via = $("fVia").value.trim();
    DS.senderName = $("fSenderName").value.trim();
    DS.drafter = parseSign($("fDrafter").value);
    DS.reviewer = parseSign($("fReviewer").value);
    DS.approver = parseSign($("fApprover").value);
    DS.handlerDept = $("fDept").value.trim();
    const t = $("fTel").value.split("/");
    DS.tel = (t[0] || "").trim(); DS.fax = (t[1] || "").trim();
    DS.openLevel = $("fOpen").value;
    DS.periodTo = $("fPeriodTo").value;
    CFG.PROJECT.periodTo = DS.periodTo || CFG.PROJECT.periodTo;
    $("hPeriod").textContent = G().dateKr(CFG.PROJECT.periodTo);
  }
  function saveSettings() {
    readForm();
    localStorage.setItem(LSK, JSON.stringify(DS));
    log("ok", "문서정보 설정 저장 완료");
    updateKPI();
  }

  const G = () => (S.G || (S.G = window.Gongmun.create(window.docx)));

  /* ── 데이터 로딩 ── */
  async function fetchJSON(url) {
    const r = await fetch(url, { cache: "no-store" });
    if (!r.ok) throw new Error("HTTP " + r.status);
    return r.json();
  }
  async function fetchText(url) {
    const r = await fetch(url, { cache: "no-store" });
    if (!r.ok) throw new Error("HTTP " + r.status);
    return r.text();
  }

  async function loadAll() {
    log("info", "데이터 원천 연계 시작");
    const SRC = CFG.SRC;

    // BMS
    try { srcState("bms", "wait", "연계중"); S.raw.bms = await fetchJSON(SRC.bms); srcState("bms", "ok", "정상"); log("ok", `BMS 연계 완료 (${(S.raw.bms.items || []).length}개 항목)`); }
    catch (e) { srcState("bms", "err", "실패"); log("err", "BMS 연계 실패 : " + e.message); }

    // WBS
    try {
      srcState("wbs", "wait", "연계중");
      const [a, b] = await Promise.all([fetchJSON(SRC.wbsSum), fetchJSON(SRC.wbsDat)]);
      S.raw.wbsSum = a; S.raw.wbsDat = b; srcState("wbs", "ok", "정상");
      log("ok", `WBS 연계 완료 (${(b.items || []).length}건)`);
    } catch (e) { srcState("wbs", "err", "실패"); log("err", "WBS 연계 실패 : " + e.message); }

    // 자산
    try { srcState("asset", "wait", "연계중"); S.raw.asset = await fetchJSON(SRC.asset); srcState("asset", "ok", "정상"); log("ok", `자산관리 연계 완료 (${((S.raw.asset.summary || {}).total_assets) || 0}점)`); }
    catch (e) { srcState("asset", "err", "실패"); log("err", "자산관리 연계 실패 : " + e.message); }

    // 인력 — JSON 우선, 실패 시 HTML 폴백
    try {
      srcState("hr", "wait", "연계중");
      S.raw.hr = await fetchJSON(SRC.hr);
      srcState("hr", "ok", "정상(JSON)");
      log("ok", `인력관리 연계 완료 — hr.json 스키마 ${(S.raw.hr.meta || {}).schema || "v1"}`);
    } catch (e) {
      log("warn", "hr.json 미연계 — HTML 파싱으로 대체 : " + e.message);
      try {
        S.raw.hr = await fetchText(SRC.hrFallback);
        srcState("hr", "warn", "대체(HTML)");
        log("warn", "인력관리 HTML 폴백 연계 — data/hr.json 배포 권장");
      } catch (e2) { srcState("hr", "err", "실패"); log("err", "인력관리 연계 실패 : " + e2.message); }
    }

    process();
  }

  function process() {
    try {
      S.proc = {
        bms: CORE.processBMS(S.raw.bms || {}),
        wbs: CORE.processWBS(S.raw.wbsSum || {}, S.raw.wbsDat || {}),
        asset: CORE.processAsset(S.raw.asset || {}),
        hr: CORE.processHR(S.raw.hr || {}),
      };
      updateKPI();
      $("sbLive").textContent = `데이터 연계 완료 · 기준 ${S.proc.bms.updatedAt || new Date().toLocaleDateString("ko-KR")}`;
      log("ok", "데이터 정규화 완료 — 문서 생성 준비됨");
      if (S.proc.hr.roster && S.proc.hr.roster.some((p) => p.ratio > 100)) {
        const n = S.proc.hr.roster.filter((p) => p.ratio > 100).length;
        log("warn", `참여율 100% 초과 인력 ${n}명 검출 — 인건비 중복계상 방지를 위한 정정 필요`);
      }
    } catch (e) { log("err", "데이터 처리 오류 : " + e.message); }
  }

  function updateKPI() {
    if (!S.proc) return;
    const { bms, wbs, asset, hr } = S.proc;
    const g = G();
    const ctx = CORE.buildCtx("monthly", CFG.PROJECT, new Date());
    $("k1").textContent = bms.execRate.toFixed(2) + "%";
    $("k1s").textContent = `${g.eokStr(bms.totalExec)} / ${g.eokStr(bms.totalBudget)}`;
    $("k2").textContent = wbs.overall.actualRate.toFixed(1) + "%";
    $("k2s").textContent = `계획 ${wbs.overall.plannedRate.toFixed(1)}% · 지연 ${wbs.overall.delayed}건`;
    $("k3").textContent = hr.total + "명";
    $("k3s").textContent = `실인원 ${hr.headcount}명 · 가동 ${hr.active}명`;
    $("k4").textContent = asset.total + "점";
    $("k4s").textContent = `${g.eokStr(asset.value)} · 등재 ${asset.issuedRate.toFixed(1)}%`;
    $("k5").textContent = "D-" + ctx.dday;
    $("k5s").textContent = `기간소진 ${ctx.timePct.toFixed(1)}%`;
    $("hToday").textContent = ctx.dateKr;
    $("hPeriod").textContent = g.dateKr(CFG.PROJECT.periodTo);
  }

  /* ── 문서 카드 ── */
  function renderDocs() {
    const order = ["official", "monthly", "quarterly", "annual", "weekly", "brief"];
    $("docs").innerHTML = order.map((k) => {
      const t = CFG.DOCTYPES[k];
      return `<div class="doc" data-k="${k}">
        <span class="kind ${k === "official" ? "o" : ""}">${t.kind}</span>
        <h3>${t.title}</h3>
        <p>${t.desc}</p>
        <div class="meta"><span>제출처 <b>${t.target}</b></span><span>${t.cycle}</span></div>
        <div class="meta"><span>${t.badge}</span><span>항목기호 ${t.style}</span></div>
      </div>`;
    }).join("");
    [...document.querySelectorAll(".doc")].forEach((el) => el.addEventListener("click", () => gen(el)));
  }

  /* ── 생성 ── */
  async function gen(el) {
    if (S.busy) return;
    if (!S.proc) { log("err", "데이터 연계가 완료되지 않았습니다."); return; }
    const type = el.dataset.k, T = CFG.DOCTYPES[type];
    S.busy = true; el.classList.add("busy");
    try {
      readForm();
      const g = G();
      const now = DS.issueDate ? new Date(DS.issueDate + "T09:00:00") : new Date();
      const ctx = CORE.buildCtx(type, CFG.PROJECT, now);
      const an = CORE.analyze(S.proc.bms, S.proc.wbs, S.proc.asset, S.proc.hr, ctx);
      const P = { bms: S.proc.bms, wbs: S.proc.wbs, asset: S.proc.asset, hr: S.proc.hr, an, ctx, CFG, DS, G: g };
      log("info", `${T.title} 생성 시작`);
      let doc;
      if (type === "official") doc = window.DocOfficial.build(window.docx, P);
      else if (type === "brief") doc = window.DocBrief.build(window.docx, P);
      else doc = window.DocReport.build(window.docx, type, P);
      const blob = await window.docx.Packer.toBlob(doc);
      const fn = `${T.file}_${ctx.year}${String(ctx.month).padStart(2, "0")}${String(ctx.day).padStart(2, "0")}.docx`;
      saveAs(blob, fn);
      addHist({ no: DS.docNo, title: T.title, file: fn, at: ctx.dateKr, open: DS.openLevel });
      DS.docNoSeq = (DS.docNoSeq || 1) + 1;
      DS.docNo = autoDocNo(now);
      $("fDocNo").value = DS.docNo;
      localStorage.setItem(LSK, JSON.stringify(DS));
      log("ok", `${T.title} 생성 완료 → ${fn}`);
    } catch (e) {
      log("err", `생성 실패 : ${e.message}`);
      console.error(e);
    } finally { S.busy = false; el.classList.remove("busy"); }
  }

  /* ── 문서등록대장 ── */
  function getHist() { try { return JSON.parse(localStorage.getItem(LSH) || "[]"); } catch (e) { return []; } }
  function addHist(r) {
    const h = getHist(); h.unshift(Object.assign({ ts: Date.now() }, r));
    localStorage.setItem(LSH, JSON.stringify(h.slice(0, 60))); renderHist();
  }
  function renderHist() {
    const h = getHist();
    if (!h.length) { $("hist").innerHTML = '<div class="empty">생성된 문서가 없습니다.</div>'; return; }
    $("hist").innerHTML = `<table class="hist"><thead><tr>
      <th style="width:60px">연번</th><th style="width:200px">문서번호</th><th>문서명</th>
      <th style="width:110px">생산일자</th><th style="width:150px">공개구분</th></tr></thead><tbody>
      ${h.map((r, i) => `<tr><td class="c">${h.length - i}</td><td class="c">${r.no || "-"}</td>
        <td>${r.title}</td><td class="c">${r.at}</td><td class="c">${r.open || "-"}</td></tr>`).join("")}
      </tbody></table>`;
  }

  /* ── 초기화 ── */
  function init() {
    fillForm(); renderDocs(); renderHist();
    $("hToday").textContent = G().dateKr(new Date());
    $("pToggle").addEventListener("click", () => $("pBody").classList.toggle("hide"));
    $("btnSave").addEventListener("click", saveSettings);
    $("btnReset").addEventListener("click", () => {
      localStorage.removeItem(LSK); DS = loadSettings(); fillForm();
      log("warn", "문서정보 설정을 기본값으로 복원했습니다.");
    });
    loadAll();
  }
  document.addEventListener("DOMContentLoaded", init);
})();

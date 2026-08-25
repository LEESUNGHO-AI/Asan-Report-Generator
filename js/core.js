/* ══════════════════════════════════════════════════════════════════════════
 *  core.js — 데이터 정규화 · 분석 엔진
 *  ※ v5.0 변경: 모든 금액을 '원' 단위 원본으로 보존(정산·감사 대응).
 *              억원 환산값은 표시용(_억)으로 별도 보관.
 * ══════════════════════════════════════════════════════════════════════════ */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(root.GovConfig || require("./config.js"));
  else root.GovCore = factory(root.GovConfig);
})(typeof self !== "undefined" ? self : this, function (CFG) {

  const { BMS_UNIT_MAP, UNIT_NAMES, BIMOK_CLEAN, BIMOK_ORDER } = CFG;

  const eok = (v) => (+v || 0) / 1e8;
  const f1 = (v) => Number(v || 0).toFixed(1);
  const f2 = (v) => Number(v || 0).toFixed(2);
  const rate = (e, b) => (b ? (e / b) * 100 : 0);

  /* ── 예산(BMS) ───────────────────────────────────────────────────────── */
  function processBMS(bms) {
    const s = bms.summary || {};
    const totalBudget = +s["총사업비"] || 0;
    const totalExec = +s["총집행액"] || 0;
    const totalRemain = +(s["총잔액"] != null ? s["총잔액"] : s["잔액"]) || 0;
    const execRate = +s["전체집행률"] || 0;

    // 재원별 — ※ 재원별 실집행 데이터 미연계. 교부비율 안분 '추정치'이며 문서에 각주 표기됨
    const sources = (bms.source_summary || []).map((r) => ({
      name: r["재원"], rate: +r["비율"] || 0, budget: +r["금액"] || 0,
      exec: Math.round((+r["금액"] || 0) * (execRate / 100)),
      remain: Math.round((+r["금액"] || 0) * (1 - execRate / 100)),
      estimated: true,
    }));

    // 비목별
    const mg = {};
    for (const b of bms.bimok_summary || []) {
      const nm = BIMOK_CLEAN[b["비목"]] || b["비목"];
      if (!mg[nm]) mg[nm] = { b: 0, e: 0, cnt: 0, raw: b["비목"] };
      mg[nm].b += +b["예산"] || 0; mg[nm].e += +b["집행"] || 0; mg[nm].cnt += +b["항목수"] || 0;
    }
    const cats = [];
    for (const nm of BIMOK_ORDER) if (mg[nm]) {
      const m = mg[nm];
      cats.push({ name: nm, code: m.raw, budget: m.b, exec: m.e, remain: m.b - m.e, rate: rate(m.e, m.b), cnt: m.cnt });
    }

    // 단위사업별 + 연도별
    const units = {}; let commonB = 0, commonE = 0;
    const years = {};
    for (const it of bms.items || []) {
      const ex = it["집행액"] != null ? +it["집행액"] : (+it["사용금액합계"] || +it["사용금액"] || 0);
      const bd = +it["총예산"] || 0;
      const num = BMS_UNIT_MAP[it["항목명"]];
      if (num) { if (!units[num]) units[num] = { b: 0, e: 0 }; units[num].b += bd; units[num].e += ex; }
      else { commonB += bd; commonE += ex; }
      for (const k of Object.keys(it)) {
        const m = String(k).match(/^(\d{4})년(예산|집행)$/);
        if (!m) continue;
        const y = m[1]; if (!years[y]) years[y] = { budget: 0, exec: 0 };
        years[y][m[2] === "예산" ? "budget" : "exec"] += +it[k] || 0;
      }
    }
    const projects = [];
    for (const num of Object.keys(UNIT_NAMES).map(Number).sort((a, b) => a - b)) {
      const u = units[num]; if (!u || (u.b === 0 && u.e === 0)) continue;
      projects.push({ num, name: UNIT_NAMES[num], budget: u.b, exec: u.e, remain: u.b - u.e, rate: rate(u.e, u.b) });
    }
    const byYear = Object.keys(years).sort().map((y) => ({
      year: y, budget: years[y].budget, exec: years[y].exec,
      remain: years[y].budget - years[y].exec, rate: rate(years[y].exec, years[y].budget),
    }));

    return {
      totalBudget, totalExec, totalRemain, execRate,
      statusCnt: s["상태별"] || {}, sources, cats, projects, byYear,
      commonB, commonE, itemCount: s["항목수"] || (bms.items || []).length,
      updatedAt: bms.updated_at || "",
    };
  }

  /* ── 공정(WBS) ───────────────────────────────────────────────────────── */
  function processWBS(sum, data) {
    const t = sum.total || {};
    const cats0 = sum.byCategory || [];
    const sumCat = (k) => cats0.reduce((a, c) => a + (+c[k] || 0), 0);
    const hasCnt = (+t.total || 0) > 0;
    const overall = {
      total: hasCnt ? +t.total : sumCat("total"),
      done: hasCnt ? +t.done : sumCat("done"),
      inProg: hasCnt ? +t.inProg : sumCat("inProg"),
      delayed: hasCnt ? +t.delayed : sumCat("delayed"),
      waiting: hasCnt ? +t.waiting : sumCat("waiting"),
      plannedRate: +t.plannedRate || 0, actualRate: +t.actualRate || 0,
      achieveRate: +t.achieveRate || 0, deviation: +t.deviation || 0,
      updatedAt: t.updatedAt || (sum.meta || {}).updatedAt || "",
    };
    const byCat = cats0.map((c) => ({
      name: String(c.name || "").replace(/[^\u3131-\uD79D\w\s()·&/.-]/g, "").trim(),
      total: +c.total || 0, done: +c.done || 0, inProg: +c.inProg || 0,
      delayed: +c.delayed || 0, waiting: +c.waiting || 0,
      plannedRate: +c.plannedRate || 0, actualRate: +c.actualRate || 0,
      deviation: +c.deviation || 0, achieveRate: +c.achieveRate || 0, note: c.note || "",
    }));
    const isDateish = (x) => /GMT|\d{4}-\d{2}|Mon |Tue |Wed |Thu |Fri |Sat |Sun /.test(String(x || ""));
    const lvl1 = [];
    for (const r of data.items || []) {
      if (String(r.level) !== "1") continue;
      const nm = String(r.name || "").replace(/^\[[^\]]+\]\s*/, "");
      if (!nm || nm === "범례" || r.wbsId === "범례") continue;
      lvl1.push({
        wbsId: r.wbsId, name: nm, category: r.category,
        org: isDateish(r.organization) ? "제일엔지니어링" : (r.organization || "제일엔지니어링"),
        weight: r.weight, planned: +r.plannedRate || 0, actual: +r.actualRate || 0,
        deviation: +r.deviation || 0, status: r.status || "-",
      });
    }
    const delayed = (data.items || []).filter((r) => r.status === "지연" && String(r.level) !== "1").map((r) => ({
      wbsId: r.wbsId, name: String(r.name || "").replace(/^\[[^\]]+\]\s*/, ""),
      org: isDateish(r.organization) ? "제일엔지니어링" : (r.organization || "-"),
      planned: +r.plannedRate || 0, actual: +r.actualRate || 0, deviation: +r.deviation || 0,
      endDate: r.endDate || "-", note: r.note || "",
    }));
    const units = [];
    for (const r of data.items || []) {
      if (String(r.level) !== "2") continue;
      const m = String(r.name || "").match(/^\[4\.(\d+)\]\s*(.+)$/);
      if (!m) continue;
      const planned = +r.plannedRate || 0, actual = +r.actualRate || 0;
      units.push({ idx: +m[1], name: m[2].trim(), planned, actual, deviation: +(actual - planned).toFixed(1), status: r.status || "-" });
    }
    units.sort((a, b) => a.idx - b.idx);
    return { overall, byCat, lvl1, units, delayed };
  }

  /* ── 자산(중요재산) ──────────────────────────────────────────────────── */
  function processAsset(a) {
    const s = a.summary || {};
    const byCat = Object.keys(s.by_category || {}).map((k) => ({
      name: k, count: s.by_category[k], value: +(s.by_category_value || {})[k] || 0,
    })).sort((x, y) => y.value - x.value);
    const byMgr = Object.keys(s.by_manager || {}).map((k) => ({ name: k, count: s.by_manager[k] })).sort((x, y) => y.count - x.count);
    const byLoc = Object.keys(s.by_location || {}).map((k) => ({ name: k, count: s.by_location[k] })).sort((x, y) => y.count - x.count);
    // 「보조금법」 제35조 중요재산: 취득가액 5천만원 이상 건별 관리대상
    const val = (x) => +x.구매금액 || +x.취득가액 || +x.acq_value || +x.value || 0;
    const dt = (x) => { const d = x.구매일자 || x.취득일 || x.acq_date; return (d && d.start) ? String(d.start).slice(0, 10) : (typeof d === "string" ? d.slice(0, 10) : "-"); };
    const major = (a.assets || []).filter((x) => val(x) >= 50000000)
      .map((x) => ({
        name: x.자산명 || x.name || "-", cat: x.자산분류 || x.category || "-",
        code: x.표준자산코드 || "-", value: val(x),
        loc: x.설치위치 || x.location || "미배치", date: dt(x),
      })).sort((p, q) => q.value - p.value).slice(0, 20);
    return {
      total: +s.total_assets || 0, value: +s.total_value || 0,
      inUse: (s.by_status || {})["사용중"] || 0, standby: (s.by_status || {})["대기중"] || 0,
      issuedRate: s.issued_rate != null ? +s.issued_rate : 0,
      byCat, byMgr, byLoc, major, syncedAt: (a.meta || {}).synced_at || "",
    };
  }

  /* ── 참여인력 ──────────────────────────────────────────────────────── */
  /**
   * 우선순위 1) hr.json (schema hr.v1) — 명부·참여율 포함
   *          2) 레거시 JSON(orgs 배열만)
   *          3) HTML 파싱 폴백 — 요약표 + org-card 명부까지 복원
   */
  function processHR(input) {
    let d = null;
    if (input && typeof input === "object" && Array.isArray(input.orgs)) d = input;
    else if (typeof input === "string" && input.trim().charAt(0) === "{") {
      try { const j = JSON.parse(input); if (Array.isArray(j.orgs)) d = j; } catch (e) { d = null; }
    }
    if (!d) d = parseHRHtml(String(input || ""));
    return normalizeHR(d);
  }

  function normalizeHR(d) {
    const orgs = (d.orgs || []).map((o) => {
      const members = (o.members || []).map((m) => ({
        name: m.name || "-", position: m.position || "-", role: m.role || "-",
        ratio: +m.ratio || 0, from: m.from || "", to: m.to || "",
        status: m.status || (m.active ? "활성" : "종료"),
      }));
      const uniq = new Map();
      members.forEach((m) => {
        const k = `${m.name}|${m.position}`;
        const p = uniq.get(k);
        if (!p) uniq.set(k, { name: m.name, position: m.position, ratio: m.ratio, status: m.status, from: m.from, to: m.to, roles: [m.role] });
        else { p.ratio += m.ratio; p.roles.push(m.role); if (m.status === "활성") p.status = "활성"; }
      });
      const persons = [...uniq.values()];
      const total = +o.total || members.length;
      const active = +o.active || persons.filter((p) => p.status === "활성").length;
      return {
        org: o.org, role: o.role || "간접보조사업자",
        total, active, ended: +o.ended || Math.max(0, total - active),
        rate: o.rate != null ? +o.rate : (total ? +((active / total) * 100).toFixed(1) : 0),
        headcount: o.headcount || persons.length || total,
        manRatio: o.manRatio != null ? +o.manRatio : +persons.reduce((a, p) => a + p.ratio, 0).toFixed(1),
        activeManRatio: o.activeManRatio != null ? +o.activeManRatio : +persons.filter((p) => p.status === "활성").reduce((a, p) => a + p.ratio, 0).toFixed(1),
        members, persons,
      };
    });
    const total = d.total != null ? +d.total : orgs.reduce((a, o) => a + o.total, 0);
    const active = d.active != null ? +d.active : orgs.reduce((a, o) => a + o.active, 0);
    const ended = d.ended != null ? +d.ended : orgs.reduce((a, o) => a + o.ended, 0);
    const headcount = orgs.reduce((a, o) => a + (o.headcount || 0), 0);
    const roster = [];
    orgs.forEach((o) => o.persons.forEach((p) => roster.push(Object.assign({ org: o.org }, p))));
    roster.sort((a, b) => (b.status === "활성") - (a.status === "활성") || b.ratio - a.ratio);
    return {
      total, active, ended, headcount,
      rate: total ? +((active / total) * 100).toFixed(1) : 0,
      manRatio: +orgs.reduce((a, o) => a + o.manRatio, 0).toFixed(1),
      activeManRatio: +orgs.reduce((a, o) => a + o.activeManRatio, 0).toFixed(1),
      orgs, roster,
      updatedAt: (d.meta || {}).updated_at || d.updatedAt || "",
      hasRoster: roster.length > 0,
    };
  }

  /** HTML 폴백 — 요약표 + 기관별 org-card 명부 복원 */
  function parseHRHtml(html) {
    const strip = (x) => String(x || "").replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&").replace(/[\u2705\u274C\u23F8\uFE0F]/g, "").replace(/\s+/g, " ").trim();
    const rowsOf = (t) => {
      const body = (t.match(/<tbody[^>]*>([\s\S]*?)<\/tbody>/i) || [null, t])[1];
      const out = []; const rre = /<tr[^>]*>([\s\S]*?)<\/tr>/gi; let m;
      while ((m = rre.exec(body))) {
        const cells = []; const cre = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi; let c;
        while ((c = cre.exec(m[1]))) cells.push(strip(c[1]));
        if (cells.length) out.push(cells);
      }
      return out;
    };
    const orgs = [];
    const tables = html.match(/<table[\s\S]*?<\/table>/gi) || [];
    for (const t of tables) {
      const head = strip((t.match(/<thead[\s\S]*?<\/thead>/i) || [""])[0]);
      if (!/기관명/.test(head) || !/총원/.test(head)) continue;
      for (const r of rowsOf(t)) {
        if (!r[0] || r[0] === "합계") continue;
        const n = (i) => parseInt(String(r[i] || "").replace(/[^\d]/g, ""), 10) || 0;
        orgs.push({ org: r[0], total: n(1), active: n(2), ended: n(3),
          rate: parseFloat(String(r[4] || "").replace(/[^\d.]/g, "")) || 0,
          role: r[5] || "", members: [] });
      }
      break;
    }
    const cards = html.split(/<div class="org-card"[^>]*>/i).slice(1);
    for (const card of cards) {
      const title = strip((card.match(/<div class="org-card-title"[^>]*>([\s\S]*?)<\/div>/i) || [])[1] || "");
      if (!title) continue;
      let org = orgs.find((o) => o.org === title) || orgs.find((o) => title.replace(/\s/g, "").indexOf(o.org.replace(/\s/g, "")) >= 0);
      if (!org) { org = { org: title, total: 0, active: 0, ended: 0, rate: 0, role: "", members: [] }; orgs.push(org); }
      const t = (card.match(/<table[\s\S]*?<\/table>/i) || [""])[0];
      for (const r of rowsOf(t)) {
        if (r.length < 5) continue;
        const pm = String(r[4] || "").match(/(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})/);
        org.members.push({ name: r[0], position: r[1], role: r[2],
          ratio: parseFloat(String(r[3] || "").replace(/[^\d.]/g, "")) || 0,
          from: pm ? pm[1] : "", to: pm ? pm[2] : "",
          status: /활성/.test(r[5] || "") ? "활성" : "종료" });
      }
    }
    if (!orgs.length) {
      const kpi = (html.match(/class="sv">(\d+)<\/div><div class="ss">[^<]*기관/) || [])[1];
      return { total: +kpi || 0, active: 0, ended: 0, orgs: [] };
    }
    return { orgs: orgs };
  }

  /* ── 보고 시점 컨텍스트 ──────────────────────────────────────────────── */
  function buildCtx(type, project, now) {
    const d = now || new Date();
    const st = new Date(project.periodFrom), en = new Date(project.periodTo);
    const dday = Math.max(0, Math.ceil((en - d) / 864e5));
    const timePct = Math.min(100, Math.max(0, ((d - st) / (en - st)) * 100));
    const monthsLeft = Math.max(0, Math.round(dday / 30.4));
    const q = Math.floor(d.getMonth() / 3) + 1;
    const period = {
      weekly: `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate() - 6 > 0 ? d.getDate() - 6 : 1}. ~ ${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}.`,
      monthly: `${d.getFullYear()}. ${d.getMonth() + 1}. 1. ~ ${d.getFullYear()}. ${d.getMonth() + 1}. ${new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate()}.`,
      quarterly: `${d.getFullYear()}년 제${q}분기(${(q - 1) * 3 + 1}. 1. ~ ${q * 3}. ${new Date(d.getFullYear(), q * 3, 0).getDate()}.)`,
      annual: `${d.getFullYear()}. 1. 1. ~ ${d.getFullYear()}. 12. 31.`,
      brief: `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}. 현재`,
      official: `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}. 현재`,
    }[type] || "";
    return {
      now: d, year: d.getFullYear(), month: d.getMonth() + 1, day: d.getDate(), quarter: q,
      dday, timePct, monthsLeft, period, type,
      dateKr: `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}.`,
    };
  }

  /* ── 분석(개조식 자동 문안) ──────────────────────────────────────────── */
  function analyze(bms, wbs, asset, hr, ctx) {
    const { timePct, monthsLeft, dday } = ctx;
    const gap = timePct - bms.execRate;
    const monthlyNeed = monthsLeft > 0 ? bms.totalRemain / monthsLeft : 0;
    const wbsGap = wbs.overall.plannedRate - wbs.overall.actualRate;
    const zero = bms.projects.filter((p) => p.rate < 1 && p.budget >= 3e8);
    const high = bms.projects.filter((p) => p.rate >= 80);
    const low = bms.projects.filter((p) => p.rate > 0 && p.rate < 30 && p.budget >= 3e8);

    const summary = [
      `예산집행: 총사업비 ${f1(eok(bms.totalBudget))}억원 중 ${f1(eok(bms.totalExec))}억원 집행(집행률 ${f2(bms.execRate)}%), 잔액 ${f1(eok(bms.totalRemain))}억원`,
      `공정추진: 계획공정률 ${f1(wbs.overall.plannedRate)}% 대비 실적공정률 ${f1(wbs.overall.actualRate)}%(${f1(Math.abs(wbsGap))}%p ${wbsGap > 0 ? "미달" : "초과"}), 완료 ${wbs.overall.done}건·진행 ${wbs.overall.inProg}건·지연 ${wbs.overall.delayed}건`,
      `인력·재산: 참여인력 ${hr.total}명(가동 ${hr.active}명, ${hr.orgs.length || 4}개 기관), 취득재산 ${asset.total}점(취득가액 ${f1(eok(asset.value))}억원)`,
      `관리쟁점: 사업기간 소진율 ${f1(timePct)}% 대비 집행률 ${f2(bms.execRate)}%로 ${f1(Math.abs(gap))}%p ${gap > 0 ? "미달" : "상회"}, 준공까지 ${monthsLeft}개월(D-${dday}) 잔여`,
    ];

    const budgetPoints = [
      `총사업비 ${f1(eok(bms.totalBudget))}억원 중 누계집행 ${f1(eok(bms.totalExec))}억원(집행률 ${f2(bms.execRate)}%), 집행잔액 ${f1(eok(bms.totalRemain))}억원`,
      `잔여 ${monthsLeft}개월 내 집행 완료를 위해 월평균 ${f1(eok(monthlyNeed))}억원 규모 집행 필요`,
      high.length ? `집행률 80% 이상 ${high.length}개 단위사업 — 검수·산출물 확정 및 정산자료 정비 단계로 전환` : `집행률 80% 이상 단위사업 없음 — 계약·집행 가속 필요`,
      zero.length ? `미집행(3억원 이상) ${zero.length}개 단위사업 — 발주계획 확정 및 계약체결 마감선 관리 필요` : `3억원 이상 미집행 단위사업 없음`,
    ];
    const wbsPoints = [
      `전체 ${wbs.overall.total}건 중 완료 ${wbs.overall.done}건(${f1((wbs.overall.done / (wbs.overall.total || 1)) * 100)}%)·진행 ${wbs.overall.inProg}건·지연 ${wbs.overall.delayed}건·대기 ${wbs.overall.waiting}건`,
      `계획공정률 ${f1(wbs.overall.plannedRate)}% 대비 실적공정률 ${f1(wbs.overall.actualRate)}%(달성률 ${f1(wbs.overall.achieveRate)}%)`,
      wbs.overall.delayed > 0 ? `지연 ${wbs.overall.delayed}건에 대한 원인분석 및 만회공정(Catch-up) 수립·이행 필요` : `지연 공정 없음 — 현 공정 유지 관리`,
    ];
    const over = (hr.roster || []).filter((p) => p.ratio > 100);
    const hrPoints = [
      `${hr.orgs.length || 4}개 기관 ${hr.total}명 등록(실인원 ${hr.headcount}명), 가동인력 ${hr.active}명(가동률 ${f1(hr.rate)}%)`,
      hr.hasRoster ? `참여율 합계 ${f1(hr.activeManRatio)}%(가동인력 기준) — 인월(M/M) 환산 시 약 ${f1(hr.activeManRatio / 100)}인 상당` : `참여인력 명부 미연계 — 인건비 정산 증빙 확보를 위해 명부 데이터 연계 필요`,
      `직접보조사업자는 PMO·발주지원·정산, 간접보조사업자는 실증·연구·기술자문 분담 수행`,
      `인건비는 실지급액 × 참여기간 × 참여율 기준 산정, 4대보험 및 급여이체 증빙 구비`,
    ];
    if (over.length) hrPoints.push(`참여율 합계 100% 초과 인력 ${over.length}명(${over.slice(0, 3).map((p) => `${p.name} ${f1(p.ratio)}%`).join(", ")}${over.length > 3 ? " 등" : ""}) — 인건비 중복계상 방지를 위한 참여율 조정 및 정정 필요`);
    const assetPoints = [
      `취득재산 ${asset.total}점, 취득가액 ${f1(eok(asset.value))}억원(사용 ${asset.inUse}점·대기 ${asset.standby}점)`,
      `자산관리대장 등재율 ${f1(asset.issuedRate)}% — 라벨링·실사·배치현황 관리 병행`,
      asset.major.length ? `「보조금법」 제35조 중요재산(취득가액 5천만원 이상) ${asset.major.length}건 별도 관리 중` : `취득가액 5천만원 이상 중요재산 해당 없음(자료 미연계 시 확인 필요)`,
    ];

    const issues = [];
    issues.push({
      issue: `사업기간 소진율(${f1(timePct)}%) 대비 예산집행률(${f2(bms.execRate)}%) 격차`,
      impact: `${f1(Math.abs(gap))}%p ${gap > 0 ? "미달" : "상회"}, 잔여 ${f1(eok(bms.totalRemain))}억원 집행 소요`,
      action: "미집행 단위사업 발주 가속, 비목간 전용 및 사업계획 변경승인 검토",
      grade: gap > 15 ? "매우 높음" : gap > 8 ? "높음" : "보통",
    });
    zero.slice(0, 3).forEach((p) => issues.push({
      issue: `${p.name} 미집행(예산 ${f1(eok(p.budget))}억원)`,
      impact: "공정·집행 동반 지연으로 사업기간 내 준공 곤란 우려",
      action: "발주계획 수립 및 행정절차(사전규격공개·협상에 의한 계약) 즉시 착수",
      grade: p.budget >= 2e9 ? "높음" : "보통",
    }));
    if (wbs.overall.delayed > 3) issues.push({
      issue: `WBS 지연공정 ${wbs.overall.delayed}건(계획 대비 ${f1(wbsGap)}%p 미달)`,
      impact: "전체 공정 만회 부담 가중 및 후속공정 연쇄 지연",
      action: "지연 원인분석 후 만회공정 재수립, 주간 단위 이행점검",
      grade: wbs.overall.delayed > 8 ? "높음" : "보통",
    });

    const achievements = [
      `예산집행관리시스템(BMS) 기준 누계집행 ${f1(eok(bms.totalExec))}억원 달성(${bms.itemCount}개 세부항목 실시간 관리)`,
      `WBS ${wbs.overall.total}건 중 완료 ${wbs.overall.done}건 관리, 실적공정률 ${f1(wbs.overall.actualRate)}%(달성률 ${f1(wbs.overall.achieveRate)}%)`,
      `${hr.orgs.length || 4}개 기관 참여인력 ${hr.total}명 투입·운영 및 기관별 인건비 정산 연계`,
      `취득재산 ${asset.total}점(${f1(eok(asset.value))}억원) 자산관리대장 등재(등재율 ${f1(asset.issuedRate)}%)`,
    ];

    const plans = [
      zero.length ? `미집행 ${zero.length}개 단위사업 발주 절차 착수 및 계약체결 마감선 확정` : `잔여 계약건 검수·준공처리 및 산출물 확정`,
      wbs.overall.delayed > 0 ? `지연 ${wbs.overall.delayed}건 만회공정 수립·이행 및 주간 점검체계 운영` : `현행 공정 유지 및 마일스톤 관리 지속`,
      `사업 준공 이후 3년간 운영단계(2027~2029) 대비 조례·예산·운영주체 확보 절차 추진`,
      `정산 대비 인건비·재산취득 증빙 정비 및 보조금 집행 적정성 사전점검`,
    ];

    const recos = [];
    if ((hr.roster || []).some((p) => p.ratio > 100)) recos.push(`참여율 정정 — 합계 100% 초과 인력의 참여율 재산정 및 사업계획 변경(참여인력 변경) 절차 이행`);
    if (gap > 5) recos.push(`집행 가속 — 잔여 ${monthsLeft}개월 기준 월평균 ${f1(eok(monthlyNeed))}억원 집행목표 설정 및 이행관리`);
    if (zero.length) recos.push(`발주일정 역산관리 — 준공기한 기준 검수·정산 소요기간 확보를 위한 계약체결 마감선 명확화`);
    if (wbs.overall.delayed > 0) recos.push(`공정 만회 — 지연 ${wbs.overall.delayed}건 만회계획 수립 및 마일스톤 재정렬`);
    recos.push(`정산 대비 — 인건비 실투입 증빙 및 중요재산 취득 증빙의 보조금관리시스템 연계 점검`);
    recos.push(`재원별 집행실적의 실계정 기준 산출체계 마련 — 현행 교부비율 안분 추정치의 정산 부적합 해소`);

    const judgment = `사업기간 소진율 대비 예산집행률 격차(${f1(Math.abs(gap))}%p) 및 공정지연 ${wbs.overall.delayed}건의 동시 관리가 사업 성패의 핵심 요인임. 잔여 ${monthsLeft}개월간 미집행 단위사업의 발주 가속과 정산증빙 정비를 병행 추진할 필요가 있음.`;

    return {
      summary, budgetPoints, wbsPoints, hrPoints, assetPoints,
      achievements, issues, recos, plans, judgment,
      zero, high, low, monthlyNeed, gap, wbsGap,
    };
  }

  return { processBMS, processWBS, processAsset, processHR, normalizeHR, parseHRHtml, analyze, buildCtx, eok, f1, f2, rate };
});

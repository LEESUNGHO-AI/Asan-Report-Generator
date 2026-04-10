#!/usr/bin/env node
/**
 * 아산시 강소형 스마트시티 주간 진도 보고서 자동 생성
 * v1.0 — BMS + WBS JSON → DOCX
 */

const fs = require("fs");
const https = require("https");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, LevelFormat, PageBreak, PageNumber
} = require("docx");

// ── Config ──
const BMS_URL = "https://leesungho-ai.github.io/Asan-Smart-City-Budget-Management-System-BMS-/data/budget.json";
const WBS_SUM_URL = "https://leesungho-ai.github.io/Asan-Smartcity-WBS/data/summary-data.json";
const WBS_DATA_URL = "https://leesungho-ai.github.io/Asan-Smartcity-WBS/data/wbs-data.json";

const PROJECT_START = new Date("2023-12-01");
const PROJECT_END = new Date("2026-12-31");
const NOW = new Date();
const KST = new Date(NOW.getTime() + 9 * 3600000);
const TODAY_STR = KST.toISOString().slice(0, 10);
const DDAY = Math.ceil((PROJECT_END - NOW) / 86400000);
const ELAPSED = Math.ceil((NOW - PROJECT_START) / 86400000);
const TOTAL_DAYS = Math.ceil((PROJECT_END - PROJECT_START) / 86400000);
const TIME_PCT = ((ELAPSED / TOTAL_DAYS) * 100).toFixed(1);

// Week number
const weekStart = new Date(KST);
weekStart.setDate(weekStart.getDate() - weekStart.getDay() + 1); // Monday
const weekEnd = new Date(weekStart);
weekEnd.setDate(weekEnd.getDate() + 4); // Friday
const weekNum = Math.ceil((KST - new Date(KST.getFullYear(), 0, 1)) / 604800000);
const WEEK_LABEL = `W${weekNum} (${weekStart.toISOString().slice(5, 10)} ~ ${weekEnd.toISOString().slice(5, 10)})`;

// BMS Unit Project Mapping
const BMS_UNIT_MAP = {
  "스마트 공공 WIFI": 1,
  "아산시 강소형 스마트시티 네트워크 구축": 1,
  "모바일 전자시민증 플랫폼 / 인프라": 2,
  "데이터기반 AI 융복합 서비스 구축": 2,
  "디지털 노마드접수/운영 및 거래관리": 2,
  "국제표준 디지털링크 공유 플랫폼": 2,
  "이노베이션 센터/ 관제 시스템 구축": 3,
  "디지털 OASIS SPOT": 4,
  "무인매장": 4,
  "SDDC Platform 구축": 5,
  "AI통합관제 및 운영 플랫폼 / 인프라": 6,
  "디지털OASIS 정보관리 시스템": 7,
  "수요응답형 DRT 서비스 운영 플랫폼 구축": 8,
  "수요응답형 DRT 서비스 운영 HW 구축": 8,
  "정보통신감리": 9,
  "스마트폴&디스플레이": 10,
  "메타버스 플랫폼": 11,
};
const UNIT_NAMES = {
  1: "유무선 네트워크 구축", 2: "서비스 인프라 플랫폼", 3: "이노베이션 센터 구축",
  4: "디지털 OASIS SPOT", 5: "SDDC Platform 구축", 6: "AI 통합관제 플랫폼",
  7: "디지털 OASIS 정보관리", 8: "DRT 수요응답형 교통", 9: "감리용역 (신설)",
  10: "스마트폴&디스플레이", 11: "메타버스 플랫폼",
};

// ── HTTP Fetch ──
function fetchJSON(url) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: { "User-Agent": "Asan-Report/1.0" } }, (res) => {
      let data = "";
      res.on("data", (c) => (data += c));
      res.on("end", () => {
        try { resolve(JSON.parse(data)); } catch (e) { reject(e); }
      });
    }).on("error", reject);
  });
}

// ── Data Processing ──
function processBMS(bms) {
  const s = bms.summary;
  const totalBudget = (s["총사업비"] || 0) / 1e8;
  const totalExec = (s["총집행액"] || 0) / 1e8;
  const totalRemain = (s["총잔액"] || 0) / 1e8;
  const execRate = s["전체집행률"] || 0;

  // Bimok
  const CLEAN = {
    "인건비(110)": "인건비", "운영비(210)": "운영비", "여비(220)": "여비",
    "연구개발비(260)": "연구개발비", "사업비배분(320)": "사업비배분",
    "사업비 배분(320)": "사업비배분", "유형자산(430)": "유형자산",
    "무형자산(440)": "무형자산(SW)", "건설비(420)": "건설비", "기타": "기타",
  };
  const merged = {};
  for (const b of (bms.bimok_summary || [])) {
    const name = CLEAN[b["비목"]] || b["비목"];
    if (!merged[name]) merged[name] = { b: 0, e: 0 };
    merged[name].b += (b["예산"] || 0) / 1e8;
    merged[name].e += (b["집행"] || 0) / 1e8;
  }
  const cats = [];
  for (const name of ["인건비", "운영비", "여비", "연구개발비", "유형자산", "무형자산(SW)", "건설비", "사업비배분", "기타"]) {
    if (!merged[name]) continue;
    const m = merged[name];
    const rate = m.b ? ((m.e / m.b) * 100).toFixed(1) : "0.0";
    cats.push({ name, budget: m.b.toFixed(1), exec: m.e.toFixed(2), rate });
  }

  // Unit projects
  const units = {};
  let commonB = 0, commonE = 0;
  for (const it of (bms.items || [])) {
    const num = BMS_UNIT_MAP[it["항목명"]];
    const exec = it["집행액"] || it["사용금액합계"] || it["사용금액"] || 0;
    if (num) {
      if (!units[num]) units[num] = { b: 0, e: 0 };
      units[num].b += (it["총예산"] || 0) / 1e8;
      units[num].e += exec / 1e8;
    } else {
      commonB += (it["총예산"] || 0) / 1e8;
      commonE += exec / 1e8;
    }
  }
  const projects = [];
  for (const num of Object.keys(UNIT_NAMES).map(Number).sort((a, b) => a - b)) {
    const u = units[num] || { b: 0, e: 0 };
    const rate = u.b ? ((u.e / u.b) * 100).toFixed(1) : "0.0";
    projects.push({ num, name: UNIT_NAMES[num], budget: u.b.toFixed(2), exec: u.e.toFixed(2), rate });
  }

  return { totalBudget, totalExec, totalRemain, execRate, cats, projects, commonB, commonE };
}

function processWBS(wbsSum, wbsData) {
  const t = wbsSum.total;
  const overall = {
    total: t.total, done: t.done, inProg: t.inProg, delayed: t.delayed,
    waiting: t.waiting, actualRate: t.actualRate, achieveRate: t.achieveRate,
    plannedRate: t.plannedRate,
  };

  const services = [];
  for (const r of (wbsData.items || [])) {
    if (r.level === "1" && r.weight > 0) {
      services.push({
        name: r.name, weight: r.weight,
        planned: r.plannedRate, actual: r.actualRate,
        deviation: r.deviation,
      });
    }
  }
  services.sort((a, b) => b.actual - a.actual);

  // Delayed items
  const delayed = [];
  for (const r of (wbsData.items || [])) {
    if (r.status === "지연" && r.level !== "1") {
      delayed.push({ name: r.name, category: r.category, org: r.organization, deviation: r.deviation });
    }
  }

  return { overall, services, delayed };
}

function generateIssues(bmsData, wbsResult) {
  const issues = [];
  const gap = parseFloat(TIME_PCT) - bmsData.execRate;
  if (gap > 25)
    issues.push({ level: "긴급", text: `예산 집행률(${bmsData.execRate}%)이 기간 소진율(${TIME_PCT}%) 대비 ${gap.toFixed(0)}%p 부족` });
  for (const p of bmsData.projects) {
    if (parseFloat(p.budget) >= 10 && parseFloat(p.rate) < 5)
      issues.push({ level: "주의", text: `#${p.num} ${p.name} (${p.budget}억) 집행률 ${p.rate}%` });
  }
  if (wbsResult.delayed.length > 5)
    issues.push({ level: "주의", text: `WBS 지연 작업 ${wbsResult.delayed.length}건 — 집중 관리 필요` });
  if (DDAY < 300)
    issues.push({ level: "정보", text: `사업 종료까지 D-${DDAY} (${(DDAY / 30).toFixed(0)}개월)` });
  return issues;
}

// ── DOCX Generation ──
function buildDoc(bmsData, wbsResult, issues) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };
  const headerShading = { fill: "1B3A5C", type: ShadingType.CLEAR };
  const altShading = { fill: "F0F5FA", type: ShadingType.CLEAR };
  const TABLE_W = 9026; // A4 with 1440 margins

  function hdrCell(text, w) {
    return new TableCell({
      borders, width: { size: w, type: WidthType.DXA }, shading: headerShading, margins: cellMargins,
      verticalAlign: "center",
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, bold: true, font: "맑은 고딕", size: 18, color: "FFFFFF" })] })],
    });
  }
  function dataCell(text, w, opts = {}) {
    return new TableCell({
      borders, width: { size: w, type: WidthType.DXA }, margins: cellMargins,
      shading: opts.shade ? altShading : undefined,
      verticalAlign: "center",
      children: [new Paragraph({
        alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({ text: String(text), font: "맑은 고딕", size: 18, bold: opts.bold, color: opts.color })],
      })],
    });
  }

  const children = [];

  // ── Cover ──
  children.push(new Paragraph({ spacing: { before: 2400 } }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "아산시 강소형 스마트시티 조성사업", font: "맑은 고딕", size: 40, bold: true, color: "1B3A5C" })],
  }));
  children.push(new Paragraph({ spacing: { before: 200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "주간 진도 보고서", font: "맑은 고딕", size: 52, bold: true, color: "1B3A5C" })],
  }));
  children.push(new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: WEEK_LABEL, font: "맑은 고딕", size: 28, color: "666666" })],
  }));
  children.push(new Paragraph({ spacing: { before: 200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: `작성일: ${TODAY_STR}`, font: "맑은 고딕", size: 24, color: "888888" })],
  }));
  children.push(new Paragraph({ spacing: { before: 1200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "제일엔지니어링 PMO팀", font: "맑은 고딕", size: 24, color: "444444" })],
  }));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ── 1. 총괄 현황 ──
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: "1. 총괄 현황", font: "맑은 고딕" })],
  }));

  const kpiCols = [2256, 2256, 2258, 2256];
  children.push(new Table({
    width: { size: TABLE_W, type: WidthType.DXA }, columnWidths: kpiCols,
    rows: [
      new TableRow({ children: [
        hdrCell("예산 집행률", kpiCols[0]), hdrCell("WBS 공정률", kpiCols[1]),
        hdrCell("기간 소진율", kpiCols[2]), hdrCell("D-Day", kpiCols[3]),
      ]}),
      new TableRow({ children: [
        dataCell(`${bmsData.execRate}%`, kpiCols[0], { align: AlignmentType.CENTER, bold: true, color: "2E75B6" }),
        dataCell(`${wbsResult.overall.actualRate}%`, kpiCols[1], { align: AlignmentType.CENTER, bold: true, color: "7030A0" }),
        dataCell(`${TIME_PCT}%`, kpiCols[2], { align: AlignmentType.CENTER, bold: true, color: "ED7D31" }),
        dataCell(`D-${DDAY}`, kpiCols[3], { align: AlignmentType.CENTER, bold: true, color: "C00000" }),
      ]}),
      new TableRow({ children: [
        dataCell(`${bmsData.totalExec.toFixed(1)}억 / ${bmsData.totalBudget.toFixed(0)}억`, kpiCols[0], { align: AlignmentType.CENTER, shade: true }),
        dataCell(`완료 ${wbsResult.overall.done} / 전체 ${wbsResult.overall.total}건`, kpiCols[1], { align: AlignmentType.CENTER, shade: true }),
        dataCell(`${ELAPSED}일 / ${TOTAL_DAYS}일`, kpiCols[2], { align: AlignmentType.CENTER, shade: true }),
        dataCell(`잔여 ${DDAY}일`, kpiCols[3], { align: AlignmentType.CENTER, shade: true }),
      ]}),
    ],
  }));
  children.push(new Paragraph({ spacing: { after: 200 } }));

  // ── 2. 예산 집행 현황 ──
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: "2. 예산 집행 현황 (비목별)", font: "맑은 고딕" })],
  }));

  const catCols = [2000, 1800, 1800, 1800, 1626];
  children.push(new Table({
    width: { size: TABLE_W, type: WidthType.DXA }, columnWidths: catCols,
    rows: [
      new TableRow({ children: [
        hdrCell("비목", catCols[0]), hdrCell("예산(억)", catCols[1]),
        hdrCell("집행(억)", catCols[2]), hdrCell("잔액(억)", catCols[3]), hdrCell("집행률", catCols[4]),
      ]}),
      ...bmsData.cats.map((c, i) => {
        const remain = (parseFloat(c.budget) - parseFloat(c.exec)).toFixed(1);
        const rateColor = parseFloat(c.rate) >= 80 ? "00B050" : (parseFloat(c.rate) >= 30 ? "ED7D31" : "C00000");
        return new TableRow({ children: [
          dataCell(c.name, catCols[0], { bold: true, shade: i % 2 === 1 }),
          dataCell(`${c.budget}억`, catCols[1], { align: AlignmentType.RIGHT, shade: i % 2 === 1 }),
          dataCell(`${c.exec}억`, catCols[2], { align: AlignmentType.RIGHT, shade: i % 2 === 1 }),
          dataCell(`${remain}억`, catCols[3], { align: AlignmentType.RIGHT, shade: i % 2 === 1 }),
          dataCell(`${c.rate}%`, catCols[4], { align: AlignmentType.CENTER, bold: true, color: rateColor, shade: i % 2 === 1 }),
        ]});
      }),
    ],
  }));
  children.push(new Paragraph({ spacing: { after: 200 } }));

  // ── 3. 단위사업별 집행 현황 ──
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: "3. 단위사업별 집행 현황", font: "맑은 고딕" })],
  }));

  const projCols = [500, 2800, 1500, 1500, 1200, 1526];
  children.push(new Table({
    width: { size: TABLE_W, type: WidthType.DXA }, columnWidths: projCols,
    rows: [
      new TableRow({ children: [
        hdrCell("#", projCols[0]), hdrCell("사업명", projCols[1]),
        hdrCell("예산(억)", projCols[2]), hdrCell("집행(억)", projCols[3]),
        hdrCell("잔액(억)", projCols[4]), hdrCell("집행률", projCols[5]),
      ]}),
      ...bmsData.projects.map((p, i) => {
        const remain = (parseFloat(p.budget) - parseFloat(p.exec)).toFixed(1);
        const rateColor = parseFloat(p.rate) >= 80 ? "00B050" : (parseFloat(p.rate) >= 30 ? "ED7D31" : "C00000");
        return new TableRow({ children: [
          dataCell(p.num, projCols[0], { align: AlignmentType.CENTER, shade: i % 2 === 1 }),
          dataCell(p.name, projCols[1], { bold: true, shade: i % 2 === 1 }),
          dataCell(`${p.budget}`, projCols[2], { align: AlignmentType.RIGHT, shade: i % 2 === 1 }),
          dataCell(`${p.exec}`, projCols[3], { align: AlignmentType.RIGHT, shade: i % 2 === 1 }),
          dataCell(`${remain}`, projCols[4], { align: AlignmentType.RIGHT, shade: i % 2 === 1 }),
          dataCell(`${p.rate}%`, projCols[5], { align: AlignmentType.CENTER, bold: true, color: rateColor, shade: i % 2 === 1 }),
        ]});
      }),
      // Common expenses
      new TableRow({ children: [
        dataCell("", projCols[0], { shade: true }),
        dataCell("공통경비 (인건비/운영비 등)", projCols[1], { bold: true, shade: true }),
        dataCell(bmsData.commonB.toFixed(1), projCols[2], { align: AlignmentType.RIGHT, shade: true }),
        dataCell(bmsData.commonE.toFixed(1), projCols[3], { align: AlignmentType.RIGHT, shade: true }),
        dataCell((bmsData.commonB - bmsData.commonE).toFixed(1), projCols[4], { align: AlignmentType.RIGHT, shade: true }),
        dataCell(`${bmsData.commonB ? ((bmsData.commonE / bmsData.commonB) * 100).toFixed(1) : "0.0"}%`, projCols[5], { align: AlignmentType.CENTER, bold: true, shade: true }),
      ]}),
    ],
  }));
  children.push(new Paragraph({ spacing: { after: 200 } }));

  // ── 4. WBS 공정률 ──
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: "4. WBS 공정률 현황", font: "맑은 고딕" })],
  }));

  // WBS Summary
  const wbsSumCols = [1504, 1504, 1506, 1504, 1504, 1504];
  children.push(new Table({
    width: { size: TABLE_W, type: WidthType.DXA }, columnWidths: wbsSumCols,
    rows: [
      new TableRow({ children: [
        hdrCell("전체", wbsSumCols[0]), hdrCell("완료", wbsSumCols[1]), hdrCell("진행", wbsSumCols[2]),
        hdrCell("지연", wbsSumCols[3]), hdrCell("대기", wbsSumCols[4]), hdrCell("달성률", wbsSumCols[5]),
      ]}),
      new TableRow({ children: [
        dataCell(`${wbsResult.overall.total}건`, wbsSumCols[0], { align: AlignmentType.CENTER, bold: true }),
        dataCell(`${wbsResult.overall.done}건`, wbsSumCols[1], { align: AlignmentType.CENTER, bold: true, color: "00B050" }),
        dataCell(`${wbsResult.overall.inProg}건`, wbsSumCols[2], { align: AlignmentType.CENTER, bold: true, color: "2E75B6" }),
        dataCell(`${wbsResult.overall.delayed}건`, wbsSumCols[3], { align: AlignmentType.CENTER, bold: true, color: "C00000" }),
        dataCell(`${wbsResult.overall.waiting}건`, wbsSumCols[4], { align: AlignmentType.CENTER, bold: true, color: "ED7D31" }),
        dataCell(`${wbsResult.overall.achieveRate}%`, wbsSumCols[5], { align: AlignmentType.CENTER, bold: true, color: "7030A0" }),
      ]}),
    ],
  }));
  children.push(new Paragraph({ spacing: { after: 200 } }));

  // WBS Level-1
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text: "4-1. Level-1 공정률 (가중평균)", font: "맑은 고딕" })],
  }));

  const svcCols = [3200, 1200, 1500, 1500, 1626];
  children.push(new Table({
    width: { size: TABLE_W, type: WidthType.DXA }, columnWidths: svcCols,
    rows: [
      new TableRow({ children: [
        hdrCell("분류", svcCols[0]), hdrCell("가중치", svcCols[1]),
        hdrCell("계획(%)", svcCols[2]), hdrCell("실적(%)", svcCols[3]), hdrCell("편차(%p)", svcCols[4]),
      ]}),
      ...wbsResult.services.map((s, i) => {
        const devColor = s.deviation >= 0 ? "00B050" : "C00000";
        return new TableRow({ children: [
          dataCell(s.name, svcCols[0], { bold: true, shade: i % 2 === 1 }),
          dataCell(`${s.weight}%`, svcCols[1], { align: AlignmentType.CENTER, shade: i % 2 === 1 }),
          dataCell(`${s.planned}%`, svcCols[2], { align: AlignmentType.CENTER, shade: i % 2 === 1 }),
          dataCell(`${s.actual}%`, svcCols[3], { align: AlignmentType.CENTER, bold: true, shade: i % 2 === 1 }),
          dataCell(`${s.deviation > 0 ? "+" : ""}${s.deviation}%p`, svcCols[4], { align: AlignmentType.CENTER, bold: true, color: devColor, shade: i % 2 === 1 }),
        ]});
      }),
    ],
  }));
  children.push(new Paragraph({ spacing: { after: 200 } }));

  // ── 5. 이슈 및 리스크 ──
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: "5. 이슈 및 리스크", font: "맑은 고딕" })],
  }));

  const issCols = [1200, 7826];
  children.push(new Table({
    width: { size: TABLE_W, type: WidthType.DXA }, columnWidths: issCols,
    rows: [
      new TableRow({ children: [hdrCell("수준", issCols[0]), hdrCell("내용", issCols[1])] }),
      ...issues.map((iss, i) => {
        const lvColor = iss.level === "긴급" ? "C00000" : (iss.level === "주의" ? "ED7D31" : "2E75B6");
        return new TableRow({ children: [
          dataCell(iss.level, issCols[0], { align: AlignmentType.CENTER, bold: true, color: lvColor, shade: i % 2 === 1 }),
          dataCell(iss.text, issCols[1], { shade: i % 2 === 1 }),
        ]});
      }),
    ],
  }));
  children.push(new Paragraph({ spacing: { after: 200 } }));

  // ── 6. 지연 작업 목록 (상위 10건) ──
  if (wbsResult.delayed.length > 0) {
    children.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text: `6. WBS 지연 작업 (${wbsResult.delayed.length}건)`, font: "맑은 고딕" })],
    }));
    const delCols = [3500, 2500, 1500, 1526];
    const delRows = wbsResult.delayed.slice(0, 10);
    children.push(new Table({
      width: { size: TABLE_W, type: WidthType.DXA }, columnWidths: delCols,
      rows: [
        new TableRow({ children: [
          hdrCell("작업명", delCols[0]), hdrCell("대분류", delCols[1]),
          hdrCell("담당기관", delCols[2]), hdrCell("편차(%p)", delCols[3]),
        ]}),
        ...delRows.map((d, i) => new TableRow({ children: [
          dataCell(d.name, delCols[0], { shade: i % 2 === 1 }),
          dataCell(d.category || "-", delCols[1], { shade: i % 2 === 1 }),
          dataCell(d.org || "-", delCols[2], { shade: i % 2 === 1 }),
          dataCell(`${d.deviation}%p`, delCols[3], { align: AlignmentType.CENTER, bold: true, color: "C00000", shade: i % 2 === 1 }),
        ]})),
      ],
    }));
  }

  // ── Build Document ──
  return new Document({
    styles: {
      default: { document: { run: { font: "맑은 고딕", size: 20 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "맑은 고딕", color: "1B3A5C" },
          paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "맑은 고딕", color: "2E75B6" },
          paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: `아산시 강소형 스마트시티 | 주간 진도 보고서 ${WEEK_LABEL}`, font: "맑은 고딕", size: 16, color: "999999" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "제일엔지니어링 PMO | ", font: "맑은 고딕", size: 16, color: "999999" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "맑은 고딕", size: 16, color: "999999" }),
              new TextRun({ text: " / ", font: "맑은 고딕", size: 16, color: "999999" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "맑은 고딕", size: 16, color: "999999" }),
            ],
          })],
        }),
      },
      children,
    }],
  });
}

// ── Main ──
async function main() {
  console.log("=" .repeat(60));
  console.log(`📋 주간 진도 보고서 생성 — ${TODAY_STR} ${WEEK_LABEL}`);
  console.log("=" .repeat(60));

  console.log("\n📦 데이터 수집...");
  const [bms, wbsSum, wbsData] = await Promise.all([
    fetchJSON(BMS_URL), fetchJSON(WBS_SUM_URL), fetchJSON(WBS_DATA_URL),
  ]);
  console.log("  BMS:", bms.updated_at);
  console.log("  WBS:", wbsSum.meta.generatedAtKst);

  console.log("\n📊 데이터 처리...");
  const bmsData = processBMS(bms);
  const wbsResult = processWBS(wbsSum, wbsData);
  const issues = generateIssues(bmsData, wbsResult);

  console.log(`  집행률: ${bmsData.execRate}%`);
  console.log(`  WBS: ${wbsResult.overall.actualRate}% (달성률 ${wbsResult.overall.achieveRate}%)`);
  console.log(`  지연: ${wbsResult.delayed.length}건, 이슈: ${issues.length}건`);

  console.log("\n📄 DOCX 생성...");
  const doc = buildDoc(bmsData, wbsResult, issues);
  const buffer = await Packer.toBuffer(doc);

  const outDir = "reports";
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  const filename = `주간진도보고서_${TODAY_STR}_${WEEK_LABEL.replace(/[\/\s()~]/g, "_")}.docx`;
  const filepath = `${outDir}/${filename}`;
  fs.writeFileSync(filepath, buffer);
  console.log(`  ✅ ${filepath} (${(buffer.length / 1024).toFixed(0)} KB)`);

  // Also save JSON snapshot
  const snapshot = {
    generated: KST.toISOString(),
    week: WEEK_LABEL,
    dday: DDAY,
    bms: { execRate: bmsData.execRate, totalExec: bmsData.totalExec, totalBudget: bmsData.totalBudget },
    wbs: wbsResult.overall,
    issues,
    filename,
  };
  fs.writeFileSync(`${outDir}/latest.json`, JSON.stringify(snapshot, null, 2));
  console.log(`  ✅ ${outDir}/latest.json`);

  console.log("\n🎉 완료!");
}

main().catch((e) => { console.error("ERROR:", e); process.exit(1); });

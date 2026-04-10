#!/usr/bin/env node
/**
 * 아산시 강소형 스마트시티 주간 진도 보고서 자동 생성 v2.0
 * ═══════════════════════════════════════════════════════
 * Phase 2: Claude API 분석 엔진 통합
 * - BMS + WBS JSON 데이터 수집
 * - Claude API로 AI 분석 (이슈 도출, 리스크 판단, 차주 계획)
 * - DOCX 보고서 자동 생성
 */

const fs = require("fs");
const https = require("https");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, LevelFormat, PageBreak, PageNumber
} = require("docx");

// ══════ Config ══════
const BMS_URL = "https://leesungho-ai.github.io/Asan-Smart-City-Budget-Management-System-BMS-/data/budget.json";
const WBS_SUM_URL = "https://leesungho-ai.github.io/Asan-Smartcity-WBS/data/summary-data.json";
const WBS_DATA_URL = "https://leesungho-ai.github.io/Asan-Smartcity-WBS/data/wbs-data.json";
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || "";

const PROJECT_START = new Date("2023-12-01");
const PROJECT_END = new Date("2026-12-31");
const NOW = new Date();
const KST = new Date(NOW.getTime() + 9 * 3600000);
const TODAY_STR = KST.toISOString().slice(0, 10);
const DDAY = Math.ceil((PROJECT_END - NOW) / 86400000);
const ELAPSED = Math.ceil((NOW - PROJECT_START) / 86400000);
const TOTAL_DAYS = Math.ceil((PROJECT_END - PROJECT_START) / 86400000);
const TIME_PCT = ((ELAPSED / TOTAL_DAYS) * 100).toFixed(1);
const weekNum = Math.ceil((KST - new Date(KST.getFullYear(), 0, 1)) / 604800000);
const ws = new Date(KST); ws.setDate(ws.getDate() - ws.getDay() + 1);
const we = new Date(ws); we.setDate(we.getDate() + 4);
const WEEK_LABEL = `W${weekNum} (${ws.toISOString().slice(5,10)} ~ ${we.toISOString().slice(5,10)})`;

// BMS → Unit Project mapping
const BMS_UNIT_MAP = {
  "스마트 공공 WIFI":1, "아산시 강소형 스마트시티 네트워크 구축":1,
  "모바일 전자시민증 플랫폼 / 인프라":2, "데이터기반 AI 융복합 서비스 구축":2,
  "디지털 노마드접수/운영 및 거래관리":2, "국제표준 디지털링크 공유 플랫폼":2,
  "이노베이션 센터/ 관제 시스템 구축":3, "디지털 OASIS SPOT":4, "무인매장":4,
  "SDDC Platform 구축":5, "AI통합관제 및 운영 플랫폼 / 인프라":6,
  "디지털OASIS 정보관리 시스템":7, "수요응답형 DRT 서비스 운영 플랫폼 구축":8,
  "수요응답형 DRT 서비스 운영 HW 구축":8, "정보통신감리":9,
  "스마트폴&디스플레이":10, "메타버스 플랫폼":11,
};
const UNIT_NAMES = {
  1:"유무선 네트워크 구축",2:"서비스 인프라 플랫폼",3:"이노베이션 센터 구축",
  4:"디지털 OASIS SPOT",5:"SDDC Platform 구축",6:"AI 통합관제 플랫폼",
  7:"디지털 OASIS 정보관리",8:"DRT 수요응답형 교통",9:"감리용역 (신설)",
  10:"스마트폴&디스플레이",11:"메타버스 플랫폼",
};

// ══════ HTTP ══════
function fetchJSON(url) {
  return new Promise((resolve, reject) => {
    https.get(url, {headers:{"User-Agent":"Asan-Report/2.0"}}, res => {
      let d=""; res.on("data",c=>d+=c);
      res.on("end",()=>{ try{resolve(JSON.parse(d))}catch(e){reject(e)} });
    }).on("error",reject);
  });
}

function callClaudeAPI(prompt) {
  if (!ANTHROPIC_API_KEY) {
    console.log("  ⚠️  ANTHROPIC_API_KEY 없음 → AI 분석 건너뜀");
    return Promise.resolve(null);
  }
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 4000,
      messages: [{ role: "user", content: prompt }],
    });
    const req = https.request({
      hostname: "api.anthropic.com", port: 443, path: "/v1/messages",
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "Content-Length": Buffer.byteLength(body),
      },
    }, res => {
      let d=""; res.on("data",c=>d+=c);
      res.on("end",()=>{
        try {
          const r = JSON.parse(d);
          if (r.content && r.content[0]) resolve(r.content[0].text);
          else { console.log("  ⚠️  Claude API 응답 이상:", d.slice(0,200)); resolve(null); }
        } catch(e) { console.log("  ⚠️  Claude API 파싱 실패:", e.message); resolve(null); }
      });
    });
    req.on("error", e => { console.log("  ⚠️  Claude API 오류:", e.message); resolve(null); });
    req.setTimeout(30000, () => { req.destroy(); resolve(null); });
    req.write(body); req.end();
  });
}

// ══════ Data Processing ══════
function processBMS(bms) {
  const s = bms.summary;
  const totalBudget = (s["총사업비"]||0)/1e8;
  const totalExec = (s["총집행액"]||0)/1e8;
  const totalRemain = (s["총잔액"]||0)/1e8;
  const execRate = s["전체집행률"]||0;

  const CLEAN = {"인건비(110)":"인건비","운영비(210)":"운영비","여비(220)":"여비",
    "연구개발비(260)":"연구개발비","사업비배분(320)":"사업비배분","사업비 배분(320)":"사업비배분",
    "유형자산(430)":"유형자산","무형자산(440)":"무형자산(SW)","건설비(420)":"건설비","기타":"기타"};
  const merged = {};
  for (const b of (bms.bimok_summary||[])) {
    const nm = CLEAN[b["비목"]]||b["비목"];
    if (!merged[nm]) merged[nm]={b:0,e:0};
    merged[nm].b += (b["예산"]||0)/1e8; merged[nm].e += (b["집행"]||0)/1e8;
  }
  const cats = [];
  for (const nm of ["인건비","운영비","여비","연구개발비","유형자산","무형자산(SW)","건설비","사업비배분","기타"]) {
    if (!merged[nm]) continue;
    const m = merged[nm];
    cats.push({name:nm, budget:m.b.toFixed(1), exec:m.e.toFixed(2), rate: m.b?(m.e/m.b*100).toFixed(1):"0.0"});
  }

  const units = {}; let commonB=0, commonE=0;
  for (const it of (bms.items||[])) {
    const num = BMS_UNIT_MAP[it["항목명"]];
    const ex = it["집행액"]||it["사용금액합계"]||it["사용금액"]||0;
    if (num) { if (!units[num]) units[num]={b:0,e:0}; units[num].b+=(it["총예산"]||0)/1e8; units[num].e+=ex/1e8; }
    else { commonB+=(it["총예산"]||0)/1e8; commonE+=ex/1e8; }
  }
  const projects = [];
  for (const num of Object.keys(UNIT_NAMES).map(Number).sort((a,b)=>a-b)) {
    const u=units[num]||{b:0,e:0};
    projects.push({num,name:UNIT_NAMES[num],budget:u.b.toFixed(2),exec:u.e.toFixed(2),rate:u.b?(u.e/u.b*100).toFixed(1):"0.0"});
  }
  return {totalBudget,totalExec,totalRemain,execRate,cats,projects,commonB,commonE};
}

function processWBS(wbsSum, wbsData) {
  const t = wbsSum.total;
  const overall = {total:t.total,done:t.done,inProg:t.inProg,delayed:t.delayed,
    waiting:t.waiting,actualRate:t.actualRate,achieveRate:t.achieveRate,plannedRate:t.plannedRate};
  const services = [];
  for (const r of (wbsData.items||[])) {
    if (r.level==="1" && r.weight>0)
      services.push({name:r.name,weight:r.weight,planned:r.plannedRate,actual:r.actualRate,deviation:r.deviation});
  }
  services.sort((a,b)=>b.actual-a.actual);
  const delayed = [];
  for (const r of (wbsData.items||[])) {
    if (r.status==="지연" && r.level!=="1")
      delayed.push({name:r.name,category:r.category,org:r.organization,deviation:r.deviation});
  }
  return {overall,services,delayed};
}

// ══════ Claude AI Analysis ══════
async function getAIAnalysis(bmsData, wbsResult) {
  console.log("\n🤖 Claude AI 분석 시작...");

  const dataContext = `
당신은 아산시 강소형 스마트시티 조성사업(240억원, 2023.12~2026.12)의 PMO 분석 전문가입니다.
아래 실시간 데이터를 기반으로 한국어로 주간 보고서의 AI 분석 섹션을 작성해주세요.

## 현재 현황
- 사업기간: ${ELAPSED}일 경과 / ${TOTAL_DAYS}일 (소진율 ${TIME_PCT}%)
- 잔여일: D-${DDAY}
- 예산 집행률: ${bmsData.execRate}% (${bmsData.totalExec.toFixed(1)}억 / ${bmsData.totalBudget.toFixed(0)}억)
- WBS 공정률: ${wbsResult.overall.actualRate}% (달성률 ${wbsResult.overall.achieveRate}%)
- WBS: 전체 ${wbsResult.overall.total}건, 완료 ${wbsResult.overall.done}, 진행 ${wbsResult.overall.inProg}, 지연 ${wbsResult.overall.delayed}, 대기 ${wbsResult.overall.waiting}

## 단위사업별 집행 현황
${bmsData.projects.map(p => `#${p.num} ${p.name}: 예산 ${p.budget}억, 집행 ${p.exec}억 (${p.rate}%)`).join("\n")}
공통경비: 예산 ${bmsData.commonB.toFixed(1)}억, 집행 ${bmsData.commonE.toFixed(1)}억

## WBS Level-1 공정률
${wbsResult.services.map(s => `${s.name}: 가중치 ${s.weight}%, 계획 ${s.planned}%, 실적 ${s.actual}%, 편차 ${s.deviation}%p`).join("\n")}

## 지연 작업 (${wbsResult.delayed.length}건)
${wbsResult.delayed.slice(0,10).map(d => `${d.name} (${d.category||"-"}, ${d.org||"-"}, 편차 ${d.deviation}%p)`).join("\n")}

## 비목별 예산
${bmsData.cats.map(c => `${c.name}: 예산 ${c.budget}억, 집행 ${c.exec}억 (${c.rate}%)`).join("\n")}

아래 JSON 형식으로만 응답해주세요. 다른 텍스트 없이 JSON만 출력하세요:
{
  "executive_summary": "3~4문장의 이번 주 핵심 요약",
  "key_achievements": ["이번 주 주요 성과 3~5개"],
  "risk_analysis": [
    {"level": "긴급|주의|정보", "title": "리스크 제목", "description": "상세 설명", "action": "대응 방안"}
  ],
  "budget_insight": "예산 집행 관련 분석 (2~3문장, 기간소진율 대비 집행률 격차 포함)",
  "wbs_insight": "WBS 공정 관련 분석 (2~3문장, 지연 원인 분석 포함)",
  "next_week_plan": ["차주 핵심 과제 3~5개"],
  "recommendations": ["PMO 권고사항 3~5개"]
}`;

  const result = await callClaudeAPI(dataContext);
  if (!result) return null;

  try {
    // Clean JSON from markdown fences
    const cleaned = result.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
    const parsed = JSON.parse(cleaned);
    console.log("  ✅ AI 분석 완료");
    return parsed;
  } catch (e) {
    console.log("  ⚠️  AI 응답 JSON 파싱 실패:", e.message);
    console.log("  원본:", result.slice(0, 300));
    return null;
  }
}

// ══════ Fallback Analysis (no API key) ══════
function getFallbackAnalysis(bmsData, wbsResult) {
  const gap = (parseFloat(TIME_PCT) - bmsData.execRate).toFixed(0);
  const issues = [];
  if (gap > 25) issues.push({level:"긴급",title:`예산 집행률 격차 ${gap}%p`,description:`기간 소진율(${TIME_PCT}%) 대비 집행률(${bmsData.execRate}%)이 ${gap}%p 부족`,action:"미집행 대형 사업 조속 발주 및 선급금 집행"});
  for (const p of bmsData.projects) {
    if (parseFloat(p.budget)>=10 && parseFloat(p.rate)<5) issues.push({level:"주의",title:`#${p.num} ${p.name} 미집행`,description:`예산 ${p.budget}억 중 집행률 ${p.rate}%`,action:"발주 일정 확인 및 집행 계획 수립"});
  }
  if (wbsResult.delayed.length>5) issues.push({level:"주의",title:`WBS 지연 ${wbsResult.delayed.length}건`,description:"계획 대비 실적 미달 작업 다수",action:"지연 원인 분석 및 만회 계획 수립"});
  if (DDAY<300) issues.push({level:"정보",title:`잔여기간 D-${DDAY}`,description:`사업 종료까지 약 ${(DDAY/30).toFixed(0)}개월`,action:"준공 로드맵 점검"});

  return {
    executive_summary: `금주 사업 현황: 예산 집행률 ${bmsData.execRate}%, WBS 공정률 ${wbsResult.overall.actualRate}%. 기간 소진율(${TIME_PCT}%) 대비 집행률 격차가 ${gap}%p로 집중 관리 필요. 잔여 D-${DDAY}일.`,
    key_achievements: ["BMS/WBS 데이터 자동 동기화 운영 중","통합 포털 v4.2 실시간 현행화 완료","GitHub Actions 기반 30분 자동 갱신 정상 가동"],
    risk_analysis: issues,
    budget_insight: `총 ${bmsData.totalBudget.toFixed(0)}억 중 ${bmsData.totalExec.toFixed(1)}억 집행(${bmsData.execRate}%). 기간 소진율(${TIME_PCT}%) 대비 ${gap}%p 부족. 잔여 ${bmsData.totalRemain.toFixed(1)}억을 ${(DDAY/30).toFixed(0)}개월 내 집행해야 하므로 월평균 ${(bmsData.totalRemain/(DDAY/30)).toFixed(1)}억 집행 필요.`,
    wbs_insight: `전체 ${wbsResult.overall.total}건 중 완료 ${wbsResult.overall.done}건(${(wbsResult.overall.done/wbsResult.overall.total*100).toFixed(0)}%), 지연 ${wbsResult.overall.delayed}건. 서비스 구축(가중치 63%) 실적 ${wbsResult.services.find(s=>s.weight>=60)?.actual||0}%로 전체 공정률에 가장 큰 영향.`,
    next_week_plan: ["미착수 대형사업 발주 진행 점검","WBS 지연 작업 만회 계획 수립","월간 집행 실적 점검 및 보고"],
    recommendations: ["미집행 대형 사업(OASIS SPOT, DRT 등) 발주 가속화","비목간 전용 검토 (여비→운영비 등)","준공 일정 역산 기반 마일스톤 재설정"],
  };
}

// ══════ DOCX Generation ══════
function buildDoc(bmsData, wbsResult, analysis) {
  const border = {style:BorderStyle.SINGLE,size:1,color:"BBBBBB"};
  const borders = {top:border,bottom:border,left:border,right:border};
  const cellM = {top:60,bottom:60,left:100,right:100};
  const hdrShade = {fill:"1B3A5C",type:ShadingType.CLEAR};
  const altShade = {fill:"F0F5FA",type:ShadingType.CLEAR};
  const aiShade = {fill:"FFF8E7",type:ShadingType.CLEAR};
  const TW = 9026;

  function hC(t,w) { return new TableCell({borders,width:{size:w,type:WidthType.DXA},shading:hdrShade,margins:cellM,verticalAlign:"center",children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:t,bold:true,font:"맑은 고딕",size:18,color:"FFFFFF"})]})]}) }
  function dC(t,w,o={}) { return new TableCell({borders,width:{size:w,type:WidthType.DXA},margins:cellM,shading:o.shade?altShade:o.aiShade?aiShade:undefined,verticalAlign:"center",children:[new Paragraph({alignment:o.align||AlignmentType.LEFT,children:[new TextRun({text:String(t),font:"맑은 고딕",size:18,bold:o.bold,color:o.color})]})]}) }

  const ch = [];

  // ── 표지 ──
  ch.push(new Paragraph({spacing:{before:2400}}));
  ch.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"아산시 강소형 스마트시티 조성사업",font:"맑은 고딕",size:40,bold:true,color:"1B3A5C"})]}));
  ch.push(new Paragraph({spacing:{before:200},alignment:AlignmentType.CENTER,children:[new TextRun({text:"주간 진도 보고서",font:"맑은 고딕",size:52,bold:true,color:"1B3A5C"})]}));
  ch.push(new Paragraph({spacing:{before:400},alignment:AlignmentType.CENTER,children:[new TextRun({text:WEEK_LABEL,font:"맑은 고딕",size:28,color:"666666"})]}));
  ch.push(new Paragraph({spacing:{before:100},alignment:AlignmentType.CENTER,children:[new TextRun({text:`작성일: ${TODAY_STR}  |  AI 분석: ${ANTHROPIC_API_KEY?"Claude Sonnet":"Fallback"}`,font:"맑은 고딕",size:20,color:"999999"})]}));
  ch.push(new Paragraph({spacing:{before:1200},alignment:AlignmentType.CENTER,children:[new TextRun({text:"제일엔지니어링 PMO팀",font:"맑은 고딕",size:24,color:"444444"})]}));
  ch.push(new Paragraph({children:[new PageBreak()]}));

  // ── 0. AI 경영진 요약 ──
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"Executive Summary",font:"맑은 고딕"})]}));
  ch.push(new Paragraph({spacing:{before:80,after:120},shading:{fill:"FFF3CD",type:ShadingType.CLEAR},
    children:[new TextRun({text:"🤖 AI 분석 기반 요약",font:"맑은 고딕",size:18,bold:true,color:"856404"}),
              new TextRun({text:`  |  ${TODAY_STR} 기준 데이터`,font:"맑은 고딕",size:16,color:"856404"})]}));
  ch.push(new Paragraph({spacing:{after:200},children:[new TextRun({text:analysis.executive_summary,font:"맑은 고딕",size:20})]}));

  // KPI 테이블
  const kC = [2256,2256,2258,2256];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:kC,rows:[
    new TableRow({children:[hC("예산 집행률",kC[0]),hC("WBS 공정률",kC[1]),hC("기간 소진율",kC[2]),hC("D-Day",kC[3])]}),
    new TableRow({children:[
      dC(`${bmsData.execRate}%`,kC[0],{align:AlignmentType.CENTER,bold:true,color:"2E75B6"}),
      dC(`${wbsResult.overall.actualRate}%`,kC[1],{align:AlignmentType.CENTER,bold:true,color:"7030A0"}),
      dC(`${TIME_PCT}%`,kC[2],{align:AlignmentType.CENTER,bold:true,color:"ED7D31"}),
      dC(`D-${DDAY}`,kC[3],{align:AlignmentType.CENTER,bold:true,color:"C00000"})
    ]}),
    new TableRow({children:[
      dC(`${bmsData.totalExec.toFixed(1)}억 / ${bmsData.totalBudget.toFixed(0)}억`,kC[0],{align:AlignmentType.CENTER,shade:true}),
      dC(`완료 ${wbsResult.overall.done} / ${wbsResult.overall.total}건`,kC[1],{align:AlignmentType.CENTER,shade:true}),
      dC(`${ELAPSED}일 / ${TOTAL_DAYS}일`,kC[2],{align:AlignmentType.CENTER,shade:true}),
      dC(`잔여 ${DDAY}일 (${(DDAY/30).toFixed(0)}개월)`,kC[3],{align:AlignmentType.CENTER,shade:true})
    ]}),
  ]}));
  ch.push(new Paragraph({spacing:{after:200}}));

  // ── 1. 주요 성과 (AI) ──
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"1. 금주 주요 성과",font:"맑은 고딕"})]}));
  for (const a of (analysis.key_achievements||[])) {
    ch.push(new Paragraph({spacing:{before:40},numbering:{reference:"bullets",level:0},children:[new TextRun({text:a,font:"맑은 고딕",size:20})]}));
  }
  ch.push(new Paragraph({spacing:{after:200}}));

  // ── 2. 예산 집행 현황 ──
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"2. 예산 집행 현황",font:"맑은 고딕"})]}));

  // AI 예산 인사이트
  ch.push(new Paragraph({spacing:{before:80,after:120},shading:{fill:"E8F4FD",type:ShadingType.CLEAR},
    children:[new TextRun({text:"💡 ",font:"맑은 고딕",size:18}),new TextRun({text:analysis.budget_insight,font:"맑은 고딕",size:18,color:"0C5460"})]}));

  // 비목별
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:"2-1. 비목별 집행 현황",font:"맑은 고딕"})]}));
  const cC = [2000,1800,1800,1800,1626];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:cC,rows:[
    new TableRow({children:[hC("비목",cC[0]),hC("예산(억)",cC[1]),hC("집행(억)",cC[2]),hC("잔액(억)",cC[3]),hC("집행률",cC[4])]}),
    ...bmsData.cats.map((c,i) => {
      const rm=(parseFloat(c.budget)-parseFloat(c.exec)).toFixed(1);
      const rc=parseFloat(c.rate)>=80?"00B050":(parseFloat(c.rate)>=30?"ED7D31":"C00000");
      return new TableRow({children:[dC(c.name,cC[0],{bold:true,shade:i%2===1}),dC(`${c.budget}억`,cC[1],{align:AlignmentType.RIGHT,shade:i%2===1}),dC(`${c.exec}억`,cC[2],{align:AlignmentType.RIGHT,shade:i%2===1}),dC(`${rm}억`,cC[3],{align:AlignmentType.RIGHT,shade:i%2===1}),dC(`${c.rate}%`,cC[4],{align:AlignmentType.CENTER,bold:true,color:rc,shade:i%2===1})]});
    }),
  ]}));
  ch.push(new Paragraph({spacing:{after:200}}));

  // 단위사업별
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:"2-2. 단위사업별 집행 현황",font:"맑은 고딕"})]}));
  const pC = [500,2800,1500,1500,1200,1526];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:pC,rows:[
    new TableRow({children:[hC("#",pC[0]),hC("사업명",pC[1]),hC("예산(억)",pC[2]),hC("집행(억)",pC[3]),hC("잔액",pC[4]),hC("집행률",pC[5])]}),
    ...bmsData.projects.map((p,i) => {
      const rm=(parseFloat(p.budget)-parseFloat(p.exec)).toFixed(1);
      const rc=parseFloat(p.rate)>=80?"00B050":(parseFloat(p.rate)>=30?"ED7D31":"C00000");
      return new TableRow({children:[dC(p.num,pC[0],{align:AlignmentType.CENTER,shade:i%2===1}),dC(p.name,pC[1],{bold:true,shade:i%2===1}),dC(p.budget,pC[2],{align:AlignmentType.RIGHT,shade:i%2===1}),dC(p.exec,pC[3],{align:AlignmentType.RIGHT,shade:i%2===1}),dC(rm,pC[4],{align:AlignmentType.RIGHT,shade:i%2===1}),dC(`${p.rate}%`,pC[5],{align:AlignmentType.CENTER,bold:true,color:rc,shade:i%2===1})]});
    }),
    new TableRow({children:[dC("",pC[0],{shade:true}),dC("공통경비",pC[1],{bold:true,shade:true}),dC(bmsData.commonB.toFixed(1),pC[2],{align:AlignmentType.RIGHT,shade:true}),dC(bmsData.commonE.toFixed(1),pC[3],{align:AlignmentType.RIGHT,shade:true}),dC((bmsData.commonB-bmsData.commonE).toFixed(1),pC[4],{align:AlignmentType.RIGHT,shade:true}),dC(`${bmsData.commonB?(bmsData.commonE/bmsData.commonB*100).toFixed(1):"0.0"}%`,pC[5],{align:AlignmentType.CENTER,bold:true,shade:true})]}),
  ]}));
  ch.push(new Paragraph({children:[new PageBreak()]}));

  // ── 3. WBS 공정률 ──
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"3. WBS 공정률 현황",font:"맑은 고딕"})]}));
  ch.push(new Paragraph({spacing:{before:80,after:120},shading:{fill:"E8F4FD",type:ShadingType.CLEAR},
    children:[new TextRun({text:"💡 ",font:"맑은 고딕",size:18}),new TextRun({text:analysis.wbs_insight,font:"맑은 고딕",size:18,color:"0C5460"})]}));

  const wC = [1504,1504,1506,1504,1504,1504];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:wC,rows:[
    new TableRow({children:[hC("전체",wC[0]),hC("완료",wC[1]),hC("진행",wC[2]),hC("지연",wC[3]),hC("대기",wC[4]),hC("달성률",wC[5])]}),
    new TableRow({children:[
      dC(`${wbsResult.overall.total}건`,wC[0],{align:AlignmentType.CENTER,bold:true}),
      dC(`${wbsResult.overall.done}건`,wC[1],{align:AlignmentType.CENTER,bold:true,color:"00B050"}),
      dC(`${wbsResult.overall.inProg}건`,wC[2],{align:AlignmentType.CENTER,bold:true,color:"2E75B6"}),
      dC(`${wbsResult.overall.delayed}건`,wC[3],{align:AlignmentType.CENTER,bold:true,color:"C00000"}),
      dC(`${wbsResult.overall.waiting}건`,wC[4],{align:AlignmentType.CENTER,bold:true,color:"ED7D31"}),
      dC(`${wbsResult.overall.achieveRate}%`,wC[5],{align:AlignmentType.CENTER,bold:true,color:"7030A0"})
    ]}),
  ]}));
  ch.push(new Paragraph({spacing:{after:120}}));

  // Level-1
  const sC = [3200,1200,1500,1500,1626];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:sC,rows:[
    new TableRow({children:[hC("분류",sC[0]),hC("가중치",sC[1]),hC("계획(%)",sC[2]),hC("실적(%)",sC[3]),hC("편차(%p)",sC[4])]}),
    ...wbsResult.services.map((s,i) => {
      const dc=s.deviation>=0?"00B050":"C00000";
      return new TableRow({children:[dC(s.name,sC[0],{bold:true,shade:i%2===1}),dC(`${s.weight}%`,sC[1],{align:AlignmentType.CENTER,shade:i%2===1}),dC(`${s.planned}%`,sC[2],{align:AlignmentType.CENTER,shade:i%2===1}),dC(`${s.actual}%`,sC[3],{align:AlignmentType.CENTER,bold:true,shade:i%2===1}),dC(`${s.deviation>0?"+":""}${s.deviation}%p`,sC[4],{align:AlignmentType.CENTER,bold:true,color:dc,shade:i%2===1})]});
    }),
  ]}));
  ch.push(new Paragraph({spacing:{after:200}}));

  // ── 4. 리스크 분석 (AI) ──
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"4. 리스크 분석 (AI)",font:"맑은 고딕"})]}));
  const rC = [1000,2500,3200,2326];
  const risks = analysis.risk_analysis||[];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:rC,rows:[
    new TableRow({children:[hC("수준",rC[0]),hC("리스크",rC[1]),hC("상세",rC[2]),hC("대응 방안",rC[3])]}),
    ...risks.map((r,i) => {
      const lc=r.level==="긴급"?"C00000":(r.level==="주의"?"ED7D31":"2E75B6");
      return new TableRow({children:[
        dC(r.level,rC[0],{align:AlignmentType.CENTER,bold:true,color:lc,shade:i%2===1}),
        dC(r.title,rC[1],{bold:true,shade:i%2===1}),
        dC(r.description,rC[2],{shade:i%2===1}),
        dC(r.action,rC[3],{shade:i%2===1}),
      ]});
    }),
  ]}));
  ch.push(new Paragraph({spacing:{after:200}}));

  // ── 5. 차주 계획 (AI) ──
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"5. 차주 핵심 과제",font:"맑은 고딕"})]}));
  for (const t of (analysis.next_week_plan||[])) {
    ch.push(new Paragraph({numbering:{reference:"numbers",level:0},children:[new TextRun({text:t,font:"맑은 고딕",size:20})]}));
  }
  ch.push(new Paragraph({spacing:{after:200}}));

  // ── 6. PMO 권고사항 (AI) ──
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"6. PMO 권고사항",font:"맑은 고딕"})]}));
  for (const r of (analysis.recommendations||[])) {
    ch.push(new Paragraph({numbering:{reference:"bullets",level:0},children:[new TextRun({text:r,font:"맑은 고딕",size:20})]}));
  }

  // ── 지연 작업 (있으면) ──
  if (wbsResult.delayed.length > 0) {
    ch.push(new Paragraph({children:[new PageBreak()]}));
    ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:`[부록] WBS 지연 작업 (${wbsResult.delayed.length}건)`,font:"맑은 고딕"})]}));
    const dCols = [3500,2500,1500,1526];
    ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:dCols,rows:[
      new TableRow({children:[hC("작업명",dCols[0]),hC("대분류",dCols[1]),hC("담당",dCols[2]),hC("편차",dCols[3])]}),
      ...wbsResult.delayed.slice(0,15).map((d,i) => new TableRow({children:[
        dC(d.name,dCols[0],{shade:i%2===1}),dC(d.category||"-",dCols[1],{shade:i%2===1}),
        dC(d.org||"-",dCols[2],{shade:i%2===1}),dC(`${d.deviation}%p`,dCols[3],{align:AlignmentType.CENTER,bold:true,color:"C00000",shade:i%2===1})
      ]})),
    ]}));
  }

  return new Document({
    numbering:{config:[
      {reference:"bullets",levels:[{level:0,format:LevelFormat.BULLET,text:"\u2022",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]},
      {reference:"numbers",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]},
    ]},
    styles:{
      default:{document:{run:{font:"맑은 고딕",size:20}}},
      paragraphStyles:[
        {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
          run:{size:28,bold:true,font:"맑은 고딕",color:"1B3A5C"},
          paragraph:{spacing:{before:360,after:120},outlineLevel:0}},
        {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
          run:{size:24,bold:true,font:"맑은 고딕",color:"2E75B6"},
          paragraph:{spacing:{before:240,after:100},outlineLevel:1}},
      ],
    },
    sections:[{
      properties:{page:{size:{width:11906,height:16838},margin:{top:1440,bottom:1440,left:1440,right:1440}}},
      headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:`아산시 스마트시티 | 주간 보고서 ${WEEK_LABEL}`,font:"맑은 고딕",size:16,color:"999999"})]})]})},
      footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[
        new TextRun({text:"제일엔지니어링 PMO | ",font:"맑은 고딕",size:16,color:"999999"}),
        new TextRun({children:[PageNumber.CURRENT],font:"맑은 고딕",size:16,color:"999999"}),
        new TextRun({text:" / ",font:"맑은 고딕",size:16,color:"999999"}),
        new TextRun({children:[PageNumber.TOTAL_PAGES],font:"맑은 고딕",size:16,color:"999999"}),
      ]})]})},
      children:ch,
    }],
  });
}

// ══════ Main ══════
async function main() {
  console.log("=".repeat(60));
  console.log(`📋 주간 진도 보고서 v2.0 (AI 분석 엔진) — ${TODAY_STR} ${WEEK_LABEL}`);
  console.log("=".repeat(60));

  console.log("\n📦 데이터 수집...");
  const [bms, wbsSum, wbsData] = await Promise.all([fetchJSON(BMS_URL), fetchJSON(WBS_SUM_URL), fetchJSON(WBS_DATA_URL)]);
  console.log(`  BMS: ${bms.updated_at}`);
  console.log(`  WBS: ${wbsSum.meta.generatedAtKst}`);

  console.log("\n📊 데이터 처리...");
  const bmsData = processBMS(bms);
  const wbsResult = processWBS(wbsSum, wbsData);
  console.log(`  집행률: ${bmsData.execRate}%, WBS: ${wbsResult.overall.actualRate}%`);
  console.log(`  지연: ${wbsResult.delayed.length}건`);

  // AI Analysis
  let analysis = await getAIAnalysis(bmsData, wbsResult);
  if (!analysis) {
    console.log("  📝 Fallback 분석 사용");
    analysis = getFallbackAnalysis(bmsData, wbsResult);
  }

  console.log("\n📄 DOCX 생성...");
  const doc = buildDoc(bmsData, wbsResult, analysis);
  const buffer = await Packer.toBuffer(doc);

  const outDir = "reports";
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, {recursive:true});
  const fn = `주간진도보고서_${TODAY_STR}_W${weekNum}.docx`;
  fs.writeFileSync(`${outDir}/${fn}`, buffer);
  console.log(`  ✅ ${outDir}/${fn} (${(buffer.length/1024).toFixed(0)} KB)`);

  const snapshot = {
    generated: KST.toISOString(), week: WEEK_LABEL, dday: DDAY,
    ai_engine: ANTHROPIC_API_KEY ? "claude-sonnet-4" : "fallback",
    bms: {execRate:bmsData.execRate,totalExec:bmsData.totalExec,totalBudget:bmsData.totalBudget},
    wbs: wbsResult.overall, issues: (analysis.risk_analysis||[]).length,
    filename: fn,
  };
  fs.writeFileSync(`${outDir}/latest.json`, JSON.stringify(snapshot,null,2));
  console.log(`  ✅ ${outDir}/latest.json`);
  console.log("\n🎉 완료!");
}

main().catch(e => { console.error("ERROR:", e); process.exit(1); });

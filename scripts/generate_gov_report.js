#!/usr/bin/env node
/**
 * 아산시 강소형 스마트시티 상위기관 보고서 자동 생성 v3.0
 * ═══════════════════════════════════════════════════════
 * 지원 보고서:
 *   monthly  - 월별 관리카드 (국토부, 익월 5일)
 *   quarter  - 분기보고서 (국토부, 분기종료 후 15일)
 *   annual   - 연간성과보고서 (국토부, 익년 1/31)
 * 
 * Usage: node generate_gov_report.js [monthly|quarter|annual]
 */

const fs = require("fs");
const https = require("https");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, LevelFormat, PageBreak, PageNumber
} = require("docx");

const TYPE = process.argv[2] || "monthly";
const API_KEY = process.env.ANTHROPIC_API_KEY || "";
const BMS_URL = "https://leesungho-ai.github.io/Asan-Smart-City-Budget-Management-System-BMS-/data/budget.json";
const WBS_SUM_URL = "https://leesungho-ai.github.io/Asan-Smartcity-WBS/data/summary-data.json";
const WBS_DATA_URL = "https://leesungho-ai.github.io/Asan-Smartcity-WBS/data/wbs-data.json";

const NOW = new Date(); const KST = new Date(NOW.getTime()+9*3600000);
const TODAY = KST.toISOString().slice(0,10);
const YEAR = KST.getFullYear(); const MONTH = KST.getMonth()+1;
const Q = Math.ceil(MONTH/3);
const PRJ_START = new Date("2023-12-01"); const PRJ_END = new Date("2026-12-31");
const DDAY = Math.ceil((PRJ_END-NOW)/86400000);
const ELAPSED = Math.ceil((NOW-PRJ_START)/86400000);
const TOTAL = Math.ceil((PRJ_END-PRJ_START)/86400000);
const TPCT = (ELAPSED/TOTAL*100).toFixed(1);

const REPORT_META = {
  monthly: {name:"월별 관리카드",to:"국토교통부 스마트시티종합계획실",due:`${MONTH===12?YEAR+1:YEAR}년 ${MONTH===12?1:MONTH+1}월 5일`,period:`${YEAR}년 ${MONTH}월`},
  quarter: {name:`${Q}분기 보고서`,to:"국토교통부 스마트시티종합계획실",due:`분기종료 후 15일`,period:`${YEAR}년 ${Q}분기`},
  annual:  {name:"연간 성과보고서",to:"국토교통부 스마트시티과",due:`${YEAR+1}년 1월 31일`,period:`${YEAR}년`},
};
const META = REPORT_META[TYPE];

const BMS_MAP = {
  "스마트 공공 WIFI":1,"아산시 강소형 스마트시티 네트워크 구축":1,
  "모바일 전자시민증 플랫폼 / 인프라":2,"데이터기반 AI 융복합 서비스 구축":2,
  "디지털 노마드접수/운영 및 거래관리":2,"국제표준 디지털링크 공유 플랫폼":2,
  "이노베이션 센터/ 관제 시스템 구축":3,"디지털 OASIS SPOT":4,"무인매장":4,
  "SDDC Platform 구축":5,"AI통합관제 및 운영 플랫폼 / 인프라":6,
  "디지털OASIS 정보관리 시스템":7,"수요응답형 DRT 서비스 운영 플랫폼 구축":8,
  "수요응답형 DRT 서비스 운영 HW 구축":8,"정보통신감리":9,
  "스마트폴&디스플레이":10,"메타버스 플랫폼":11,
};
const NAMES = {1:"유무선 네트워크 구축",2:"서비스 인프라 플랫폼",3:"이노베이션 센터 구축",
  4:"디지털 OASIS SPOT",5:"SDDC Platform 구축",6:"AI 통합관제 플랫폼",
  7:"디지털 OASIS 정보관리",8:"DRT 수요응답형 교통",9:"감리용역",
  10:"스마트폴&디스플레이",11:"메타버스 플랫폼"};

// ── HTTP/API ──
function fetchJSON(u){return new Promise((ok,no)=>{https.get(u,{headers:{"User-Agent":"Asan/3.0"}},r=>{let d="";r.on("data",c=>d+=c);r.on("end",()=>{try{ok(JSON.parse(d))}catch(e){no(e)}})}).on("error",no)})}

function callAI(prompt){
  if(!API_KEY) return Promise.resolve(null);
  return new Promise(ok=>{
    const b=JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:3000,messages:[{role:"user",content:prompt}]});
    const req=https.request({hostname:"api.anthropic.com",port:443,path:"/v1/messages",method:"POST",
      headers:{"Content-Type":"application/json","x-api-key":API_KEY,"anthropic-version":"2023-06-01","Content-Length":Buffer.byteLength(b)}},
      r=>{let d="";r.on("data",c=>d+=c);r.on("end",()=>{try{const j=JSON.parse(d);ok(j.content?.[0]?.text||null)}catch{ok(null)}})});
    req.on("error",()=>ok(null));req.setTimeout(60000,()=>{req.destroy();ok(null)});
    req.write(b);req.end();
  });
}

// ── Data ──
function processBMS(bms){
  const s=bms.summary;
  const CLEAN={"인건비(110)":"인건비","운영비(210)":"운영비","여비(220)":"여비","연구개발비(260)":"연구개발비",
    "사업비배분(320)":"사업비배분","사업비 배분(320)":"사업비배분","유형자산(430)":"유형자산",
    "무형자산(440)":"무형자산(SW)","건설비(420)":"건설비","기타":"기타"};
  const mg={};
  for(const b of bms.bimok_summary||[]){const n=CLEAN[b["비목"]]||b["비목"];if(!mg[n])mg[n]={b:0,e:0};mg[n].b+=(b["예산"]||0)/1e8;mg[n].e+=(b["집행"]||0)/1e8}
  const cats=[];
  for(const n of ["인건비","운영비","여비","연구개발비","유형자산","무형자산(SW)","건설비","사업비배분","기타"]){
    if(!mg[n])continue; const m=mg[n];
    cats.push({name:n,budget:m.b.toFixed(1),exec:m.e.toFixed(2),rate:m.b?(m.e/m.b*100).toFixed(1):"0.0"});
  }
  const units={};let cB=0,cE=0;
  for(const it of bms.items||[]){
    const num=BMS_MAP[it["항목명"]];const ex=it["집행액"]||it["사용금액합계"]||it["사용금액"]||0;
    if(num){if(!units[num])units[num]={b:0,e:0};units[num].b+=(it["총예산"]||0)/1e8;units[num].e+=ex/1e8}
    else{cB+=(it["총예산"]||0)/1e8;cE+=ex/1e8}
  }
  const prj=[];
  for(const n of Object.keys(NAMES).map(Number).sort((a,b)=>a-b)){
    const u=units[n]||{b:0,e:0};
    prj.push({num:n,name:NAMES[n],budget:u.b.toFixed(2),exec:u.e.toFixed(2),rate:u.b?(u.e/u.b*100).toFixed(1):"0.0"});
  }
  return{totalBudget:(s["총사업비"]||0)/1e8,totalExec:(s["총집행액"]||0)/1e8,
    totalRemain:(s["총잔액"]||0)/1e8,execRate:s["전체집행률"]||0,cats,projects:prj,commonB:cB,commonE:cE};
}

function processWBS(ws,wd){
  const t=ws.total;
  const svcs=[];
  for(const r of wd.items||[]){if(r.level==="1"&&r.weight>0)svcs.push({name:r.name,weight:r.weight,planned:r.plannedRate,actual:r.actualRate,deviation:r.deviation})}
  svcs.sort((a,b)=>b.actual-a.actual);
  const delayed=(wd.items||[]).filter(r=>r.status==="지연"&&r.level!=="1");
  return{overall:{total:t.total,done:t.done,inProg:t.inProg,delayed:t.delayed,waiting:t.waiting,
    actualRate:t.actualRate,achieveRate:t.achieveRate,plannedRate:t.plannedRate},svcs,delayed};
}

// ── DOCX Helpers ──
const TW=9026;
const bdr={style:BorderStyle.SINGLE,size:1,color:"BBBBBB"};
const borders={top:bdr,bottom:bdr,left:bdr,right:bdr};
const cm={top:60,bottom:60,left:100,right:100};
const hs={fill:"1B3A5C",type:ShadingType.CLEAR};
const as={fill:"F0F5FA",type:ShadingType.CLEAR};

function hC(t,w){return new TableCell({borders,width:{size:w,type:WidthType.DXA},shading:hs,margins:cm,verticalAlign:"center",children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:t,bold:true,font:"맑은 고딕",size:18,color:"FFFFFF"})]})]})}
function dC(t,w,o={}){return new TableCell({borders,width:{size:w,type:WidthType.DXA},margins:cm,shading:o.shade?as:undefined,verticalAlign:"center",children:[new Paragraph({alignment:o.align||AlignmentType.LEFT,children:[new TextRun({text:String(t),font:"맑은 고딕",size:18,bold:o.bold,color:o.color})]})]})}
function rc(v){return parseFloat(v)>=80?"00B050":(parseFloat(v)>=30?"ED7D31":"C00000")}

// ── Report Builders ──
function buildMonthly(bms,wbs,ai){
  const ch=[];
  // Cover
  ch.push(new Paragraph({spacing:{before:1800}}));
  ch.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:400},
    children:[new TextRun({text:"스마트시티 조성사업 월간 추진현황 보고",font:"맑은 고딕",size:44,bold:true,color:"1B3A5C"})]}));
  ch.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:600},
    children:[new TextRun({text:`(${META.period})`,font:"맑은 고딕",size:32,color:"444444"})]}));

  const infoCols=[3000,6026];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:infoCols,rows:[
    new TableRow({children:[hC("사업명",infoCols[0]),dC("아산시 강소형 스마트시티 조성사업",infoCols[1],{bold:true})]}),
    new TableRow({children:[hC("사업위치",infoCols[0]),dC("충청남도 아산시 도고면·배방읍 일원",infoCols[1])]}),
    new TableRow({children:[hC("사업기간",infoCols[0]),dC("2023.12 ~ 2026.12 (3년, 12개월 연장)",infoCols[1])]}),
    new TableRow({children:[hC("총사업비",infoCols[0]),dC(`240억원 (국비 120억 · 도비 28.8억 · 시비 91.2억)`,infoCols[1])]}),
    new TableRow({children:[hC("제출일",infoCols[0]),dC(TODAY,infoCols[1])]}),
    new TableRow({children:[hC("제출처",infoCols[0]),dC("아산시 스마트도시팀",infoCols[1])]}),
  ]}));
  ch.push(new Paragraph({spacing:{before:800},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:"아 산 시",font:"맑은 고딕",size:36,bold:true,color:"1B3A5C"})]}));
  ch.push(new Paragraph({children:[new PageBreak()]}));

  // 1. 총괄현황
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"1. 사업 총괄 현황",font:"맑은 고딕"})]}));
  const kC=[2256,2256,2258,2256];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:kC,rows:[
    new TableRow({children:[hC("예산 집행률",kC[0]),hC("WBS 공정률",kC[1]),hC("기간 소진율",kC[2]),hC("D-Day",kC[3])]}),
    new TableRow({children:[
      dC(`${bms.execRate}%`,kC[0],{align:AlignmentType.CENTER,bold:true,color:"2E75B6"}),
      dC(`${wbs.overall.actualRate}%`,kC[1],{align:AlignmentType.CENTER,bold:true,color:"7030A0"}),
      dC(`${TPCT}%`,kC[2],{align:AlignmentType.CENTER,bold:true,color:"ED7D31"}),
      dC(`D-${DDAY}`,kC[3],{align:AlignmentType.CENTER,bold:true,color:"C00000"})
    ]}),
  ]}));
  ch.push(new Paragraph({spacing:{after:200}}));

  // AI Summary
  if(ai?.executive_summary){
    ch.push(new Paragraph({spacing:{before:80,after:120},shading:{fill:"E8F4FD",type:ShadingType.CLEAR},
      children:[new TextRun({text:"💡 AI 분석: ",font:"맑은 고딕",size:18,bold:true,color:"0C5460"}),
                new TextRun({text:ai.executive_summary,font:"맑은 고딕",size:18,color:"0C5460"})]}));
  }

  // 2. 예산 집행
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"2. 예산 집행 현황",font:"맑은 고딕"})]}));
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:"2-1. 재원별 집행현황",font:"맑은 고딕"})]}));
  const srcCols=[1500,1500,1500,1500,1500,1526];
  const natBudget=120,doBudget=28.8,siBudget=91.2;
  const natRate=(bms.totalExec*natBudget/240/natBudget*100).toFixed(1);
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:srcCols,rows:[
    new TableRow({children:[hC("재원",srcCols[0]),hC("총예산",srcCols[1]),hC("집행액",srcCols[2]),hC("잔액",srcCols[3]),hC("집행률",srcCols[4]),hC("비고",srcCols[5])]}),
    new TableRow({children:[dC("국비(50%)",srcCols[0],{bold:true}),dC("120.0억",srcCols[1],{align:AlignmentType.RIGHT}),
      dC(`${(bms.totalExec*0.5).toFixed(1)}억`,srcCols[2],{align:AlignmentType.RIGHT}),
      dC(`${(120-bms.totalExec*0.5).toFixed(1)}억`,srcCols[3],{align:AlignmentType.RIGHT}),
      dC(`${bms.execRate}%`,srcCols[4],{align:AlignmentType.CENTER,bold:true,color:rc(bms.execRate)}),dC("",srcCols[5])]}),
    new TableRow({children:[dC("도비(12%)",srcCols[0],{bold:true,shade:true}),dC("28.8억",srcCols[1],{align:AlignmentType.RIGHT,shade:true}),
      dC(`${(bms.totalExec*0.12).toFixed(1)}억`,srcCols[2],{align:AlignmentType.RIGHT,shade:true}),
      dC(`${(28.8-bms.totalExec*0.12).toFixed(1)}억`,srcCols[3],{align:AlignmentType.RIGHT,shade:true}),
      dC(`${bms.execRate}%`,srcCols[4],{align:AlignmentType.CENTER,bold:true,shade:true}),dC("",srcCols[5],{shade:true})]}),
    new TableRow({children:[dC("시비(38%)",srcCols[0],{bold:true}),dC("91.2억",srcCols[1],{align:AlignmentType.RIGHT}),
      dC(`${(bms.totalExec*0.38).toFixed(1)}억`,srcCols[2],{align:AlignmentType.RIGHT}),
      dC(`${(91.2-bms.totalExec*0.38).toFixed(1)}억`,srcCols[3],{align:AlignmentType.RIGHT}),
      dC(`${bms.execRate}%`,srcCols[4],{align:AlignmentType.CENTER,bold:true}),dC("",srcCols[5])]}),
    new TableRow({children:[dC("합계",srcCols[0],{bold:true,shade:true}),dC("240.0억",srcCols[1],{align:AlignmentType.RIGHT,bold:true,shade:true}),
      dC(`${bms.totalExec.toFixed(1)}억`,srcCols[2],{align:AlignmentType.RIGHT,bold:true,shade:true}),
      dC(`${bms.totalRemain.toFixed(1)}억`,srcCols[3],{align:AlignmentType.RIGHT,bold:true,shade:true}),
      dC(`${bms.execRate}%`,srcCols[4],{align:AlignmentType.CENTER,bold:true,color:rc(bms.execRate),shade:true}),dC("",srcCols[5],{shade:true})]}),
  ]}));
  ch.push(new Paragraph({spacing:{after:200}}));

  // 비목별
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:"2-2. 비목별 집행현황",font:"맑은 고딕"})]}));
  const cC=[2000,1800,1800,1800,1626];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:cC,rows:[
    new TableRow({children:[hC("비목",cC[0]),hC("예산(억)",cC[1]),hC("집행(억)",cC[2]),hC("잔액(억)",cC[3]),hC("집행률",cC[4])]}),
    ...bms.cats.map((c,i)=>{const rm=(parseFloat(c.budget)-parseFloat(c.exec)).toFixed(1);
      return new TableRow({children:[dC(c.name,cC[0],{bold:true,shade:i%2===1}),dC(`${c.budget}`,cC[1],{align:AlignmentType.RIGHT,shade:i%2===1}),
        dC(`${c.exec}`,cC[2],{align:AlignmentType.RIGHT,shade:i%2===1}),dC(rm,cC[3],{align:AlignmentType.RIGHT,shade:i%2===1}),
        dC(`${c.rate}%`,cC[4],{align:AlignmentType.CENTER,bold:true,color:rc(c.rate),shade:i%2===1})]})
    }),
  ]}));
  ch.push(new Paragraph({children:[new PageBreak()]}));

  // 3. 세부사업별
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"3. 세부사업별 추진현황",font:"맑은 고딕"})]}));
  const pC=[500,2800,1300,1300,1300,1826];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:pC,rows:[
    new TableRow({children:[hC("#",pC[0]),hC("세부사업명",pC[1]),hC("예산(억)",pC[2]),hC("집행(억)",pC[3]),hC("집행률",pC[4]),hC("상태",pC[5])]}),
    ...bms.projects.map((p,i)=>{
      const st=parseFloat(p.rate)>=90?"🟢 완료":(parseFloat(p.rate)>0?"🔵 진행":"⬜ 미착수");
      return new TableRow({children:[dC(p.num,pC[0],{align:AlignmentType.CENTER,shade:i%2===1}),dC(p.name,pC[1],{bold:true,shade:i%2===1}),
        dC(p.budget,pC[2],{align:AlignmentType.RIGHT,shade:i%2===1}),dC(p.exec,pC[3],{align:AlignmentType.RIGHT,shade:i%2===1}),
        dC(`${p.rate}%`,pC[4],{align:AlignmentType.CENTER,bold:true,color:rc(p.rate),shade:i%2===1}),
        dC(st,pC[5],{align:AlignmentType.CENTER,shade:i%2===1})]})
    }),
  ]}));
  ch.push(new Paragraph({spacing:{after:200}}));

  // 4. WBS 공정률
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"4. WBS 공정 현황",font:"맑은 고딕"})]}));
  if(ai?.wbs_insight){
    ch.push(new Paragraph({spacing:{after:120},shading:{fill:"E8F4FD",type:ShadingType.CLEAR},
      children:[new TextRun({text:"💡 ",font:"맑은 고딕",size:18}),new TextRun({text:ai.wbs_insight,font:"맑은 고딕",size:18,color:"0C5460"})]}));
  }
  const wC=[1504,1504,1506,1504,1504,1504];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:wC,rows:[
    new TableRow({children:[hC("전체",wC[0]),hC("완료",wC[1]),hC("진행",wC[2]),hC("지연",wC[3]),hC("대기",wC[4]),hC("달성률",wC[5])]}),
    new TableRow({children:[dC(`${wbs.overall.total}건`,wC[0],{align:AlignmentType.CENTER,bold:true}),
      dC(`${wbs.overall.done}건`,wC[1],{align:AlignmentType.CENTER,bold:true,color:"00B050"}),
      dC(`${wbs.overall.inProg}건`,wC[2],{align:AlignmentType.CENTER,bold:true,color:"2E75B6"}),
      dC(`${wbs.overall.delayed}건`,wC[3],{align:AlignmentType.CENTER,bold:true,color:"C00000"}),
      dC(`${wbs.overall.waiting}건`,wC[4],{align:AlignmentType.CENTER,bold:true,color:"ED7D31"}),
      dC(`${wbs.overall.achieveRate}%`,wC[5],{align:AlignmentType.CENTER,bold:true,color:"7030A0"})]}),
  ]}));
  ch.push(new Paragraph({spacing:{after:200}}));

  // 5. 이슈
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"5. 주요 이슈 및 대응현황",font:"맑은 고딕"})]}));
  if(ai?.risk_analysis?.length){
    const rC=[1000,2500,3200,2326];
    ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:rC,rows:[
      new TableRow({children:[hC("수준",rC[0]),hC("이슈",rC[1]),hC("영향",rC[2]),hC("대응방안",rC[3])]}),
      ...ai.risk_analysis.map((r,i)=>{
        const lc=r.level==="긴급"?"C00000":(r.level==="주의"?"ED7D31":"2E75B6");
        return new TableRow({children:[dC(r.level,rC[0],{align:AlignmentType.CENTER,bold:true,color:lc,shade:i%2===1}),
          dC(r.title,rC[1],{bold:true,shade:i%2===1}),dC(r.description,rC[2],{shade:i%2===1}),dC(r.action,rC[3],{shade:i%2===1})]})
      }),
    ]}));
  }
  ch.push(new Paragraph({spacing:{after:200}}));

  // 6. 향후 계획
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"6. 향후 추진계획",font:"맑은 고딕"})]}));
  if(ai?.next_week_plan){
    for(const p of ai.next_week_plan)
      ch.push(new Paragraph({numbering:{reference:"bullets",level:0},children:[new TextRun({text:p,font:"맑은 고딕",size:20})]}));
  }
  ch.push(new Paragraph({spacing:{after:120}}));
  if(ai?.recommendations){
    ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:"6-1. PMO 권고사항",font:"맑은 고딕"})]}));
    for(const r of ai.recommendations)
      ch.push(new Paragraph({numbering:{reference:"bullets",level:0},children:[new TextRun({text:r,font:"맑은 고딕",size:20})]}));
  }

  return ch;
}

function buildQuarter(bms,wbs,ai){
  // 분기보고서는 월별 + 성과지표 + 분기별 비교 추가
  const ch = buildMonthly(bms,wbs,ai);

  // 추가: 성과지표 섹션
  ch.push(new Paragraph({children:[new PageBreak()]}));
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"7. 성과지표 달성 현황",font:"맑은 고딕"})]}));
  const kpiCols=[3000,2000,2000,2026];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:kpiCols,rows:[
    new TableRow({children:[hC("성과지표",kpiCols[0]),hC("목표",kpiCols[1]),hC("실적",kpiCols[2]),hC("달성률",kpiCols[3])]}),
    new TableRow({children:[dC("예산 집행률",kpiCols[0],{bold:true}),dC("100%",kpiCols[1],{align:AlignmentType.CENTER}),
      dC(`${bms.execRate}%`,kpiCols[2],{align:AlignmentType.CENTER}),dC(`${bms.execRate}%`,kpiCols[3],{align:AlignmentType.CENTER,bold:true,color:rc(bms.execRate)})]}),
    new TableRow({children:[dC("WBS 공정률",kpiCols[0],{bold:true,shade:true}),dC("100%",kpiCols[1],{align:AlignmentType.CENTER,shade:true}),
      dC(`${wbs.overall.actualRate}%`,kpiCols[2],{align:AlignmentType.CENTER,shade:true}),dC(`${wbs.overall.achieveRate}%`,kpiCols[3],{align:AlignmentType.CENTER,bold:true,color:rc(wbs.overall.achieveRate),shade:true})]}),
    new TableRow({children:[dC("서비스 구축 완료",kpiCols[0],{bold:true}),dC(`${wbs.overall.total}건`,kpiCols[1],{align:AlignmentType.CENTER}),
      dC(`${wbs.overall.done}건`,kpiCols[2],{align:AlignmentType.CENTER}),dC(`${(wbs.overall.done/wbs.overall.total*100).toFixed(0)}%`,kpiCols[3],{align:AlignmentType.CENTER,bold:true})]}),
  ]}));

  // Level-1 WBS 상세
  ch.push(new Paragraph({spacing:{before:200},heading:HeadingLevel.HEADING_2,children:[new TextRun({text:"7-1. WBS Level-1 분류별 공정률",font:"맑은 고딕"})]}));
  const sC=[3200,1200,1500,1500,1626];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:sC,rows:[
    new TableRow({children:[hC("분류",sC[0]),hC("가중치",sC[1]),hC("계획(%)",sC[2]),hC("실적(%)",sC[3]),hC("편차(%p)",sC[4])]}),
    ...wbs.svcs.map((s,i)=>{const dc=s.deviation>=0?"00B050":"C00000";
      return new TableRow({children:[dC(s.name,sC[0],{bold:true,shade:i%2===1}),dC(`${s.weight}%`,sC[1],{align:AlignmentType.CENTER,shade:i%2===1}),
        dC(`${s.planned}%`,sC[2],{align:AlignmentType.CENTER,shade:i%2===1}),dC(`${s.actual}%`,sC[3],{align:AlignmentType.CENTER,bold:true,shade:i%2===1}),
        dC(`${s.deviation>0?"+":""}${s.deviation}%p`,sC[4],{align:AlignmentType.CENTER,bold:true,color:dc,shade:i%2===1})]})
    }),
  ]}));

  return ch;
}

function buildAnnual(bms,wbs,ai){
  // 연간보고서는 분기보고서 + 연간 종합 + 추진 경과
  const ch = buildQuarter(bms,wbs,ai);

  ch.push(new Paragraph({children:[new PageBreak()]}));
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"8. 연간 추진 경과",font:"맑은 고딕"})]}));
  ch.push(new Paragraph({spacing:{after:120},children:[new TextRun({text:`${YEAR}년 사업 추진 주요 경과를 월별로 정리하였습니다.`,font:"맑은 고딕",size:20})]}));

  // 추진체계
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:"9. 사업 추진체계",font:"맑은 고딕"})]}));
  const orgCols=[2000,3000,4026];
  ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:orgCols,rows:[
    new TableRow({children:[hC("구분",orgCols[0]),hC("기관",orgCols[1]),hC("역할",orgCols[2])]}),
    new TableRow({children:[dC("시행기관",orgCols[0],{bold:true}),dC("아산시",orgCols[1]),dC("스마트도시팀, 사업총괄",orgCols[2])]}),
    new TableRow({children:[dC("직접보조사업자",orgCols[0],{bold:true,shade:true}),dC("제일엔지니어링",orgCols[1],{shade:true}),dC("PMO, 사업관리·기술지원",orgCols[2],{shade:true})]}),
    new TableRow({children:[dC("간접보조사업자",orgCols[0],{bold:true}),dC("호서대학교",orgCols[1]),dC("이노베이션센터 운영",orgCols[2])]}),
    new TableRow({children:[dC("간접보조사업자",orgCols[0],{bold:true,shade:true}),dC("충남연구원",orgCols[1],{shade:true}),dC("연구·정책지원",orgCols[2],{shade:true})]}),
    new TableRow({children:[dC("간접보조사업자",orgCols[0],{bold:true}),dC("KAIST",orgCols[1]),dC("기술자문·연구지원",orgCols[2])]}),
  ]}));

  return ch;
}

// ── Document Assembly ──
function assemble(ch,type){
  return new Document({
    numbering:{config:[
      {reference:"bullets",levels:[{level:0,format:LevelFormat.BULLET,text:"\u2022",alignment:AlignmentType.LEFT,
        style:{paragraph:{indent:{left:720,hanging:360}}}}]},
    ]},
    styles:{default:{document:{run:{font:"맑은 고딕",size:20}}},
      paragraphStyles:[
        {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
          run:{size:28,bold:true,font:"맑은 고딕",color:"1B3A5C"},paragraph:{spacing:{before:360,after:120},outlineLevel:0}},
        {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
          run:{size:24,bold:true,font:"맑은 고딕",color:"2E75B6"},paragraph:{spacing:{before:240,after:100},outlineLevel:1}},
      ]},
    sections:[{
      properties:{page:{size:{width:11906,height:16838},margin:{top:1440,bottom:1440,left:1440,right:1440}}},
      headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,
        children:[new TextRun({text:`아산시 스마트시티 | ${META.name} (${META.period})`,font:"맑은 고딕",size:16,color:"999999"})]})]})},
      footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[
        new TextRun({text:"아산시 스마트도시팀 | 제일엔지니어링 PMO | ",font:"맑은 고딕",size:16,color:"999999"}),
        new TextRun({children:[PageNumber.CURRENT],font:"맑은 고딕",size:16,color:"999999"}),
        new TextRun({text:` / `,font:"맑은 고딕",size:16,color:"999999"}),
        new TextRun({children:[PageNumber.TOTAL_PAGES],font:"맑은 고딕",size:16,color:"999999"}),
      ]})]})},
      children:ch,
    }],
  });
}

// ── AI Analysis ──
async function getAI(bms,wbs){
  const prompt = `당신은 국토교통부에 제출하는 스마트시티 사업 ${META.name} 작성 전문가입니다.
아산시 강소형 스마트시티 조성사업(240억원, 2023.12~2026.12) 데이터:
- 예산 집행률: ${bms.execRate}% (${bms.totalExec.toFixed(1)}억/${bms.totalBudget.toFixed(0)}억)
- WBS: ${wbs.overall.actualRate}% (달성률 ${wbs.overall.achieveRate}%)
- 작업: 전체${wbs.overall.total} 완료${wbs.overall.done} 지연${wbs.overall.delayed} 대기${wbs.overall.waiting}
- 기간소진율: ${TPCT}%, D-${DDAY}
단위사업: ${bms.projects.map(p=>`#${p.num}${p.name}(${p.rate}%)`).join(", ")}

JSON만 응답: {"executive_summary":"4문장","key_achievements":["3개"],"risk_analysis":[{"level":"긴급|주의","title":"","description":"","action":""}],"budget_insight":"3문장","wbs_insight":"3문장","next_week_plan":["향후계획3개"],"recommendations":["권고3개"]}`;

  console.log("\n🤖 Claude AI 분석...");
  const r = await callAI(prompt);
  if(r){try{return JSON.parse(r.replace(/```json\n?/g,"").replace(/```\n?/g,"").trim())}catch{}}
  console.log("  📝 Fallback 사용");
  const gap=(parseFloat(TPCT)-bms.execRate).toFixed(0);
  return {
    executive_summary:`${META.period} 기준 예산 집행률 ${bms.execRate}%, WBS 공정률 ${wbs.overall.actualRate}%. 기간 소진율(${TPCT}%) 대비 집행률 격차 ${gap}%p. 잔여 D-${DDAY}일.`,
    key_achievements:["BMS/WBS 자동 동기화 시스템 운영","통합 포털 실시간 현행화","GitHub Actions 30분 자동 갱신"],
    risk_analysis:[
      {level:"긴급",title:`집행률 격차 ${gap}%p`,description:`기간소진율 대비 집행률 부족`,action:"미착수 사업 발주 가속화"},
      {level:"주의",title:`WBS 지연 ${wbs.delayed.length}건`,description:"계획 대비 실적 미달",action:"지연 원인 분석 및 만회 계획"},
    ],
    budget_insight:`총 240억 중 ${bms.totalExec.toFixed(1)}억 집행(${bms.execRate}%). 잔여 ${bms.totalRemain.toFixed(1)}억을 ${(DDAY/30).toFixed(0)}개월 내 집행 필요.`,
    wbs_insight:`전체 ${wbs.overall.total}건 중 완료 ${wbs.overall.done}건, 지연 ${wbs.overall.delayed}건.`,
    next_week_plan:["미착수 사업 발주 진행","지연 작업 만회 계획 수립","월간 집행 실적 점검"],
    recommendations:["대형 미집행 사업 발주 가속화","비목간 전용 검토","준공 로드맵 재설정"],
  };
}

// ── Main ──
async function main(){
  console.log("=".repeat(60));
  console.log(`📋 ${META.name} 생성 — ${META.period} (${TODAY})`);
  console.log(`   제출처: ${META.to} | 기한: ${META.due}`);
  console.log("=".repeat(60));

  console.log("\n📦 데이터 수집...");
  const [bms_raw,ws,wd] = await Promise.all([fetchJSON(BMS_URL),fetchJSON(WBS_SUM_URL),fetchJSON(WBS_DATA_URL)]);
  const bms = processBMS(bms_raw);
  const wbs = processWBS(ws,wd);
  console.log(`  집행률: ${bms.execRate}%, WBS: ${wbs.overall.actualRate}%`);

  const ai = await getAI(bms,wbs);

  console.log("\n📄 DOCX 생성...");
  let children;
  if(TYPE==="annual") children=buildAnnual(bms,wbs,ai);
  else if(TYPE==="quarter") children=buildQuarter(bms,wbs,ai);
  else children=buildMonthly(bms,wbs,ai);

  const doc = assemble(children,TYPE);
  const buf = await Packer.toBuffer(doc);

  const dir="reports"; if(!fs.existsSync(dir))fs.mkdirSync(dir,{recursive:true});
  const typeLabel={monthly:"월별관리카드",quarter:`${Q}분기보고서`,annual:"연간성과보고서"};
  const fn = `${typeLabel[TYPE]}_${TODAY}.docx`;
  fs.writeFileSync(`${dir}/${fn}`,buf);
  console.log(`  ✅ ${dir}/${fn} (${(buf.length/1024).toFixed(0)} KB)`);

  fs.writeFileSync(`${dir}/latest_${TYPE}.json`,JSON.stringify({
    type:TYPE,generated:KST.toISOString(),period:META.period,
    bms:{execRate:bms.execRate,totalExec:bms.totalExec},wbs:wbs.overall,
    ai_engine:API_KEY?"claude-sonnet-4":"fallback",filename:fn,
  },null,2));
  console.log(`  ✅ ${dir}/latest_${TYPE}.json`);
  console.log("\n🎉 완료!");
}
main().catch(e=>{console.error("ERROR:",e);process.exit(1)});

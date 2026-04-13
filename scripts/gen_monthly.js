const fs = require("fs");
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,Header,Footer,
  AlignmentType,HeadingLevel,BorderStyle,WidthType,ShadingType,LevelFormat,
  PageBreak,PageNumber,VerticalAlign} = require("docx");

const bms = JSON.parse(fs.readFileSync("/home/claude/bms.json","utf8"));
const wsum = JSON.parse(fs.readFileSync("/home/claude/wbs_sum.json","utf8"));

const TW=9026, now=new Date(), Y=now.getFullYear(), M=now.getMonth()+1;
const br={style:BorderStyle.SINGLE,size:1,color:"000000"};
const bo={top:br,bottom:br,left:br,right:br};
const cm={top:60,bottom:60,left:100,right:100};
const hs={fill:"D9E2F3",type:ShadingType.CLEAR};
const gs={fill:"F2F2F2",type:ShadingType.CLEAR};

function hC(t,w,opt={}){return new TableCell({borders:bo,width:{size:w,type:WidthType.DXA},
  shading:hs,margins:cm,verticalAlign:VerticalAlign.CENTER,rowSpan:opt.rs,columnSpan:opt.cs,
  children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:t,bold:true,font:"맑은 고딕",size:18})]})]})}
function dC(t,w,opt={}){return new TableCell({borders:bo,width:{size:w,type:WidthType.DXA},
  margins:cm,shading:opt.shade?gs:undefined,verticalAlign:VerticalAlign.CENTER,rowSpan:opt.rs,columnSpan:opt.cs,
  children:[new Paragraph({alignment:opt.align||AlignmentType.LEFT,spacing:{before:40,after:40},
    children:[new TextRun({text:String(t),font:"맑은 고딕",size:18,bold:opt.bold,color:opt.color})]})]})}
function sP(t,opt={}){return new Paragraph({spacing:{before:opt.before||0,after:opt.after||0},
  alignment:opt.align||AlignmentType.LEFT,heading:opt.heading,
  children:[new TextRun({text:t,font:"맑은 고딕",size:opt.size||20,bold:opt.bold,color:opt.color})]})}

const s = bms.summary;
const ch = [];

// ═══ 표지 ═══
ch.push(new Paragraph({spacing:{before:3000}}));
ch.push(sP("강소형 스마트시티 조성사업",{align:AlignmentType.CENTER,size:44,bold:true}));
ch.push(sP("관리카드",{align:AlignmentType.CENTER,size:52,bold:true,before:200}));
ch.push(new Paragraph({spacing:{before:600}}));
ch.push(sP(`${Y}. ${String(M).padStart(2,"0")}.`,{align:AlignmentType.CENTER,size:32,before:400}));
ch.push(new Paragraph({spacing:{before:1200}}));
ch.push(sP("아 산 시",{align:AlignmentType.CENTER,size:40,bold:true}));
ch.push(new Paragraph({children:[new PageBreak()]}));

// ═══ 사업 기본정보 ═══
const ic=[2500,6526];
ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:ic,rows:[
  new TableRow({children:[hC("담당부서",ic[0]),dC("아산시 스마트도시과 스마트도시관리팀",ic[1])]}),
  new TableRow({children:[hC("팀장",ic[0]),dC("박상국",ic[1])]}),
  new TableRow({children:[hC("주무관",ic[0]),dC("임용훈",ic[1])]}),
  new TableRow({children:[hC("사업명",ic[0]),dC("아산시 강소형 스마트시티 조성사업",ic[1],{bold:true})]}),
  new TableRow({children:[hC("대상지",ic[0]),dC("충청남도 아산시 도고면 및 배방읍 일원 등",ic[1])]}),
  new TableRow({children:[hC("사업비",ic[0]),dC("총 240억원 (국비 120억, 지방비 120억)",ic[1])]}),
  new TableRow({children:[hC("사업기간",ic[0]),dC("2023. 08. ~ 2026. 12.31 (총 41개월)",ic[1])]}),
]}));
ch.push(sP("",{after:200}));

// 추진체계 표
const tc=[1500,2000,5526];
ch.push(sP("□ 추진체계 및 참여기관",{bold:true,before:200}));
ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:tc,rows:[
  new TableRow({children:[hC("구분",tc[0]),hC("기관·기업명",tc[1]),hC("주요역할",tc[2])]}),
  new TableRow({children:[dC("총괄",tc[0],{align:AlignmentType.CENTER,bold:true}),dC("아산시",tc[1]),dC("사업주관, 사업비 매칭",tc[2])]}),
  new TableRow({children:[dC("참여기업",tc[0],{align:AlignmentType.CENTER,bold:true}),dC("제일엔지니어링",tc[1]),dC("디지털 OASIS 서비스/인프라 구축, 비즈니스 인큐베이팅 서비스 구축",tc[2])]}),
  new TableRow({children:[dC("참여기관",tc[0],{align:AlignmentType.CENTER,bold:true,rs:3}),dC("충남연구원",tc[1]),dC("아산 스마트시티 거버넌스 관리 및 정책지원",tc[2])]}),
  new TableRow({children:[dC("호서대학교",tc[1]),dC("이노베이션 센터 창업교육지원, 운영매뉴얼 개발, 리빙랩 구축 운영, 순찰로봇 개발",tc[2])]}),
  new TableRow({children:[dC("한국과학기술원(KAIST)",tc[1]),dC("시설물 위치 기반 관리체계 마련 및 표준화",tc[2])]}),
]}));
ch.push(new Paragraph({children:[new PageBreak()]}));

// ═══ Ⅰ 사업 주요내용 ═══
ch.push(sP("Ⅰ 사업 주요내용",{heading:HeadingLevel.HEADING_1}));
ch.push(sP("□ 추진경과",{bold:true,before:200}));
const milestones = [
  ["2023. 12.","지자체-참여기관 협약 체결"],
  ["2024. 06.","국토부 실시계획 승인"],
  ["2024. 09.","서비스 인프라 구축 입찰 공고"],
  ["2024. 11.","이노베이션센터 착수보고 및 인테리어 공사"],
  ["2025. 01.","이노베이션 센터 실시설계 및 시공 용역 준공"],
  ["2025. 03.","서비스 인프라 구축 용역 계약"],
  ["2025. 04.","서비스 인프라 구축 용역 계약 해지(하도급 위반)"],
  ["2025. 08.","유무선 네트워크 구축 용역 입찰 공고"],
  ["2025. 10.","유무선 네트워크 구축 용역 계약 체결"],
  ["2025. 12.","사업기간 연장 국토부 실시계획 변경 승인, 서비스 인프라 구축 용역 계약"],
  ["2026. 01.","서비스 인프라 구축 착수 보고, OASIS SPOT 도시관리계획 변경 용역 공고/계약"],
];
for(const [d,c] of milestones){
  ch.push(new Paragraph({spacing:{before:20},indent:{left:360},
    children:[new TextRun({text:`ㅇ ${d}`,font:"맑은 고딕",size:18,bold:true}),new TextRun({text:`  • ${c}`,font:"맑은 고딕",size:18})]}));
}

// ═══ 사업별 세부내용 ═══
ch.push(sP("",{after:100}));
ch.push(sP("□ 사업별 세부내용",{bold:true,before:200}));
const dc=[1500,1500,3026,1500,1005];
const projs = [
  ["사업관리","사업관리","사업관리, 추진과제 세부 설계, PMO","제일ENG","1,836.5"],
  ["스마트 인프라\n및 서비스\n구축","디지털 OASIS\n사용자 서비스","OASIS SPOT 구성, 전자시민증, DRT, 스마트폴, 무인매장, 공공WiFi, 순찰로봇","제일ENG\n호서대","7,935"],
  ["","디지털 OASIS\n운영 서비스","AI시티 플랫폼 구축 및 연계","제일ENG","3,600"],
  ["","디지털 OASIS\n정보관리 서비스","정보관리시스템 구축, 시설물 위치기반 표준 서비스","제일ENG\nKAIST","2,700"],
  ["","비즈니스\n인큐베이팅 서비스","이노베이션센터/관제, AI융복합, SDDC, 메타버스, ESG교육","제일ENG","5,950"],
  ["도시산업\n육성","이노베이션센터","리빙랩, 창업교육, 운영매뉴얼","호서대","700"],
  ["","거버넌스","협의체 구성·운영, 정책개발","충남연구원","200"],
  ["사업성과\n확산","성과홍보","KPI 도출, 홍보영상, 전시참여","제일ENG","1,078.5"],
];
ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:dc,rows:[
  new TableRow({children:[hC("구분",dc[0]),hC("세부과제",dc[1]),hC("세부내용",dc[2]),hC("추진주체",dc[3]),hC("사업비\n(백만원)",dc[4])]}),
  ...projs.map((p,i)=>new TableRow({children:[
    dC(p[0],dc[0],{align:AlignmentType.CENTER,shade:i%2===1,bold:!!p[0]}),
    dC(p[1],dc[1],{align:AlignmentType.CENTER,shade:i%2===1}),
    dC(p[2],dc[2],{shade:i%2===1}),
    dC(p[3],dc[3],{align:AlignmentType.CENTER,shade:i%2===1}),
    dC(p[4],dc[4],{align:AlignmentType.RIGHT,shade:i%2===1}),
  ]}))
]}));
ch.push(new Paragraph({children:[new PageBreak()]}));

// ═══ 공정관리 문제점 조치계획 ═══
ch.push(sP("□ 공정관리 문제점 조치계획",{bold:true,before:200}));
const pcC=[4513,4513];
ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:pcC,rows:[
  new TableRow({children:[hC("공정관리 문제점",pcC[0]),hC("조치계획",pcC[1])]}),
  new TableRow({children:[
    dC("ㅇ 디지털 OASIS SPOT 대상지 미확정에 따른 후행 사업 지연\n- 디지털 OASIS 생태계는 다수의 서비스가 SPOT과 유기적으로 연계되는 통합형 구조\n- 대상지 확정 및 국토부 실시계획 변경 승인이 선행되어야 후속 사업 착수 가능",pcC[0]),
    dC("ㅇ 절차 병행 추진으로 지연 만회\n- 국토부 사업기간 연장 승인과 대상지 변경 승인 절차 병행\n- 현장 무관 구축 분야 우선 추진\n- SPOT 확정 즉시 공고할 수 있도록 사전규격·검토·내부결재 선행\n- 기곡리 296-4, 300-1 부지 용도 폐지 용역 진행 중",pcC[1]),
  ]}),
]}));

// ═══ 당월실적 및 익월계획 ═══
ch.push(sP("",{after:200}));
ch.push(sP("□ 당월실적 및 익월계획",{bold:true,before:300}));
const mc=[1800,3613,3613];
const monthData = [
  {cat:"1. 사업관리",items:[
    {sub:"추진 실적",cur:`ㅇ 서비스 인프라 구축 용역 상세 설계 진행\nㅇ 유무선 네트워크 구축 사업 장비 입고/구축\nㅇ 잔여 단위사업 발주준비\nㅇ 예산 및 공정률 시스템화`,next:`ㅇ 서비스 인프라 구축 용역 상세 설계 진행\nㅇ 유무선 네트워크 이노베이션센터 先 구축\nㅇ 잔여 단위사업 발주`},
    {sub:"사업비 관리",cur:"ㅇ 인건비 및 일반수용비 등 집행\nㅇ 보조사업 회계 정산 협의",next:"ㅇ 인건비 및 일반수용비 등 집행\nㅇ 국토부 보조금 집행점검 대응"},
  ]},
  {cat:"2. 스마트 인프라 및 서비스 구축",items:[
    {sub:"OASIS SPOT",cur:"ㅇ 대상지 기존 용도(주차장, 수도시설) 폐지 절차 진행\nㅇ 입찰공고 준비(공고문, 제안요청서 등)",next:"ㅇ 용도 폐지 절차 진행\nㅇ 국토부 승인 후 입찰공고"},
    {sub:"서비스 인프라\n구축 용역",cur:"ㅇ 착수 보고 및 상세 설계\nㅇ 하도급 승인 절차 진행",next:"ㅇ 상세설계 진행\nㅇ HW 인프라 장비실 구축 착수"},
    {sub:"유무선 네트워크",cur:"ㅇ 이노베이션 센터 현장 답사 및 구축 진행",next:"ㅇ 이노베이션 센터 공공WiFi 및 네트워크 구축\nㅇ 구축 물량 입고 및 검수"},
  ]},
  {cat:"3. 도시산업 육성",items:[
    {sub:"거버넌스 운영",cur:"ㅇ 거버넌스 운영/관리 계획 수립\nㅇ 스마트시티 우수사례 조사",next:"ㅇ 3차 거버넌스 실무자문단 추진\nㅇ 아산시 공공기관 업무협의"},
    {sub:"이노베이션센터\n운영 관리",cur:"ㅇ 사업정산 및 종합보고\nㅇ 이노베이션센터 운영 관리",next:"ㅇ 리빙랩 아카이빙 홍보\nㅇ 창업 입주기업 모집"},
  ]},
];
const mRows = [new TableRow({children:[hC("구분",mc[0]),hC(`당월 추진실적\n(${Y}.${String(M).padStart(2,"0")}.01~${String(M).padStart(2,"0")}.${new Date(Y,M,0).getDate()})`,mc[1]),hC(`익월 추진계획\n(${Y}.${String(M===12?1:M+1).padStart(2,"0")}.01~${String(M===12?1:M+1).padStart(2,"0")}.${new Date(Y,M===12?M+1:M+1,0).getDate()})`,mc[2])]})];
for(const g of monthData){
  mRows.push(new TableRow({children:[
    new TableCell({borders:bo,width:{size:mc[0],type:WidthType.DXA},margins:cm,shading:hs,columnSpan:3,
      children:[new Paragraph({children:[new TextRun({text:g.cat,bold:true,font:"맑은 고딕",size:18})]})]}),
  ]}));
  for(const it of g.items){
    mRows.push(new TableRow({children:[dC(it.sub,mc[0],{align:AlignmentType.CENTER,bold:true}),dC(it.cur,mc[1]),dC(it.next,mc[2])]}));
  }
}
ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:mc,rows:mRows}));
ch.push(new Paragraph({children:[new PageBreak()]}));

// ═══ Ⅲ 사업 성과관리 ═══
ch.push(sP("Ⅲ 사업 성과관리",{heading:HeadingLevel.HEADING_1}));
ch.push(sP("□ 성과목표 (실시계획 구축단계 KPI)",{bold:true,before:200}));
const kc=[2000,3513,3513];
const kpis = [
  ["디지털 OASIS\nSPOT 구축","이동형 체류 공간 30기 구축","국토부 실시계획 변경 승인 완료 후 입찰공고 예정\n부지 용도 폐지 용역 진행 중"],
  ["전자 시민증","모바일 전자시민증 발급 1,000건/년 이상","상세 기능정의 및 설계 진행 예정"],
  ["수요응답형\n모빌리티(DRT)","DRT 이용자 비율 30% 미만","수요응답형 버스 운송사업자 모집 예정\n버스 2대, 스마트 정거장 2개소 구축 예정"],
  ["SDDC 기반\nIT인프라","시스템 다운타임 99.99% 가용성","상세 설계 진행 중\nHW 인프라 장비실 구축 착수"],
  ["무인 순찰 로봇","도시 공공데이터 수집","구축 완료, 시범운영 중"],
];
ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:kc,rows:[
  new TableRow({children:[hC("분야",kc[0]),hC("성과목표",kc[1]),hC("달성현황 (예정계획)",kc[2])]}),
  ...kpis.map((k,i)=>new TableRow({children:[
    dC(k[0],kc[0],{align:AlignmentType.CENTER,bold:true,shade:i%2===1}),
    dC(k[1],kc[1],{shade:i%2===1}),dC(k[2],kc[2],{shade:i%2===1}),
  ]})),
]}));
ch.push(new Paragraph({children:[new PageBreak()]}));

// ═══ Ⅳ 사업비 관리 ═══
ch.push(sP("Ⅳ 사업비 관리",{heading:HeadingLevel.HEADING_1}));
ch.push(sP("□ 사업비 확보 현황",{bold:true,before:200}));
ch.push(sP("(단위 : 백만원)",{align:AlignmentType.RIGHT,size:16}));
const fc=[1500,1500,1500,1500,1013,1013,1000];
ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:fc,rows:[
  new TableRow({children:[hC("구분",fc[0]),hC("총계",fc[1]),hC("2023년",fc[2]),hC("2024년",fc[3]),hC("2025년\n1차",fc[4]),hC("2025년\n2차",fc[5]),hC("잔액",fc[6])]}),
  new TableRow({children:[dC("국비",fc[0],{bold:true}),dC("12,000",fc[1],{align:AlignmentType.RIGHT}),dC("6,000",fc[2],{align:AlignmentType.RIGHT}),dC("3,000",fc[3],{align:AlignmentType.RIGHT}),dC("2,100",fc[4],{align:AlignmentType.RIGHT}),dC("900",fc[5],{align:AlignmentType.RIGHT}),dC("0",fc[6],{align:AlignmentType.RIGHT})]}),
  new TableRow({children:[dC("지방비",fc[0],{bold:true,shade:true}),dC("8,100",fc[1],{align:AlignmentType.RIGHT,shade:true}),dC("750",fc[2],{align:AlignmentType.RIGHT,shade:true}),dC("2,800",fc[3],{align:AlignmentType.RIGHT,shade:true}),dC("4,550",fc[4],{align:AlignmentType.RIGHT,shade:true}),dC("",fc[5],{shade:true}),dC("",fc[6],{shade:true})]}),
]}));

// 기관별 집행 현황
ch.push(sP("□ 기관별 집행 현황",{bold:true,before:300}));
ch.push(sP("(단위 : 백만원)",{align:AlignmentType.RIGHT,size:16}));
const ec=[2500,1500,1800,1200,2026];
const orgs = [
  ["제일엔지니어링","22,700",`${(s["총집행액"]/1e6-153-167-833).toFixed(0)}`],
  ["충남연구원","200","153"],
  ["한국과학기술원","200","167"],
  ["호서대학교 산학협력단","900","833"],
];
ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:ec,rows:[
  new TableRow({children:[hC("보조사업자",ec[0]),hC("총 사업비",ec[1]),hC("집행 누계",ec[2]),hC("집행율",ec[3]),hC("향후 집행계획",ec[4])]}),
  ...orgs.map((o,i)=>{const rate=((parseFloat(o[2])/parseFloat(o[1].replace(/,/g,"")))*100).toFixed(1);
    return new TableRow({children:[dC(o[0],ec[0],{bold:true,shade:i%2===1}),dC(o[1],ec[1],{align:AlignmentType.RIGHT,shade:i%2===1}),
      dC(o[2],ec[2],{align:AlignmentType.RIGHT,shade:i%2===1}),dC(`${rate}%`,ec[3],{align:AlignmentType.CENTER,shade:i%2===1}),
      dC("인건비, 용역비 등 집행",ec[4],{shade:i%2===1})]})
  }),
  new TableRow({children:[dC("합계",ec[0],{bold:true}),dC("24,000",ec[1],{align:AlignmentType.RIGHT,bold:true}),
    dC(`${(s["총집행액"]/1e6).toFixed(0)}`,ec[2],{align:AlignmentType.RIGHT,bold:true}),
    dC(`${s["전체집행률"].toFixed(1)}%`,ec[3],{align:AlignmentType.CENTER,bold:true}),dC("",ec[4])]}),
]}));

// 세부과제별 집행현황
ch.push(sP("□ 세부과제별 집행현황",{bold:true,before:300}));
ch.push(sP("(단위 : 백만원)",{align:AlignmentType.RIGHT,size:16}));
const sc=[2800,1200,1200,1000,1000,1826];
const subs = [
  ["디지털 OASIS SPOT (제일ENG)","3,500","15","0.4%"],
  ["전자시민증 (제일ENG)","1,200","546","45.5%"],
  ["수요응답형 모빌리티 (제일ENG)","1,000","0","0.0%"],
  ["순찰로봇 (호서대)","200","104","52.0%"],
  ["스마트폴 (제일ENG)","600","0","0.0%"],
  ["무인매장 (제일ENG)","700","0","0.0%"],
  ["스마트 공공 WiFi (제일ENG)","735","269","36.6%"],
  ["디지털 노마드 접수/운영 (제일ENG)","2,000","820","41.0%"],
  ["AI 시티관제 플랫폼 (제일ENG)","1,600","929","58.1%"],
  ["디지털 OASIS 정보관리 (제일ENG)","2,500","1,093","43.7%"],
  ["KAIST 표준화 (KAIST)","200","75","37.5%"],
  ["시설물 위치기반 서비스 (제일ENG)","200","0","0.0%"],
  ["이노베이션 센터 구축 (제일ENG)","1,300","1,186","91.2%"],
  ["AI 융복합 서비스 (제일ENG)","750","437","58.3%"],
  ["SDDC 플랫폼 & 네트워크 (제일ENG)","2,700","709","26.3%"],
  ["네트워크 구축 (제일ENG)","400","146","36.5%"],
  ["메타버스 (제일ENG)","600","0","0.0%"],
  ["ESG 교육 (제일ENG)","200","0","0.0%"],
  ["이노베이션센터 운영 (호서대)","700","83","11.9%"],
  ["거버넌스 운영 (충남연구원)","200","27","13.5%"],
  ["사업관리 및 설계 (제일ENG)","2,715","1,316","48.5%"],
];
ch.push(new Table({width:{size:TW,type:WidthType.DXA},columnWidths:sc,rows:[
  new TableRow({children:[hC("세부과제 (수행기관)",sc[0]),hC("총계",sc[1]),hC("집행 누계",sc[2]),hC("집행율",sc[3]),hC("이월액",sc[4]),hC("비고",sc[5])]}),
  ...subs.map((r,i)=>new TableRow({children:[
    dC(r[0],sc[0],{shade:i%2===1}),dC(r[1],sc[1],{align:AlignmentType.RIGHT,shade:i%2===1}),
    dC(r[2],sc[2],{align:AlignmentType.RIGHT,shade:i%2===1}),
    dC(r[3],sc[3],{align:AlignmentType.CENTER,bold:true,shade:i%2===1,color:parseFloat(r[3])>=50?"00B050":(parseFloat(r[3])>0?"ED7D31":"C00000")}),
    dC("",sc[4],{shade:i%2===1}),dC("",sc[5],{shade:i%2===1}),
  ]})),
]}));

const doc = new Document({
  numbering:{config:[{reference:"bullets",levels:[{level:0,format:LevelFormat.BULLET,text:"\u2022",alignment:AlignmentType.LEFT,
    style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
  styles:{default:{document:{run:{font:"맑은 고딕",size:20}}},
    paragraphStyles:[
      {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:32,bold:true,font:"맑은 고딕"},paragraph:{spacing:{before:360,after:200},outlineLevel:0}},
      {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:26,bold:true,font:"맑은 고딕"},paragraph:{spacing:{before:240,after:120},outlineLevel:1}},
    ]},
  sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1440,right:1440}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,
      children:[new TextRun({text:`아산시 강소형 스마트시티 | 관리카드 ${Y}.${String(M).padStart(2,"0")}`,font:"맑은 고딕",size:16,color:"999999"})]})]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[
      new TextRun({text:"- ",font:"맑은 고딕",size:18}),
      new TextRun({children:[PageNumber.CURRENT],font:"맑은 고딕",size:18}),
      new TextRun({text:" -",font:"맑은 고딕",size:18})]})]}),},
    children:ch}]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync("/mnt/user-data/outputs/월별관리카드_양식.docx",buf);
  console.log(`✅ 월별관리카드 생성 완료: ${(buf.length/1024).toFixed(0)} KB`);
});

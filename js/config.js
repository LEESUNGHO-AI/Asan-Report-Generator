/* ══════════════════════════════════════════════════════════════════════════
 *  config.js — 사업정보 · 공문서 기본설정
 *  아산시 강소형 스마트시티 조성사업 | 공문서 생성 시스템 v5.0
 *  ※ 하드코딩 최소화: 담당자·문서번호·수신처·결재선은 UI에서 수정 가능
 * ══════════════════════════════════════════════════════════════════════════ */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.GovConfig = factory();
})(typeof self !== "undefined" ? self : this, function () {

  /* ── 1. 사업 기본정보 (보조금 교부결정 기준) ────────────────────────── */
  const PROJECT = {
    name: "아산시 강소형 스마트시티 조성사업",
    brand: "디지털 OASIS",
    type: "강소형 스마트시티 조성사업(기존도시형)",
    ministry: "국토교통부",
    intermediary: "한국스마트도시협회",
    owner: "아산시",
    ownerDept: "아산시 도시계획과 스마트도시팀",
    ownerManager: "이현경 팀장",          // 2026. 7. 인사이동 반영 (前 박상국 팀장)
    location: "충청남도 아산시 도고면·배방읍 일원",
    periodFrom: "2023-12-01",
    periodTo: "2026-12-31",              // ※ 협약서 기준 사업기간 — 설정 패널에서 변경 가능
    totalBudget: 24000000000,            // 원
    fund: [
      { name: "국비", rate: 50, amount: 12000000000, ministry: "국토교통부" },
      { name: "시비", rate: 38, amount: 9120000000, ministry: "아산시" },
      { name: "도비", rate: 12, amount: 2880000000, ministry: "충청남도" },
    ],
    consortium: [
      { org: "㈜제일엔지니어링종합건축사사무소", role: "직접보조사업자(수행기관·총괄 PMO)", scope: "사업관리(PMO)·발주지원·인프라 구축·정산" },
      { org: "호서대학교 산학협력단", role: "간접보조사업자", scope: "아산 이노베이션 스퀘어 운영·실증" },
      { org: "충남연구원", role: "간접보조사업자", scope: "정책연구·성과지표 설계·리빙랩" },
      { org: "한국과학기술원(KAIST)", role: "간접보조사업자", scope: "AI·데이터 기술자문" },
    ],
    legalBasis: [
      "「스마트도시 조성 및 산업진흥 등에 관한 법률」 제12조",
      "「보조금 관리에 관한 법률」 제27조 및 같은 법 시행령 제12조",
      "국토교통부-아산시 강소형 스마트시티 조성사업 협약서",
    ],
  };

  /* ── 2. 공문서 기본설정 (행정업무규정 별지 제1호서식 결문 요소) ─────── */
  const DOCSET = {
    senderOrg: "㈜제일엔지니어링종합건축사사무소",
    senderName: "대표이사",                     // 발신명의
    docNoPrefix: "제일엔지니어링스마트시티",     // 문서번호 앞부분(처리과 기관코드)
    docNoSeq: 1,                                // 일련번호(문서등록대장 연동)
    recipient: "아산시장",
    recipientVia: "",                           // (경유) — 해당 없으면 공란
    handlerDept: "스마트시티사업본부 PMO팀",
    drafter: { title: "차장", name: "강문석" },   // 기안자
    reviewer: { title: "상무", name: "이성호" },  // 검토자 ※ '이사' 아님
    approver: { title: "대표이사", name: "김재현" }, // 결재권자
    cooperator: "",
    zip: "31460",
    address: "충청남도 아산시 도고면 기곡리 174-1, 아산 이노베이션 스퀘어",
    homepage: "https://leesungho-ai.github.io/Asan-Smartcity-integration-Portal/",
    tel: "041-538-1234",
    fax: "041-538-1235",
    email: "asan.smartcity@jeil-eng.co.kr",
    openLevel: "부분공개(제9조제1항제7호)",       // 공개 / 부분공개 / 비공개
    keepYears: "준영구",                          // 보존기간(정산종료 2031년 이후)
  };

  /* ── 3. 문서 유형 정의 ──────────────────────────────────────────────── */
  const DOCTYPES = {
    official: {
      key: "official", kind: "시행문",
      title: "사업추진현황 통보(시행문)",
      file: "시행문_사업추진현황통보",
      desc: "행정업무규정 별지 제1호서식 — 두문·본문·결문·붙임·끝 표기 완비",
      target: "아산시장(스마트도시팀)", cycle: "수시",
      style: "1.가.1)", badge: "별지 제1호서식",
    },
    weekly: {
      key: "weekly", kind: "보고자료",
      title: "주간 업무추진 보고",
      file: "주간업무추진보고",
      desc: "PMO 내부 주간 진도관리 — 금주 실적·차주 계획·조치필요사항",
      target: "PMO 내부", cycle: "매주 금요일",
      style: "□○-", badge: "개조식 보고자료",
    },
    monthly: {
      key: "monthly", kind: "실적보고서",
      title: "보조사업 추진실적 보고(월간)",
      file: "월간_보조사업추진실적보고",
      desc: "보조금법 제27조 — 교부·집행·잔액·공정·인력·중요재산 종합",
      target: "아산시 / 충청남도", cycle: "익월 5일",
      style: "□○-", badge: "보조금법 제27조",
    },
    quarterly: {
      key: "quarterly", kind: "실적보고서",
      title: "분기 추진실적 보고",
      file: "분기_추진실적보고",
      desc: "국토교통부·한국스마트도시협회 제출 — 성과지표·리스크 포함",
      target: "국토교통부 / 한국스마트도시협회", cycle: "분기말 +15일",
      style: "□○-", badge: "국토부 제출용",
    },
    annual: {
      key: "annual", kind: "실적보고서",
      title: "연간 사업추진 실적보고서",
      file: "연간_사업추진실적보고서",
      desc: "연차 결산 — 연도별 교부·집행·이월·반납, 중요재산 명세 포함",
      target: "국토교통부 / 아산시", cycle: "익년 1월 31일",
      style: "□○-", badge: "정산 연계",
    },
    brief: {
      key: "brief", kind: "보고자료",
      title: "핵심 추진현황 보고(2매)",
      file: "핵심추진현황보고_2매",
      desc: "시장·부시장 대면보고용 2매 요약 — 현황·쟁점·건의사항",
      target: "아산시장 / 부시장", cycle: "수시",
      style: "□○-", badge: "간부보고 2매",
    },
  };

  /* ── 4. 단위사업 / 비목 매핑 ────────────────────────────────────────── */
  const BMS_UNIT_MAP = {
    "스마트 공공 WIFI": 1, "아산시 강소형 스마트시티 네트워크 구축": 1, "네트워크 구축": 1,
    "모바일 전자시민증 플랫폼 / 인프라": 2,
    "이노베이션센터 구축": 3, "이노베이션 센터/ 관제 시스템 구축": 3,
    "디지털 OASIS SPOT": 4, "무인매장": 4,
    "SDDC Platform 구축": 5,
    "AI통합관제 및 운영 플랫폼 / 인프라": 6,
    "디지털OASIS 정보관리 시스템": 7,
    "수요응답형 DRT 서비스 운영 플랫폼 구축": 8, "수요응답형 DRT 서비스 운영 HW 구축": 8,
    "정보통신감리": 9,
    "스마트폴&디스플레이": 10,
    "메타버스 플랫폼": 11,
    "디지털 노마드접수/운영 및 거래관리": 12,
    "데이터기반 AI 융복합 서비스 구축": 13,
    "국제표준 디지털링크 공유 플랫폼": 14,
    "시설물 위치기반 표준 서비스 플랫폼": 14,
  };
  const UNIT_NAMES = {
    1: "유무선 네트워크 구축", 2: "모바일 전자시민증(ECC)", 3: "이노베이션 스퀘어 구축",
    4: "디지털 OASIS SPOT·무인매장", 5: "SDDC Platform 구축", 6: "AI 통합관제 플랫폼",
    7: "디지털 OASIS 정보관리", 8: "DRT 수요응답형 교통", 9: "정보통신 감리용역",
    10: "스마트폴 & 디스플레이", 11: "메타버스 플랫폼", 12: "디지털 노마드(NOP)",
    13: "AI 융복합 서비스", 14: "디지털링크 표준 플랫폼",
  };
  const BIMOK_CLEAN = {
    "인건비(110)": "인건비", "운영비(210)": "운영비", "여비(220)": "여비",
    "연구개발비(260)": "연구개발비", "사업비배분(320)": "사업비 배분", "사업비 배분(320)": "사업비 배분",
    "유형자산(430)": "유형자산", "무형자산(440)": "무형자산(소프트웨어)", "건설비(420)": "건설비", "기타": "기타",
  };
  const BIMOK_ORDER = ["건설비", "무형자산(소프트웨어)", "유형자산", "인건비", "연구개발비", "운영비", "여비", "사업비 배분", "기타"];

  /* ── 5. 데이터 소스 (Slack → Notion → GitHub Pages) ─────────────────── */
  const SRC = {
    bms: "https://leesungho-ai.github.io/Asan-Smart-City-Budget-Management-System-BMS-/data/budget.json",
    wbsSum: "https://leesungho-ai.github.io/Asan-Smartcity-WBS/data/summary-data.json",
    wbsDat: "https://leesungho-ai.github.io/Asan-Smartcity-WBS/data/wbs-data.json",
    asset: "https://leesungho-ai.github.io/Asan-asset-management/data/assets.json",
    hr: "https://leesungho-ai.github.io/Asan-HR-Management-Portal/data/hr.json",   // 1순위(JSON)
    hrFallback: "https://leesungho-ai.github.io/Asan-HR-Management-Portal/index.html", // 2순위(HTML 파싱)
  };

  return { PROJECT, DOCSET, DOCTYPES, BMS_UNIT_MAP, UNIT_NAMES, BIMOK_CLEAN, BIMOK_ORDER, SRC };
});

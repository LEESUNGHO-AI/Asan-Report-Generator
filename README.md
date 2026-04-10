# 아산시 강소형 스마트시티 — 보고서 자동 생성 시스템

## 개요
BMS(예산) + WBS(공정률) 데이터를 자동 수집하여 주간/월간 진도 보고서를 DOCX로 생성합니다.

## 데이터 소스
| 소스 | URL | 내용 |
|------|-----|------|
| BMS | `budget.json` | 비목별/단위사업별 예산 집행 현황 |
| WBS | `summary-data.json` + `wbs-data.json` | 공정률, 지연 작업, Level-1 가중평균 |

## 보고서 유형
- **주간 진도 보고서**: 매주 금요일 17:00 자동 생성
- 월간 집행 현황 리포트 (예정)
- 상위기관 보고서 (예정)

## 수동 실행
GitHub Actions → "주간 진도 보고서 자동 생성" → Run workflow

## 파일 구조
```
├── scripts/
│   └── generate_weekly_report.js   # 주간 보고서 생성 스크립트
├── reports/
│   ├── 주간진도보고서_YYYY-MM-DD_Wxx.docx
│   └── latest.json                 # 최신 보고서 메타데이터
└── .github/workflows/
    └── weekly-report.yml           # GitHub Actions 워크플로우
```

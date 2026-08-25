/* ══════════════════════════════════════════════════════════════════════════
 *  doc-report.js — 보조사업 추진실적 보고서 빌더 (주간/월간/분기/연간)
 *  개조식(□·○·-) 보고자료 체계 + 공문서 규격(A4/여백/서체/금액/날짜)
 * ══════════════════════════════════════════════════════════════════════════ */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.DocReport = factory();
})(typeof self !== "undefined" ? self : this, function () {

  function build(docx, type, P) {
    const { Document, Paragraph, TableRow, AlignmentType } = docx;
    const { bms, wbs, asset, hr, an, ctx, CFG, DS, G } = P;
    const g = G;                                   // Gongmun instance
    const { TW, C, SZ, run, EMPTY, BREAK, sq, ci, dash, note, H1, H2, hc, dc, table,
      tblCaption, tblNote, box, comma, money, moneyKr, eokStr, pct } = g;
    const PJ = CFG.PROJECT;
    const T = CFG.DOCTYPES[type];
    const A = AlignmentType;
    const f1 = (v) => Number(v || 0).toFixed(1);
    const f2 = (v) => Number(v || 0).toFixed(2);
    const isAnnual = type === "annual", isWeekly = type === "weekly";

    const ch = [];
    // 장 번호 자동 채번 — 주간보고는 Ⅰ장(사업개요) 생략에 따라 번호 자동 시프트
    const RN = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ"];
    let _sn = 0;
    const CH = (t) => H1(`${RN[_sn++]}. ${t}`);
    const chapters = (isWeekly ? [] : ["사업 개요"]).concat(
      ["추진현황 총괄", "예산 집행현황", "공정(WBS) 추진현황", "참여인력 운영현황",
       "재산 취득·관리 현황", "주요 이슈 및 조치계획", "향후 추진계획", "종합의견"]);

    /* ══════════ 표지 ══════════ */
    if (isWeekly) {
      // 주간보고는 내부 진도관리용 — 표지·목차·사업개요 생략, 약식 표제부 사용
      ch.push(new Paragraph({
        alignment: A.RIGHT, spacing: { before: 0, after: 40, line: 250 },
        children: [run(`문서번호 ${DS.docNo}  /  ${DS.openLevel}`, { size: SZ.tiny, color: C.gray })],
      }));
      ch.push(new Paragraph({
        alignment: A.CENTER, spacing: { before: 40, after: 50, line: 300 },
        children: [run(`${PJ.name} ${T.title}`, { size: 32, bold: true, color: C.navy })],
      }));
      ch.push(new Paragraph({
        alignment: A.CENTER, spacing: { before: 0, after: 180, line: 260 },
        border: { bottom: { style: docx.BorderStyle.SINGLE, size: 12, color: C.navy } },
        children: [run(`보고기간 ${ctx.period}`, { size: SZ.tiny, color: C.gray })],
      }));
      ch.push(new Paragraph({
        alignment: A.CENTER, spacing: { before: 0, after: 180, line: 260 },
        border: { bottom: { style: docx.BorderStyle.SINGLE, size: 12, color: C.navy } },
        children: [run(`${DS.senderOrg} ${DS.handlerDept}  ·  작성 ${ctx.dateKr}  ·  검토 ${DS.reviewer.title} ${DS.reviewer.name}`, { size: SZ.tiny, color: C.gray })],
      }));
    } else {
    ch.push(new Paragraph({
      alignment: A.RIGHT, spacing: { before: 0, after: 60, line: 260 },
      children: [run(`문서번호 : ${DS.docNo}`, { size: SZ.tiny, color: C.gray })],
    }));
    ch.push(new Paragraph({
      alignment: A.RIGHT, spacing: { before: 0, after: 160, line: 260 },
      children: [run(`공개구분 : ${DS.openLevel}`, { size: SZ.tiny, color: C.gray })],
    }));
    ch.push(g.approvalBox(DS));

    ch.push(EMPTY(1200));
    ch.push(new Paragraph({
      alignment: A.CENTER, spacing: { before: 900, after: 80, line: 300 },
      children: [run(PJ.name, { size: 32, bold: true, color: C.navy })],
    }));
    ch.push(new Paragraph({
      alignment: A.CENTER, spacing: { before: 0, after: 400, line: 300 },
      children: [run(`( ${PJ.brand} )`, { size: SZ.sub, color: C.gray })],
    }));
    ch.push(new Paragraph({
      alignment: A.CENTER, spacing: { before: 200, after: 200, line: 360 },
      border: { top: { style: docx.BorderStyle.SINGLE, size: 16, color: C.navy }, bottom: { style: docx.BorderStyle.SINGLE, size: 16, color: C.navy } },
      children: [run(T.title, { size: 44, bold: true, color: C.navy })],
    }));
    ch.push(new Paragraph({
      alignment: A.CENTER, spacing: { before: 240, after: 1000, line: 300 },
      children: [run(`보고대상기간 : ${ctx.period}`, { size: SZ.body, color: C.gray })],
    }));

    const cvw = [2400, TW - 2400];
    ch.push(table(cvw, [
      new TableRow({ children: [hc("보 고 유 형", cvw[0]), dc(`${T.kind} — ${T.badge}`, cvw[1], { align: A.LEFT })] }),
      new TableRow({ children: [hc("제 출 처", cvw[0]), dc(T.target, cvw[1], { align: A.LEFT })] }),
      new TableRow({ children: [hc("사 업 기 간", cvw[0]), dc(`${g.dateKr(PJ.periodFrom)} ~ ${g.dateKr(PJ.periodTo)}`, cvw[1], { align: A.LEFT })] }),
      new TableRow({ children: [hc("총 사 업 비", cvw[0]), dc(moneyKr(PJ.totalBudget), cvw[1], { align: A.LEFT })] }),
      new TableRow({ children: [hc("수 행 기 관", cvw[0]), dc(`${PJ.consortium[0].org} 외 ${PJ.consortium.length - 1}개 기관`, cvw[1], { align: A.LEFT })] }),
      new TableRow({ children: [hc("작 성 일", cvw[0]), dc(ctx.dateKr, cvw[1], { align: A.LEFT })] }),
    ]));

    ch.push(EMPTY(600));
    ch.push(new Paragraph({
      alignment: A.CENTER, spacing: { before: 700, after: 40, line: 300 },
      children: [run(DS.senderOrg, { size: SZ.sub, bold: true })],
    }));
    ch.push(new Paragraph({
      alignment: A.CENTER, spacing: { before: 0, after: 0, line: 300 },
      children: [run(`${DS.handlerDept}  (직접보조사업자 · 총괄 PMO)`, { size: SZ.small, color: C.gray })],
    }));
    ch.push(BREAK());

    /* ══════════ 목 차 ══════════ */
    const toc = chapters.map((t, i) => [RN[i], t]).concat([["붙임", "세부 현황자료"]]);
    ch.push(new Paragraph({
      alignment: A.CENTER, spacing: { before: 400, after: 400, line: 300 },
      children: [run("목      차", { size: 34, bold: true, color: C.navy })],
    }));
    const tw2 = [1200, TW - 1200];
    ch.push(table(tw2, toc.map((r) => new TableRow({
      children: [dc(r[0], tw2[0], { bold: true, color: C.navy }), dc(r[1], tw2[1], { align: A.LEFT })],
    }))));
    ch.push(BREAK());

    /* ══════════ Ⅰ. 사업 개요 ══════════ */
    ch.push(CH("사업 개요"));
    ch.push(H2("1. 추진근거"));
    PJ.legalBasis.forEach((b) => ci(b) && ch.push(ci(b)));

    ch.push(H2("2. 사업 기본현황"));
    const bw = [1900, 2900, 1900, TW - 6700];
    ch.push(table(bw, [
      new TableRow({ children: [hc("사 업 명", bw[0]), dc(PJ.name, bw[1], { align: A.LEFT }), hc("브 랜 드", bw[2]), dc(PJ.brand, bw[3])] }),
      new TableRow({ children: [hc("사 업 유 형", bw[0]), dc(PJ.type, bw[1], { align: A.LEFT }), hc("주 관 부 처", bw[2]), dc(PJ.ministry, bw[3])] }),
      new TableRow({ children: [hc("사 업 위 치", bw[0]), dc(PJ.location, bw[1], { align: A.LEFT }), hc("전 담 기 관", bw[2]), dc(PJ.intermediary, bw[3])] }),
      new TableRow({ children: [hc("사 업 기 간", bw[0]), dc(`${g.dateKr(PJ.periodFrom)} ~ ${g.dateKr(PJ.periodTo)}`, bw[1], { align: A.LEFT }), hc("사 업 주 체", bw[2]), dc(PJ.owner, bw[3])] }),
      new TableRow({ children: [hc("총 사 업 비", bw[0]), dc(money(PJ.totalBudget), bw[1], { align: A.LEFT, bold: true }), hc("주 관 부 서", bw[2]), dc(PJ.ownerDept, bw[3])] }),
    ]));

    ch.push(H2("3. 총사업비 재원구성"));
    const fw = [2000, 1500, 3000, TW - 6500];
    const frows = [new TableRow({ children: [hc("재 원", fw[0]), hc("부담비율", fw[1]), hc("교부(예정)액", fw[2]), hc("교부기관", fw[3])] })];
    PJ.fund.forEach((f) => frows.push(new TableRow({
      children: [dc(f.name, fw[0], { bold: true }), dc(pct(f.rate, 0), fw[1]), dc(money(f.amount), fw[2], { align: A.RIGHT }), dc(f.ministry, fw[3])],
    })));
    frows.push(new TableRow({
      children: [dc("합 계", fw[0], { bold: true, fill: C.alt }), dc("100%", fw[1], { bold: true, fill: C.alt }),
        dc(money(PJ.totalBudget), fw[2], { align: A.RIGHT, bold: true, fill: C.alt }), dc("-", fw[3], { fill: C.alt })],
    }));
    ch.push(table(fw, frows));
    ch.push(tblNote("(단위 : 원)"));
    ch.push(note(`총사업비 ${moneyKr(PJ.totalBudget)}`));

    ch.push(H2("4. 추진체계 및 컨소시엄 구성"));
    const cw = [700, 3300, 2500, TW - 6500];
    const crows = [new TableRow({ children: [hc("연번", cw[0]), hc("기관명", cw[1]), hc("구분", cw[2]), hc("주요 수행범위", cw[3])] })];
    PJ.consortium.forEach((c, i) => crows.push(new TableRow({
      children: [dc(String(i + 1), cw[0]), dc(c.org, cw[1], { align: A.LEFT, bold: i === 0 }), dc(c.role, cw[2], { size: SZ.tiny }), dc(c.scope, cw[3], { align: A.LEFT, size: SZ.tiny })],
    })));
    ch.push(table(cw, crows));
    ch.push(note(`사업주체(${PJ.owner}) 총괄 하에 직접보조사업자 1개 기관 및 간접보조사업자 ${PJ.consortium.length - 1}개 기관이 협약에 따라 사업 수행 중임`));
    ch.push(BREAK());
    }

    /* ══════════ Ⅱ. 추진현황 총괄 ══════════ */
    ch.push(CH("추진현황 총괄"));
    ch.push(H2("1. 핵심 관리지표"));
    const kw = [1928, 1928, 1928, 1927, TW - 7711];
    ch.push(table(kw, [
      new TableRow({ children: [hc("예산집행률", kw[0], { dark: true }), hc("실적공정률", kw[1], { dark: true }), hc("참여인력", kw[2], { dark: true }), hc("취득재산", kw[3], { dark: true }), hc("준공까지", kw[4], { dark: true })] }),
      new TableRow({
        children: [
          dc(pct(bms.execRate, 2), kw[0], { bold: true, size: SZ.body }),
          dc(pct(wbs.overall.actualRate), kw[1], { bold: true, size: SZ.body }),
          dc(`${hr.total}명`, kw[2], { bold: true, size: SZ.body }),
          dc(`${asset.total}점`, kw[3], { bold: true, size: SZ.body }),
          dc(`D-${ctx.dday}`, kw[4], { bold: true, size: SZ.body }),
        ],
      }),
      new TableRow({
        children: [
          dc(`${eokStr(bms.totalExec)} / ${eokStr(bms.totalBudget)}`, kw[0], { size: SZ.tiny, color: C.gray }),
          dc(`계획 ${pct(wbs.overall.plannedRate)}`, kw[1], { size: SZ.tiny, color: C.gray }),
          dc(`가동 ${hr.active}명`, kw[2], { size: SZ.tiny, color: C.gray }),
          dc(eokStr(asset.value), kw[3], { size: SZ.tiny, color: C.gray }),
          dc(`기간소진 ${pct(ctx.timePct)}`, kw[4], { size: SZ.tiny, color: C.gray }),
        ],
      }),
    ]));

    ch.push(H2("2. 보고 요지"));
    ch.push(box([
      ...an.summary.map((s, i) => new Paragraph({
        spacing: { before: 40, after: 40, line: 320, lineRule: "auto" },
        children: [run(`${i + 1}. `, { bold: true, size: SZ.small, color: C.navy }), run(s, { size: SZ.small })],
      })),
    ]));

    ch.push(H2("3. 주요 추진성과"));
    an.achievements.forEach((a) => ch.push(ci(a)));
    ch.push(BREAK());

    /* ══════════ Ⅲ. 예산 집행현황 ══════════ */
    ch.push(CH("예산 집행현황"));
    ch.push(sq("집행 총괄"));
    an.budgetPoints.forEach((p) => ci(p) && ch.push(ci(p)));
    ch.push(note(`기준일 : ${bms.updatedAt || ctx.dateKr}  /  자료출처 : 예산집행관리시스템(BMS) 실시간 연계`));

    ch.push(H2("1. 재원별 집행현황"));
    const sw = [1100, 800, 1950, 1950, 1950, TW - 7750];
    const srows = [new TableRow({ children: [hc("재원", sw[0]), hc("비율", sw[1]), hc("예산액", sw[2]), hc("집행액", sw[3]), hc("집행잔액", sw[4]), hc("집행률", sw[5])] })];
    bms.sources.forEach((s) => srows.push(new TableRow({
      children: [dc(s.name, sw[0], { bold: true }), dc(pct(s.rate, 0), sw[1]),
        dc(comma(s.budget), sw[2], { align: A.RIGHT }), dc(comma(s.exec), sw[3], { align: A.RIGHT }),
        dc(comma(s.remain), sw[4], { align: A.RIGHT }), dc(pct(bms.execRate, 2), sw[5])],
    })));
    srows.push(new TableRow({
      children: [dc("합 계", sw[0], { bold: true, fill: C.alt }), dc("100%", sw[1], { fill: C.alt }),
        dc(comma(bms.totalBudget), sw[2], { align: A.RIGHT, bold: true, fill: C.alt }),
        dc(comma(bms.totalExec), sw[3], { align: A.RIGHT, bold: true, fill: C.alt }),
        dc(comma(bms.totalRemain), sw[4], { align: A.RIGHT, bold: true, fill: C.alt }),
        dc(pct(bms.execRate, 2), sw[5], { bold: true, fill: C.alt })],
    }));
    ch.push(table(sw, srows));
    ch.push(tblNote("(단위 : 원)"));
    ch.push(note("재원별 집행액은 재원별 실계정 집행자료 미연계로 교부비율에 따라 안분한 추정치이며, 정산 시에는 보조금 교부·집행 실적 기준으로 재산출하여야 함"));

    if (isAnnual || type === "quarterly") {
      ch.push(H2("2. 연도별 집행현황"));
      const yw = [1400, 2100, 2100, 2100, TW - 7700];
      const yrows = [new TableRow({ children: [hc("회계연도", yw[0]), hc("예산액", yw[1]), hc("집행액", yw[2]), hc("집행잔액", yw[3]), hc("집행률", yw[4])] })];
      bms.byYear.forEach((y) => yrows.push(new TableRow({
        children: [dc(`${y.year}년`, yw[0], { bold: true }), dc(comma(y.budget), yw[1], { align: A.RIGHT }),
          dc(comma(y.exec), yw[2], { align: A.RIGHT }), dc(comma(y.remain), yw[3], { align: A.RIGHT }), dc(pct(y.rate, 2), yw[4])],
      })));
      if (!bms.byYear.length) yrows.push(new TableRow({ children: [dc("연도별 자료 미연계", yw[0], { span: 5, color: C.gray })] }));
      ch.push(table(yw, yrows));
      ch.push(tblNote("(단위 : 원)"));
    }

    ch.push(H2(`${isAnnual || type === "quarterly" ? "3" : "2"}. 비목별 집행현황`));
    const mw = [2000, 600, 1950, 1950, 1950, TW - 8450];
    const mrows = [new TableRow({ children: [hc("비목", mw[0]), hc("건수", mw[1]), hc("예산액", mw[2]), hc("집행액", mw[3]), hc("집행잔액", mw[4]), hc("집행률", mw[5])] })];
    bms.cats.forEach((c) => mrows.push(new TableRow({
      children: [dc(c.name, mw[0], { align: A.LEFT, bold: true }), dc(String(c.cnt), mw[1]),
        dc(comma(c.budget), mw[2], { align: A.RIGHT }), dc(comma(c.exec), mw[3], { align: A.RIGHT }),
        dc(comma(c.remain), mw[4], { align: A.RIGHT }),
        dc(pct(c.rate, 2), mw[5], { bold: true, color: c.rate >= 80 ? C.ok : c.rate >= 30 ? C.warn : C.bad })],
    })));
    mrows.push(new TableRow({
      children: [dc("합 계", mw[0], { bold: true, fill: C.alt }), dc(String(bms.itemCount), mw[1], { fill: C.alt }),
        dc(comma(bms.totalBudget), mw[2], { align: A.RIGHT, bold: true, fill: C.alt }),
        dc(comma(bms.totalExec), mw[3], { align: A.RIGHT, bold: true, fill: C.alt }),
        dc(comma(bms.totalRemain), mw[4], { align: A.RIGHT, bold: true, fill: C.alt }),
        dc(pct(bms.execRate, 2), mw[5], { bold: true, fill: C.alt })],
    }));
    ch.push(table(mw, mrows));
    ch.push(tblNote("(단위 : 원)"));

    ch.push(H2(`${isAnnual || type === "quarterly" ? "4" : "3"}. 단위사업별 집행현황`));
    const uw = [600, 2700, 1780, 1780, 1780, TW - 8640];
    const urows = [new TableRow({ children: [hc("연번", uw[0]), hc("단위사업명", uw[1]), hc("예산액", uw[2]), hc("집행액", uw[3]), hc("집행잔액", uw[4]), hc("집행률", uw[5])] })];
    const projList = isWeekly
      ? bms.projects.slice().sort((a, b) => b.rate - a.rate).filter((_, i, arr) => i < 3 || i >= arr.length - 3)
      : bms.projects;
    projList.forEach((p, i) => urows.push(new TableRow({
      children: [dc(String(i + 1), uw[0]), dc(p.name, uw[1], { align: A.LEFT }),
        dc(comma(p.budget), uw[2], { align: A.RIGHT }), dc(comma(p.exec), uw[3], { align: A.RIGHT }),
        dc(comma(p.remain), uw[4], { align: A.RIGHT }),
        dc(pct(p.rate, 1), uw[5], { bold: true, color: p.rate >= 80 ? C.ok : p.rate >= 30 ? C.warn : C.bad })],
    })));
    if (!isWeekly && bms.commonB > 0) urows.push(new TableRow({
      children: [dc("-", uw[0]), dc("공통비(인건비·운영비 등)", uw[1], { align: A.LEFT, color: C.gray }),
        dc(comma(bms.commonB), uw[2], { align: A.RIGHT }), dc(comma(bms.commonE), uw[3], { align: A.RIGHT }),
        dc(comma(bms.commonB - bms.commonE), uw[4], { align: A.RIGHT }), dc(pct((bms.commonE / (bms.commonB || 1)) * 100, 1), uw[5])],
    }));
    ch.push(table(uw, urows));
    ch.push(tblNote(isWeekly ? "(단위 : 원) ※ 집행률 상·하위 각 3건" : "(단위 : 원)"));
    if (!isWeekly) ch.push(BREAK());

    /* ══════════ Ⅳ. 공정 추진현황 ══════════ */
    ch.push(CH("공정(WBS) 추진현황"));
    ch.push(sq("공정 총괄"));
    an.wbsPoints.forEach((p) => ch.push(ci(p)));
    ch.push(note(`기준일 : ${wbs.overall.updatedAt || ctx.dateKr}  /  자료출처 : 사업공정관리시스템(WBS) 실시간 연계`));

    ch.push(H2("1. 공정 종합현황"));
    const pw = [1376, 1376, 1376, 1376, 1376, 1376, TW - 8256];
    ch.push(table(pw, [
      new TableRow({ children: [hc("전체", pw[0]), hc("완료", pw[1]), hc("진행", pw[2]), hc("지연", pw[3]), hc("대기", pw[4]), hc("계획공정률", pw[5]), hc("실적공정률", pw[6])] }),
      new TableRow({
        children: [dc(`${wbs.overall.total}건`, pw[0], { bold: true }), dc(`${wbs.overall.done}건`, pw[1], { color: C.ok }),
          dc(`${wbs.overall.inProg}건`, pw[2]), dc(`${wbs.overall.delayed}건`, pw[3], { color: C.bad, bold: true }),
          dc(`${wbs.overall.waiting}건`, pw[4]), dc(pct(wbs.overall.plannedRate), pw[5]),
          dc(pct(wbs.overall.actualRate), pw[6], { bold: true })],
      }),
    ]));

    ch.push(H2("2. 대분류별 공정현황"));
    const cw2 = [2600, 700, 700, 700, 700, 1180, 1180, TW - 7760];
    const crows2 = [new TableRow({ children: [hc("대분류", cw2[0]), hc("전체", cw2[1]), hc("완료", cw2[2]), hc("지연", cw2[3]), hc("대기", cw2[4]), hc("계획", cw2[5]), hc("실적", cw2[6]), hc("편차", cw2[7])] })];
    wbs.byCat.forEach((c) => crows2.push(new TableRow({
      children: [dc(c.name, cw2[0], { align: A.LEFT }), dc(String(c.total), cw2[1]), dc(String(c.done), cw2[2]),
        dc(String(c.delayed), cw2[3], { color: c.delayed ? C.bad : C.text }), dc(String(c.waiting), cw2[4]),
        dc(pct(c.plannedRate), cw2[5]), dc(pct(c.actualRate), cw2[6], { bold: true }),
        dc(`${c.deviation > 0 ? "+" : ""}${f1(c.deviation)}%p`, cw2[7], { color: c.deviation < 0 ? C.bad : C.ok })],
    })));
    ch.push(table(cw2, crows2));

    if (wbs.units.length) {
      ch.push(H2("3. 단위사업별 추진현황"));
      const vw = [700, 3600, 1500, 1500, 1300, TW - 8600];
      const vrows = [new TableRow({ children: [hc("연번", vw[0]), hc("단위사업명", vw[1]), hc("계획공정", vw[2]), hc("실적공정", vw[3]), hc("편차", vw[4]), hc("상태", vw[5])] })];
      wbs.units.forEach((u, i) => vrows.push(new TableRow({
        children: [dc(String(i + 1), vw[0]), dc(u.name, vw[1], { align: A.LEFT }),
          dc(pct(u.planned), vw[2]), dc(pct(u.actual), vw[3], { bold: true }),
          dc(`${u.deviation > 0 ? "+" : ""}${f1(u.deviation)}%p`, vw[4], { color: u.deviation < 0 ? C.bad : C.ok }),
          dc(u.status, vw[5], { color: u.status === "지연" ? C.bad : C.text })],
      })));
      ch.push(table(vw, vrows));
    }
    if (!isWeekly) ch.push(BREAK());

    /* ══════════ Ⅴ. 참여인력 ══════════ */
    ch.push(CH("참여인력 운영현황"));
    ch.push(sq("인력 운영 총괄"));
    an.hrPoints.forEach((p) => ch.push(ci(p)));

    ch.push(H2("1. 기관별 참여인력 현황"));
    const hw = [2600, 1500, 1000, 1000, 1000, 1000, TW - 8100];
    const hrows = [new TableRow({ children: [hc("기관명", hw[0]), hc("구분", hw[1]), hc("등록", hw[2]), hc("실인원", hw[3]), hc("가동", hw[4]), hc("종료", hw[5]), hc("참여율계", hw[6])] })];
    (hr.orgs || []).forEach((o) => hrows.push(new TableRow({
      children: [dc(o.org, hw[0], { align: A.LEFT, size: SZ.tiny }), dc(o.role || "-", hw[1], { size: SZ.tiny }),
        dc(`${o.total}명`, hw[2]), dc(`${o.headcount}명`, hw[3]), dc(`${o.active}명`, hw[4], { bold: true }),
        dc(`${o.ended}명`, hw[5], { color: C.gray }), dc(pct(o.activeManRatio), hw[6])],
    })));
    hrows.push(new TableRow({
      children: [dc("합 계", hw[0], { bold: true, fill: C.alt }), dc("-", hw[1], { fill: C.alt }),
        dc(`${hr.total}명`, hw[2], { bold: true, fill: C.alt }), dc(`${hr.headcount}명`, hw[3], { bold: true, fill: C.alt }),
        dc(`${hr.active}명`, hw[4], { bold: true, fill: C.alt }), dc(`${hr.ended}명`, hw[5], { fill: C.alt }),
        dc(pct(hr.activeManRatio), hw[6], { bold: true, fill: C.alt })],
    }));
    ch.push(table(hw, hrows));
    ch.push(note("등록인원은 역할별 등재 건수, 실인원은 성명·직급 기준 중복 제거 인원임. 참여율계는 가동인력의 협약상 참여율 합계"));
    ch.push(note("인건비 정산은 「보조금 관리에 관한 법률」 및 사업비 집행기준에 따라 실지급액 × 참여기간 × 참여율 기준으로 산정하며, 참여율 변경 시 사전 변경승인 절차 이행"));

    /* ══════════ Ⅵ. 재산 ══════════ */
    ch.push(CH("재산 취득·관리 현황"));
    ch.push(sq("재산 총괄"));
    an.assetPoints.forEach((p) => ch.push(ci(p)));

    ch.push(H2("1. 분류별 취득현황"));
    const aw = [3000, 1300, 2400, 1600, TW - 8300];
    const arows = [new TableRow({ children: [hc("자산분류", aw[0]), hc("수량", aw[1]), hc("취득가액", aw[2]), hc("구성비", aw[3]), hc("비고", aw[4])] })];
    asset.byCat.forEach((c) => arows.push(new TableRow({
      children: [dc(c.name, aw[0], { align: A.LEFT }), dc(`${c.count}점`, aw[1]), dc(comma(c.value), aw[2], { align: A.RIGHT }),
        dc(pct((c.value / (asset.value || 1)) * 100), aw[3]), dc("", aw[4])],
    })));
    arows.push(new TableRow({
      children: [dc("합 계", aw[0], { bold: true, fill: C.alt }), dc(`${asset.total}점`, aw[1], { bold: true, fill: C.alt }),
        dc(comma(asset.value), aw[2], { align: A.RIGHT, bold: true, fill: C.alt }), dc("100.0%", aw[3], { fill: C.alt }), dc("", aw[4], { fill: C.alt })],
    }));
    ch.push(table(aw, arows));
    ch.push(tblNote("(단위 : 원)"));
    ch.push(note("「보조금 관리에 관한 법률」 제35조에 따른 중요재산은 처분 시 중앙관서의 장의 승인을 받아야 하며, 자산관리대장·재물조사 결과를 정산 시 함께 제출하여야 함"));

    if (!isWeekly && asset.major && asset.major.length) {
      ch.push(H2("2. 중요재산 관리대상 (취득가액 5천만원 이상)"));
      const jw = [700, 3400, 1700, 2000, TW - 7800];
      const jrows = [new TableRow({ children: [hc("연번", jw[0]), hc("재산명", jw[1]), hc("분류", jw[2]), hc("취득가액", jw[3]), hc("설치위치", jw[4])] })];
      asset.major.slice(0, 15).forEach((m, i) => jrows.push(new TableRow({
        children: [dc(String(i + 1), jw[0]), dc(m.name, jw[1], { align: A.LEFT, size: SZ.tiny }), dc(m.cat, jw[2], { size: SZ.tiny }),
          dc(comma(m.value), jw[3], { align: A.RIGHT }), dc(m.loc, jw[4], { size: SZ.tiny })],
      })));
      ch.push(table(jw, jrows));
      ch.push(tblNote("(단위 : 원)"));
    }
    if (!isWeekly) ch.push(BREAK());

    /* ══════════ Ⅶ. 이슈 ══════════ */
    ch.push(CH("주요 이슈 및 조치계획"));
    const iw = [600, 3000, 2400, 2400, TW - 8400];
    const irows = [new TableRow({ children: [hc("연번", iw[0]), hc("주요 이슈", iw[1]), hc("영향", iw[2]), hc("조치계획", iw[3]), hc("중요도", iw[4])] })];
    an.issues.forEach((s, i) => irows.push(new TableRow({
      children: [dc(String(i + 1), iw[0]), dc(s.issue, iw[1], { align: A.LEFT, size: SZ.tiny }),
        dc(s.impact, iw[2], { align: A.LEFT, size: SZ.tiny }), dc(s.action, iw[3], { align: A.LEFT, size: SZ.tiny }),
        dc(s.grade, iw[4], { bold: true, size: SZ.tiny, color: s.grade === "매우 높음" ? C.bad : s.grade === "높음" ? C.warn : C.text })],
    })));
    ch.push(table(iw, irows));

    /* ══════════ Ⅷ. 향후계획 ══════════ */
    ch.push(CH("향후 추진계획"));
    ch.push(sq(isWeekly ? "차주 추진계획" : "차기 보고기간 추진계획"));
    an.plans.forEach((p) => ch.push(ci(p)));
    ch.push(sq("협조 및 조치 요청사항"));
    an.recos.forEach((r) => ch.push(ci(r)));

    /* ══════════ Ⅸ. 종합의견 ══════════ */
    ch.push(CH("종합의견"));
    ch.push(box([
      new Paragraph({
        spacing: { before: 40, after: 80, line: 330, lineRule: "auto" },
        children: [run("□ 총괄 판단", { bold: true, size: SZ.small, color: C.navy })],
      }),
      new Paragraph({
        spacing: { before: 20, after: 60, line: 330, lineRule: "auto" },
        indent: { left: 200 }, children: [run(an.judgment, { size: SZ.small })],
      }),
    ]));
    ch.push(EMPTY(120));
    ch.push(ci(`작성 : ${DS.senderOrg} ${DS.handlerDept}`));
    ch.push(ci(`검토 : ${DS.reviewer.title} ${DS.reviewer.name}`));
    ch.push(ci(`보고일 : ${ctx.dateKr}`));
    if (!isWeekly) ch.push(BREAK());

    /* ══════════ 붙임 ══════════ */
    ch.push(H1("붙임. 세부 현황자료"));
    ch.push(H2("붙임 1. 공정 지연 목록"));
    if (wbs.delayed.length) {
      const dw = [1400, 3600, 1300, 1300, TW - 7600];
      const drows = [new TableRow({ children: [hc("WBS ID", dw[0]), hc("작업명", dw[1]), hc("계획", dw[2]), hc("실적", dw[3]), hc("종료예정일", dw[4])] })];
      wbs.delayed.slice(0, 25).forEach((d) => drows.push(new TableRow({
        children: [dc(d.wbsId || "-", dw[0], { size: SZ.tiny }), dc(d.name, dw[1], { align: A.LEFT, size: SZ.tiny }),
          dc(pct(d.planned), dw[2], { size: SZ.tiny }), dc(pct(d.actual), dw[3], { size: SZ.tiny, color: C.bad }),
          dc(String(d.endDate || "-").slice(0, 10), dw[4], { size: SZ.tiny })],
      })));
      ch.push(table(dw, drows));
    } else ch.push(ci("해당 없음"));

    ch.push(H2("붙임 2. 미집행 단위사업 목록"));
    if (an.zero.length) {
      const zw = [700, 4600, 2200, TW - 7500];
      const zrows = [new TableRow({ children: [hc("연번", zw[0]), hc("단위사업명", zw[1]), hc("예산액", zw[2]), hc("조치사항", zw[3])] })];
      an.zero.forEach((z, i) => zrows.push(new TableRow({
        children: [dc(String(i + 1), zw[0]), dc(z.name, zw[1], { align: A.LEFT }), dc(comma(z.budget), zw[2], { align: A.RIGHT }),
          dc("발주절차 착수", zw[3], { size: SZ.tiny })],
      })));
      ch.push(table(zw, zrows));
      ch.push(tblNote("(단위 : 원)"));
    } else ch.push(ci("해당 없음"));

    if (!isWeekly && hr.hasRoster) {
      ch.push(H2("붙임 3. 참여인력 명부 (가동인력)"));
      const rw = [600, 2600, 1500, 1100, 1100, TW - 6900];
      const rrows = [new TableRow({ children: [hc("연번", rw[0]), hc("소속기관", rw[1]), hc("성명", rw[2]), hc("직급", rw[3]), hc("참여율", rw[4]), hc("참여기간", rw[5])] })];
      hr.roster.filter((p) => p.status === "활성").slice(0, 40).forEach((p, i) => rrows.push(new TableRow({
        children: [dc(String(i + 1), rw[0]), dc(p.org, rw[1], { align: A.LEFT, size: SZ.tiny }),
          dc(p.name, rw[2], { size: SZ.tiny }), dc(p.position, rw[3], { size: SZ.tiny }),
          dc(pct(p.ratio, 0), rw[4], { size: SZ.tiny, bold: p.ratio > 100, color: p.ratio > 100 ? C.bad : C.text }),
          dc(`${p.from} ~ ${p.to}`, rw[5], { size: SZ.tiny })],
      })));
      ch.push(table(rw, rrows));
      ch.push(note("참여율이 100%를 초과하는 경우 인건비 중복계상에 해당하므로 참여율 재산정 및 정정이 필요함"));
    }

    ch.push(EMPTY(200));
    ch.push(new Paragraph({
      alignment: A.RIGHT, spacing: { before: 200, after: 0, line: 300 },
      children: [run("끝.", { bold: true, size: SZ.body })],
    }));

    return new Document({
      creator: DS.senderOrg, title: `${PJ.name} ${T.title}`,
      description: `${ctx.period} / 문서번호 ${DS.docNo}`,
      styles: g.styles(),
      sections: [{
        properties: g.pageProps(),
        titlePage: !isWeekly,
        headers: { default: g.header(`${PJ.name} · ${T.title}`), first: g.headerBlank() },
        footers: { default: g.footer() },
        children: ch,
      }],
    });
  }

  return { build };
});

/* ══════════════════════════════════════════════════════════════════════════
 *  doc-brief.js — 핵심 추진현황 보고(2매) 빌더
 *  아산시장·부시장 대면보고용. 1매: 현황총괄 / 2매: 쟁점·건의사항
 * ══════════════════════════════════════════════════════════════════════════ */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.DocBrief = factory();
})(typeof self !== "undefined" ? self : this, function () {

  function build(docx, P) {
    const { Document, Paragraph, TableRow } = docx;
    const { bms, wbs, asset, hr, an, ctx, CFG, DS, G } = P;
    const g = G;
    const { TW, C, SZ, run, EMPTY, BREAK, sq, ci, dash, note, hc, dc, table, tblNote, box,
      comma, money, eokStr, pct, AlignmentType: A } = g;
    const PJ = CFG.PROJECT;
    const f1 = (v) => Number(v || 0).toFixed(1);
    const ch = [];
    // 2매 제약 — 컴팩트 문단 헬퍼(기본 개조식 대비 여백·자간 축소)
    const SQ = (t) => new Paragraph({
      spacing: { before: 130, after: 40, line: 300, lineRule: "auto" },
      children: [run("□ ", { bold: true, size: SZ.small, color: C.navy }), run(t, { bold: true, size: SZ.small, color: C.navy })],
    });
    const CI = (t) => new Paragraph({
      indent: { left: 300 }, spacing: { before: 20, after: 20, line: 290, lineRule: "auto" },
      children: [run("○ ", { size: SZ.small }), run(t, { size: SZ.small })],
    });
    const DA = (t) => new Paragraph({
      indent: { left: 620 }, spacing: { before: 14, after: 14, line: 280, lineRule: "auto" },
      children: [run("- ", { size: SZ.tiny }), run(t, { size: SZ.tiny, color: C.gray })],
    });

    /* ── 제목부 ── */
    ch.push(new Paragraph({
      alignment: A.RIGHT, spacing: { before: 0, after: 40, line: 240 },
      children: [run(`문서번호 ${DS.docNo}  /  ${DS.openLevel}`, { size: SZ.tiny, color: C.gray })],
    }));
    ch.push(new Paragraph({
      alignment: A.CENTER, spacing: { before: 60, after: 60, line: 300 },
      children: [run(`${PJ.name} 추진현황`, { size: 34, bold: true, color: C.navy })],
    }));
    ch.push(new Paragraph({
      alignment: A.CENTER, spacing: { before: 0, after: 160, line: 260 },
      border: { bottom: { style: docx.BorderStyle.SINGLE, size: 12, color: C.navy } },
      children: [run(`${ctx.period}  ·  ${DS.senderOrg} ${DS.handlerDept}`, { size: SZ.small, color: C.gray })],
    }));

    /* ── 1. 핵심지표 ── */
    ch.push(SQ("핵심 관리지표"));
    const kw = [1928, 1928, 1928, 1927, TW - 7711];
    ch.push(table(kw, [
      new TableRow({ children: [hc("예산집행률", kw[0], { dark: true }), hc("실적공정률", kw[1], { dark: true }), hc("참여인력", kw[2], { dark: true }), hc("취득재산", kw[3], { dark: true }), hc("준공까지", kw[4], { dark: true })] }),
      new TableRow({
        children: [dc(pct(bms.execRate, 2), kw[0], { bold: true, size: SZ.body }),
          dc(pct(wbs.overall.actualRate), kw[1], { bold: true, size: SZ.body }),
          dc(`${hr.total}명`, kw[2], { bold: true, size: SZ.body }),
          dc(`${asset.total}점`, kw[3], { bold: true, size: SZ.body }),
          dc(`D-${ctx.dday}`, kw[4], { bold: true, size: SZ.body })],
      }),
      new TableRow({
        children: [dc(`${eokStr(bms.totalExec)}/${eokStr(bms.totalBudget)}`, kw[0], { size: SZ.tiny, color: C.gray }),
          dc(`계획 ${pct(wbs.overall.plannedRate)}`, kw[1], { size: SZ.tiny, color: C.gray }),
          dc(`가동 ${hr.active}명`, kw[2], { size: SZ.tiny, color: C.gray }),
          dc(eokStr(asset.value), kw[3], { size: SZ.tiny, color: C.gray }),
          dc(`기간소진 ${pct(ctx.timePct)}`, kw[4], { size: SZ.tiny, color: C.gray })],
      }),
    ]));

    /* ── 2. 추진현황 요지 ── */
    ch.push(SQ("추진현황 요지"));
    an.summary.forEach((s) => ch.push(CI(s)));

    /* ── 3. 예산 ── */
    ch.push(SQ("예산 집행현황"));
    const sw = [1800, 2400, 2400, TW - 8000];
    const srows = [new TableRow({ children: [hc("재원", sw[0]), hc("예산액", sw[1]), hc("집행액", sw[2]), hc("집행률", sw[3])] })];
    bms.sources.forEach((s) => srows.push(new TableRow({
      children: [dc(`${s.name}(${pct(s.rate, 0)})`, sw[0], { bold: true }), dc(comma(s.budget), sw[1], { align: A.RIGHT }),
        dc(comma(s.exec), sw[2], { align: A.RIGHT }), dc(pct(bms.execRate, 2), sw[3])],
    })));
    srows.push(new TableRow({
      children: [dc("합 계", sw[0], { bold: true, fill: C.alt }), dc(comma(bms.totalBudget), sw[1], { align: A.RIGHT, bold: true, fill: C.alt }),
        dc(comma(bms.totalExec), sw[2], { align: A.RIGHT, bold: true, fill: C.alt }), dc(pct(bms.execRate, 2), sw[3], { bold: true, fill: C.alt })],
    }));
    ch.push(table(sw, srows));
    ch.push(tblNote("(단위 : 원) ※ 재원별 집행액은 교부비율 안분 추정치"));

    /* ── 4. 공정 상위/하위 ── */
    ch.push(SQ("단위사업별 추진현황 (집행률 상·하위)"));
    const top = bms.projects.slice().sort((a, b) => b.rate - a.rate).slice(0, 3);
    const bot = bms.projects.slice().sort((a, b) => a.rate - b.rate).slice(0, 3);
    const uw = [4200, 1600, 1600, TW - 7400];
    const urows = [new TableRow({ children: [hc("단위사업명", uw[0]), hc("예산액", uw[1]), hc("집행액", uw[2]), hc("집행률", uw[3])] })];
    top.concat(bot).forEach((p, i) => urows.push(new TableRow({
      children: [dc((i < top.length ? "▲ " : "▼ ") + p.name, uw[0], { align: A.LEFT, size: SZ.tiny }),
        dc(eokStr(p.budget), uw[1], { size: SZ.tiny }), dc(eokStr(p.exec), uw[2], { size: SZ.tiny }),
        dc(pct(p.rate, 1), uw[3], { bold: true, size: SZ.tiny, color: p.rate >= 80 ? C.ok : p.rate >= 30 ? C.warn : C.bad })],
    })));
    ch.push(table(uw, urows));
    ch.push(BREAK());

    /* ══ 2매 ══ */
    ch.push(SQ("주요 쟁점 및 조치계획"));
    const iw = [3400, 3000, TW - 6400];
    const irows = [new TableRow({ children: [hc("쟁점사항", iw[0]), hc("영향", iw[1]), hc("조치계획", iw[2])] })];
    an.issues.slice(0, 4).forEach((s) => irows.push(new TableRow({
      children: [dc(s.issue, iw[0], { align: A.LEFT, size: SZ.tiny }), dc(s.impact, iw[1], { align: A.LEFT, size: SZ.tiny }),
        dc(s.action, iw[2], { align: A.LEFT, size: SZ.tiny })],
    })));
    ch.push(table(iw, irows));

    ch.push(SQ("공정 지연현황"));
    ch.push(CI(`전체 ${wbs.overall.total}건 중 지연 ${wbs.overall.delayed}건(계획 대비 ${f1(Math.abs(an.wbsGap))}%p ${an.wbsGap > 0 ? "미달" : "상회"})`));
    wbs.delayed.slice(0, 4).forEach((d) => ch.push(DA(`${d.name} — 계획 ${pct(d.planned)} / 실적 ${pct(d.actual)}`)));

    ch.push(SQ("향후 추진계획"));
    an.plans.slice(0, 3).forEach((p) => ch.push(CI(p)));

    ch.push(SQ("건의사항"));
    an.recos.slice(0, 3).forEach((r) => ch.push(CI(r)));

    ch.push(EMPTY(60));
    ch.push(box([
      new Paragraph({ spacing: { before: 20, after: 60, line: 320, lineRule: "auto" }, children: [run("□ 종합의견", { bold: true, size: SZ.small, color: C.navy })] }),
      new Paragraph({ spacing: { before: 0, after: 20, line: 320, lineRule: "auto" }, indent: { left: 200 }, children: [run(an.judgment, { size: SZ.small })] }),
    ]));

    ch.push(EMPTY(120));
    ch.push(new Paragraph({
      alignment: A.RIGHT, spacing: { before: 160, after: 0, line: 300 },
      children: [run(`${ctx.dateKr}   ${DS.senderOrg}  ${DS.reviewer.title} ${DS.reviewer.name}`, { size: SZ.small })],
    }));
    ch.push(new Paragraph({
      alignment: A.RIGHT, spacing: { before: 60, after: 0, line: 300 },
      children: [run("끝.", { bold: true, size: SZ.body })],
    }));

    return new Document({
      creator: DS.senderOrg, title: `${PJ.name} 핵심 추진현황 보고`,
      styles: g.styles(),
      sections: [{
        properties: g.pageProps(),
        footers: { default: g.footer() },
        children: ch,
      }],
    });
  }

  return { build };
});

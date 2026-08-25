/* ══════════════════════════════════════════════════════════════════════════
 *  doc-official.js — 시행문(일반기안문) 빌더
 *  「행정업무의 운영 및 혁신에 관한 규정 시행규칙」 별지 제1호서식 준거
 *  두문(기관명·수신·경유·제목) → 본문(1.가.1)) → 붙임 → 끝. → 결문
 * ══════════════════════════════════════════════════════════════════════════ */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.DocOfficial = factory();
})(typeof self !== "undefined" ? self : this, function () {

  function build(docx, P) {
    const { Document, Paragraph } = docx;
    const { bms, wbs, asset, hr, an, ctx, CFG, DS, G } = P;
    const g = G;
    const { TW, C, SZ, run, EMPTY, item, sq, ci, note, comma, money, moneyKr, eokStr, pct, AlignmentType: A } = g;
    const PJ = CFG.PROJECT;
    const f1 = (v) => Number(v || 0).toFixed(1);

    const subject = `「${PJ.name}」 추진현황 통보 및 협조 요청(${ctx.year}. ${ctx.month}. 기준)`;
    const ch = [];

    /* 결재란(수기결재 시 사용) */
    ch.push(g.approvalBox(DS));
    ch.push(EMPTY(120));

    /* 두문 */
    g.head(Object.assign({}, DS, { recipientDept: DS.recipientDept || PJ.ownerDept, via: DS.via || "" }), subject).forEach((p) => ch.push(p));

    /* 본문 — 공문서 항목기호 1. → 가. → 1) */
    ch.push(item(1, 0, "귀 기관의 무궁한 발전을 기원합니다."));
    ch.push(EMPTY(60));
    ch.push(item(1, 1, `「스마트도시 조성 및 산업진흥 등에 관한 법률」 제12조 및 「${PJ.name}」 협약에 따라 추진 중인 사업의 ${ctx.period} 기준 추진현황을 다음과 같이 통보하오니 업무에 참고하여 주시기 바랍니다.`));

    ch.push(item(2, 0, `예산 집행현황 : 총사업비 ${money(bms.totalBudget)} 중 ${money(bms.totalExec)} 집행(집행률 ${pct(bms.execRate, 2)}), 집행잔액 ${money(bms.totalRemain)}`));
    ch.push(item(3, 0, `재원구성 : 국비 ${pct(PJ.fund[0].rate, 0)} · 시비 ${pct(PJ.fund[1].rate, 0)} · 도비 ${pct(PJ.fund[2].rate, 0)}`));
    ch.push(item(3, 1, `사업기간 소진율 ${pct(ctx.timePct)} 대비 집행률 ${pct(bms.execRate, 2)}로 ${f1(Math.abs(an.gap))}%p ${an.gap > 0 ? "미달" : "상회"}`));

    ch.push(item(2, 1, `공정 추진현황 : 계획공정률 ${pct(wbs.overall.plannedRate)} 대비 실적공정률 ${pct(wbs.overall.actualRate)}(달성률 ${pct(wbs.overall.achieveRate)})`));
    ch.push(item(3, 0, `전체 ${wbs.overall.total}건 중 완료 ${wbs.overall.done}건 · 진행 ${wbs.overall.inProg}건 · 지연 ${wbs.overall.delayed}건 · 대기 ${wbs.overall.waiting}건`));

    ch.push(item(2, 2, `참여인력 및 취득재산 : ${hr.orgs.length || 4}개 기관 ${hr.total}명 참여(가동 ${hr.active}명), 취득재산 ${asset.total}점(취득가액 ${money(asset.value)})`));

    ch.push(item(2, 3, `준공 잔여기간 : ${ctx.monthsLeft}개월(${g.dateKr(PJ.periodTo)} 준공 예정)`));

    ch.push(EMPTY(60));
    ch.push(item(1, 2, "아울러 사업의 원활한 추진을 위하여 다음 사항에 대한 협조를 요청드립니다."));
    an.recos.slice(0, 4).forEach((r, i) => ch.push(item(2, i, r)));

    ch.push(EMPTY(60));
    ch.push(item(1, 3, `본 통보사항에 대한 세부 현황은 붙임 자료를 참고하시기 바라며, 관련 문의사항은 ${DS.handlerDept}(${DS.tel})로 연락하여 주시기 바랍니다.`));

    /* 요약 표 */
    ch.push(EMPTY(140));
    const cw = [2200, 2400, 2200, TW - 6800];
    const { hc, dc, table } = g;
    const { TableRow } = docx;
    ch.push(table(cw, [
      new TableRow({ children: [hc("구 분", cw[0]), hc("계획 / 예산", cw[1]), hc("실적 / 집행", cw[2]), hc("달성 · 집행률", cw[3])] }),
      new TableRow({ children: [dc("예산집행", cw[0], { bold: true }), dc(comma(bms.totalBudget) + "원", cw[1], { align: A.RIGHT }), dc(comma(bms.totalExec) + "원", cw[2], { align: A.RIGHT }), dc(pct(bms.execRate, 2), cw[3], { bold: true })] }),
      new TableRow({ children: [dc("공정추진", cw[0], { bold: true }), dc(pct(wbs.overall.plannedRate), cw[1]), dc(pct(wbs.overall.actualRate), cw[2]), dc(pct(wbs.overall.achieveRate), cw[3], { bold: true })] }),
      new TableRow({ children: [dc("참여인력", cw[0], { bold: true }), dc(`${hr.total}명`, cw[1]), dc(`가동 ${hr.active}명`, cw[2]), dc(pct(hr.rate), cw[3], { bold: true })] }),
      new TableRow({ children: [dc("취득재산", cw[0], { bold: true }), dc("-", cw[1]), dc(`${asset.total}점 / ${comma(asset.value)}원`, cw[2], { align: A.RIGHT }), dc(pct(asset.issuedRate), cw[3], { bold: true })] }),
    ]));
    ch.push(note(`기준일 : ${ctx.dateKr}  /  자료출처 : 예산집행관리시스템(BMS)·사업공정관리시스템(WBS)·자산관리시스템 실시간 연계`));

    /* 붙임 + 끝. */
    g.attachments([
      { name: `${PJ.name} 추진현황 세부자료`, copies: "1부" },
      { name: "예산 집행현황 및 공정현황 총괄표", copies: "1부" },
      { name: "주요 이슈 및 조치계획", copies: "1부" },
    ]).forEach((p) => ch.push(p));

    /* 결문 */
    g.foot(DS, ctx).forEach((p) => ch.push(p));

    return new Document({
      creator: DS.senderOrg, title: subject,
      description: `문서번호 ${DS.docNo} / 시행일 ${ctx.dateKr}`,
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

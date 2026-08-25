/* ══════════════════════════════════════════════════════════════════════════
 *  gongmun.js — 공문서 규격 엔진
 *  근거: 「행정업무의 운영 및 혁신에 관한 규정」(대통령령) 및 같은 규정 시행규칙
 *        별지 제1호서식(일반기안문) / 제7조(문서 작성의 일반원칙)
 *  - 용지 A4, 여백 위30·아래15·좌우20mm
 *  - 글꼴 맑은 고딕 12pt, 줄간격 160%
 *  - 항목기호 1. → 가. → 1) → 가) → (1) → (가) → ①
 *  - 날짜 "2026. 8. 25." / 금액 "금10,000원(금일만원)"
 *  - 본문 종결 "끝." 표기, 붙임 "붙임  1. ○○ 1부."
 * ══════════════════════════════════════════════════════════════════════════ */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Gongmun = factory();
})(typeof self !== "undefined" ? self : this, function () {

  const MM = 56.7;                       // 1mm = 56.7 twip
  const PAGE = { width: 11906, height: 16838 };
  const MARGIN = {
    top: Math.round(30 * MM),            // 30mm
    bottom: Math.round(15 * MM),         // 15mm
    left: Math.round(20 * MM),           // 20mm
    right: Math.round(20 * MM),          // 20mm
    header: Math.round(10 * MM),
    footer: Math.round(8 * MM),
  };
  const TW = PAGE.width - MARGIN.left - MARGIN.right;   // 본문 폭 9638 twip
  const FONT = "맑은 고딕";
  const FONT_ALT = "휴먼명조";

  // 색상 — 공공기관 문서 톤(무채색 기조 + 남색 강조). 컬러 남용 금지
  const C = {
    text: "000000", navy: "1F3864", head: "2F5496", line: "808080",
    alt: "F2F2F2", head2: "D9E2F3", gray: "595959", light: "A6A6A6",
    ok: "1F6F3F", warn: "9C6500", bad: "9C0006", white: "FFFFFF",
  };

  const SZ = { body: 24, small: 20, tiny: 18, h1: 28, h2: 24, title: 40, sub: 26 };
  const LINE = { line: 384, lineRule: "auto" };          // 줄간격 160%

  /* ── 숫자·금액·날짜 표기 ────────────────────────────────────────────── */
  const comma = (n) => Math.round(+n || 0).toLocaleString("ko-KR");

  function hanNumber(n) {
    const D = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"];
    const S = ["", "십", "백", "천"];
    const B = ["", "만", "억", "조", "경"];
    n = Math.floor(Math.abs(+n || 0));
    if (n === 0) return "영";
    let out = "", bi = 0;
    while (n > 0) {
      const chunk = n % 10000; n = Math.floor(n / 10000);
      if (chunk) {
        let cs = "";
        for (let i = 0; i < 4; i++) {
          const d = Math.floor(chunk / Math.pow(10, i)) % 10;
          if (d) cs = D[d] + S[i] + cs;
        }
        out = cs + B[bi] + out;
      }
      bi++;
    }
    return out;
  }
  /** 공문서 금액표기: 금10,205,027,657원(금일백이억오백이만칠천육백오십칠원) */
  const moneyKr = (n) => `금${comma(n)}원(금${hanNumber(n)}원)`;
  const money = (n) => `${comma(n)}원`;
  const eokStr = (n, d) => `${(( +n || 0) / 1e8).toFixed(d == null ? 1 : d)}억원`;
  const pct = (n, d) => `${Number(n || 0).toFixed(d == null ? 1 : d)}%`;
  /** 공문서 날짜표기: 2026. 8. 25. */
  const dateKr = (d) => { const x = d instanceof Date ? d : new Date(d); return `${x.getFullYear()}. ${x.getMonth() + 1}. ${x.getDate()}.`; };
  const ymKr = (d) => { const x = d instanceof Date ? d : new Date(d); return `${x.getFullYear()}. ${x.getMonth() + 1}.`; };

  /* ── 팩토리: docx 네임스페이스 주입 ─────────────────────────────────── */
  function create(docx) {
    const {
      Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
      AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
      PageBreak, PageNumber, HeadingLevel,
    } = docx;

    const brd = (sz, col) => ({ style: BorderStyle.SINGLE, size: sz || 4, color: col || C.line });
    const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const boxAll = { top: brd(), bottom: brd(), left: brd(), right: brd() };
    const boxNone = { top: NONE, bottom: NONE, left: NONE, right: NONE };
    const CM = { top: 60, bottom: 60, left: 90, right: 90 };

    /* 텍스트 런 */
    const run = (t, o) => {
      o = o || {};
      return new TextRun({
        text: String(t == null ? "" : t), font: o.font || FONT,
        size: o.size || SZ.body, bold: !!o.bold, italics: !!o.italics,
        color: o.color || C.text, underline: o.underline ? {} : undefined,
        break: o.break, allCaps: false,
      });
    };

    /* 문단 */
    const P = (t, o) => {
      o = o || {};
      return new Paragraph({
        alignment: o.align, spacing: Object.assign({ before: 30, after: 30 }, LINE, o.spacing || {}),
        indent: o.indent, border: o.border, shading: o.shading,
        children: Array.isArray(t) ? t : [run(t, o)],
      });
    };
    const EMPTY = (h) => new Paragraph({ spacing: { before: 0, after: h || 60, line: 240 }, children: [run("")] });
    const BREAK = () => new Paragraph({ children: [new PageBreak()] });

    /* ── 공문서 항목기호 체계 (1. → 가. → 1) → 가) → (1)) ───────────── */
    const IND = [0, 400, 800, 1200, 1600, 2000];
    const KOR = ["가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하"];
    const mark = (lv, i) => {
      switch (lv) {
        case 1: return `${i + 1}.`;
        case 2: return `${KOR[i % 14]}.`;
        case 3: return `${i + 1})`;
        case 4: return `${KOR[i % 14]})`;
        case 5: return `(${i + 1})`;
        default: return `(${KOR[i % 14]})`;
      }
    };
    /** 공문서 본문 항목 — lv 1~6, idx 0-base */
    const item = (lv, idx, text, o) => {
      o = o || {};
      return new Paragraph({
        indent: { left: IND[Math.min(lv, 5)], hanging: 0 },
        spacing: Object.assign({ before: 40, after: 40 }, LINE),
        children: [run(`${mark(lv, idx)}  `, { bold: lv <= 2, size: o.size || SZ.body }), run(text, { size: o.size || SZ.body, bold: o.bold })],
      });
    };

    /* ── 보고자료 개조식 기호 (□ → ○ → -) ───────────────────────────── */
    const sq = (t, o) => new Paragraph({
      spacing: { before: 180, after: 60, line: 340, lineRule: "auto" },
      children: [run("□ ", { bold: true, size: SZ.body, color: C.navy }), run(t, { bold: true, size: SZ.body, color: C.navy })],
    });
    const ci = (t, o) => new Paragraph({
      indent: { left: 340 }, spacing: { before: 40, after: 40, line: 330, lineRule: "auto" },
      children: [run("○ ", { size: SZ.body }), run(t, Object.assign({ size: SZ.body }, o || {}))],
    });
    const dash = (t, o) => new Paragraph({
      indent: { left: 700 }, spacing: { before: 24, after: 24, line: 320, lineRule: "auto" },
      children: [run("- ", { size: SZ.small }), run(t, Object.assign({ size: SZ.small, color: C.gray }, o || {}))],
    });
    const note = (t) => new Paragraph({
      indent: { left: 340 }, spacing: { before: 60, after: 60, line: 300, lineRule: "auto" },
      children: [run("※ ", { size: SZ.tiny, color: C.gray }), run(t, { size: SZ.tiny, color: C.gray })],
    });

    /* ── 제목(장/절) ────────────────────────────────────────────────── */
    const H1 = (t) => new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 360, after: 160, line: 300, lineRule: "auto" },
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
      children: [run(t, { size: SZ.h1, bold: true, color: C.navy })],
    });
    const H2 = (t) => new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 100, line: 300, lineRule: "auto" },
      children: [run(t, { size: SZ.h2, bold: true, color: C.head })],
    });

    /* ── 표 ─────────────────────────────────────────────────────────── */
    const hc = (t, w, o) => {
      o = o || {};
      return new TableCell({
        borders: boxAll, width: { size: w, type: WidthType.DXA },
        shading: { fill: o.dark ? C.head : C.head2, type: ShadingType.CLEAR },
        margins: CM, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span,
        children: [new Paragraph({
          alignment: o.align || AlignmentType.CENTER, spacing: { before: 20, after: 20, line: 260, lineRule: "auto" },
          children: [run(t, { bold: true, size: o.size || SZ.small, color: o.dark ? C.white : C.navy })],
        })],
      });
    };
    const dc = (t, w, o) => {
      o = o || {};
      return new TableCell({
        borders: boxAll, width: { size: w, type: WidthType.DXA },
        shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR } : undefined,
        margins: CM, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span, rowSpan: o.rowSpan,
        children: [new Paragraph({
          alignment: o.align || AlignmentType.CENTER, spacing: { before: 20, after: 20, line: 260, lineRule: "auto" },
          children: [run(t, { size: o.size || SZ.small, bold: o.bold, color: o.color || C.text })],
        })],
      });
    };
    const table = (widths, rows, o) => new Table({
      width: { size: (o && o.width) || TW, type: WidthType.DXA },
      columnWidths: widths, rows, layout: docx.TableLayoutType ? docx.TableLayoutType.FIXED : undefined,
      alignment: AlignmentType.CENTER,
    });
    const tblCaption = (t) => new Paragraph({
      spacing: { before: 120, after: 60, line: 260, lineRule: "auto" },
      children: [run(t, { size: SZ.small, bold: true, color: C.gray })],
    });
    const tblNote = (t) => new Paragraph({
      alignment: AlignmentType.RIGHT, spacing: { before: 40, after: 120, line: 260, lineRule: "auto" },
      children: [run(t, { size: SZ.tiny, color: C.gray })],
    });

    /** 단일 셀 강조 박스 */
    const box = (children, fill) => new Table({
      width: { size: TW, type: WidthType.DXA }, columnWidths: [TW],
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: brd(8, C.navy), bottom: brd(8, C.navy), left: brd(8, C.navy), right: brd(8, C.navy) },
          width: { size: TW, type: WidthType.DXA }, margins: { top: 140, bottom: 140, left: 200, right: 200 },
          shading: { fill: fill || C.alt, type: ShadingType.CLEAR }, children,
        })],
      })],
    });

    /* ── 시행문 두문 (기관명 / 수신 / 경유 / 제목) ─────────────────── */
    function head(ds, subject) {
      const out = [];
      out.push(new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200, line: 300, lineRule: "auto" },
        children: [run(ds.senderOrg, { size: SZ.title, bold: true, color: C.navy })],
      }));
      out.push(new Paragraph({
        spacing: { before: 0, after: 200 },
        border: { bottom: { style: BorderStyle.DOUBLE, size: 8, color: C.navy } }, children: [run("")],
      }));
      const lbl = (t) => run(t, { bold: true, size: SZ.body });
      out.push(new Paragraph({
        spacing: { before: 60, after: 60, line: 320, lineRule: "auto" },
        children: [lbl("수신  "), run(ds.recipient + (ds.recipientDept ? ` (${ds.recipientDept})` : ""))],
      }));
      out.push(new Paragraph({
        spacing: { before: 60, after: 60, line: 320, lineRule: "auto" },
        children: [lbl("(경유)  "), run(ds.via || "")],
      }));
      out.push(new Paragraph({
        spacing: { before: 60, after: 200, line: 320, lineRule: "auto" },
        children: [lbl("제목  "), run(subject, { bold: true })],
      }));
      return out;
    }

    /* ── 붙임 표기 + "끝." (규정: 마지막 글자에서 2타 띄우고 끝.) ───── */
    function attachments(list) {
      const out = [EMPTY(120)];
      if (!list || !list.length) {
        out.push(new Paragraph({
          alignment: AlignmentType.RIGHT, spacing: { before: 120, after: 120, line: 320, lineRule: "auto" },
          children: [run("끝.", { bold: true })],
        }));
        return out;
      }
      list.forEach((a, i) => {
        const last = i === list.length - 1;
        out.push(new Paragraph({
          indent: { left: i === 0 ? 0 : 660 }, spacing: { before: 50, after: 50, line: 320, lineRule: "auto" },
          children: [
            i === 0 ? run("붙임  ", { bold: true }) : run(""),
            run(`${i + 1}. ${a.name} ${a.copies || "1부"}.`),
            last ? run("  끝.", { bold: true }) : run(""),
          ],
        }));
      });
      return out;
    }

    /* ── 시행문 결문 (발신명의 / 결재선 / 시행·접수 / 기관정보 / 공개) ─ */
    function foot(ds, ctx) {
      const out = [];
      out.push(new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 340, after: 220, line: 320, lineRule: "auto" },
        keepNext: true, keepLines: true,
        children: [
          run(ds.senderOrg + "  ", { size: SZ.sub, bold: true }),
          run(ds.senderName, { size: SZ.sub, bold: true }),
          run("      (직인)", { size: SZ.small, color: C.light }),
        ],
      }));
      out.push(new Paragraph({
        spacing: { before: 0, after: 50 }, keepNext: true,
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } }, children: [run("")],
      }));
      const l = (t, v) => new Paragraph({
        spacing: { before: 16, after: 16, line: 250, lineRule: "auto" }, keepLines: true,
        children: [run(t, { size: SZ.tiny, bold: true }), run(v, { size: SZ.tiny })],
      });
      const sign = (r) => `${r.title || ""} ${r.name || "(   )"}`.trim();
      out.push(l("기안자 ", `${sign(ds.drafter)}      검토자 ${sign(ds.reviewer)}      결재권자 ${sign(ds.approver)}`));
      out.push(l("협조자 ", ds.cooperator || ""));
      out.push(l("시행  ", `${ds.docNo}  (${ctx.dateKr})        접수  ${" ".repeat(10)}(          )`));
      out.push(l("우 ", `${ds.zip}  ${ds.address}   /  ${ds.homepage}`));
      out.push(new Paragraph({
        spacing: { before: 16, after: 40, line: 250, lineRule: "auto" }, keepLines: true,
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.line } },
        children: [run("전화 ", { size: SZ.tiny, bold: true }),
          run(`${ds.tel}      전송 ${ds.fax}      /  ${ds.email}      /  ${ds.openLevel}`, { size: SZ.tiny })],
      }));
      return out;
    }

    /* ── 결재란 (수기결재용, 우측 상단) ──────────────────────────────── */
    function approvalBox(ds) {
      const w = 1300, cols = [w, w, w];
      const cell = (t, sh) => new TableCell({
        borders: boxAll, width: { size: w, type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 40, right: 40 },
        shading: sh ? { fill: C.alt, type: ShadingType.CLEAR } : undefined, verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 10, after: 10, line: 240 }, children: [run(t, { size: SZ.tiny, bold: sh })] })],
      });
      const blank = () => new TableCell({
        borders: boxAll, width: { size: w, type: WidthType.DXA }, margins: { top: 200, bottom: 200, left: 40, right: 40 },
        children: [new Paragraph({ children: [run("")] })],
      });
      return new Table({
        width: { size: w * 3, type: WidthType.DXA }, columnWidths: cols,
        alignment: AlignmentType.RIGHT,
        rows: [
          new TableRow({ children: [cell("담  당", true), cell("검  토", true), cell("결  재", true)] }),
          new TableRow({ children: [cell(`${ds.drafter.title || ""} ${ds.drafter.name || ""}`.trim() || " ", false), cell(`${ds.reviewer.title || ""} ${ds.reviewer.name || ""}`.trim(), false), cell(`${ds.approver.title || ""} ${ds.approver.name || ""}`.trim() || " ", false)] }),
          new TableRow({ children: [blank(), blank(), blank()] }),
        ],
      });
    }

    /* ── 머리글 / 바닥글 ────────────────────────────────────────────── */
    const header = (t) => new Header({
      children: [new Paragraph({
        alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 40, line: 240 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.light } },
        children: [run(t, { size: SZ.tiny, color: C.gray })],
      })],
    });
    const headerBlank = () => new Header({ children: [new Paragraph({ spacing: { before: 0, after: 0, line: 240 }, children: [run("")] })] });
    const footer = () => new Footer({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 40, after: 0, line: 240 },
        children: [run("- ", { size: SZ.tiny, color: C.gray }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ.tiny, color: C.gray }),
          run(" -", { size: SZ.tiny, color: C.gray })],
      })],
    });

    const pageProps = () => ({
      page: { size: PAGE, margin: MARGIN },
    });

    const styles = () => ({
      default: { document: { run: { font: FONT, size: SZ.body, color: C.text }, paragraph: { spacing: LINE } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: SZ.h1, bold: true, font: FONT, color: C.navy }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: SZ.h2, bold: true, font: FONT, color: C.head }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
      ],
    });

    return {
      TW, C, SZ, FONT, MARGIN, PAGE,
      run, P, EMPTY, BREAK, item, mark, sq, ci, dash, note, H1, H2,
      hc, dc, table, tblCaption, tblNote, box,
      head, foot, attachments, approvalBox, header, headerBlank, footer, pageProps, styles,
      comma, money, moneyKr, hanNumber, eokStr, pct, dateKr, ymKr,
      AlignmentType,
    };
  }

  return { create, MM, PAGE, MARGIN, TW, FONT, C, SZ, comma, money, moneyKr, hanNumber, eokStr, pct, dateKr, ymKr };
});

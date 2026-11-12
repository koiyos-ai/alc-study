const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, HeadingLevel, ShadingType, BorderStyle, LevelFormat, PageBreak,
  Footer, PageNumber,
} = require('docx');

// スマホ向け：幅10.5cm×高さ18.6cm、余白0.7cm、本文11pt
const PAGE = { width: 5953, height: 10546 };
const M = 397;
const W = PAGE.width - M * 2;

const FONT = { ascii: 'Yu Gothic', eastAsia: 'Yu Gothic', hAnsi: 'Yu Gothic', cs: 'Yu Gothic' };
const C = { accent: '1F4E79', q: 'EAF1F8', a: 'EEF6EC', pt: 'FFF4E0', card: 'F5F7FA', border: 'A6B4C2' };
const BASE = 22;

const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || BASE, bold: o.bold, color: o.color });
const p = (runs, o = {}) => new Paragraph({
  children: (Array.isArray(runs) ? runs : [runs]).map(r => typeof r === 'string' ? t(r, o) : r),
  spacing: { before: o.before ?? 40, after: o.after ?? 100, line: 320 },
  alignment: o.align, keepNext: o.keepNext,
});
const bullet = (text, lvl = 0) => new Paragraph({
  children: [t(text)], numbering: { reference: 'bul', level: lvl }, spacing: { after: 80, line: 320 },
});
const num = (text, ref) => new Paragraph({
  children: [t(text)], numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 320 },
});
const h1 = s => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: s, font: FONT })] });
const h2 = s => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: s, font: FONT })] });
const h3 = s => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: s, font: FONT })] });
const brk = () => new Paragraph({ children: [new PageBreak()] });
const gap = () => p('', { before: 0, after: 60 });

const bd = { style: BorderStyle.SINGLE, size: 4, color: C.border };
const borders = { top: bd, bottom: bd, left: bd, right: bd };

// 1列だけの色付き囲み（横スクロールが出ない）
function box(label, fill, paras, keep = false) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ cantSplit: keep, children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, borders,
      shading: { fill, type: ShadingType.CLEAR, color: 'auto' },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [...(label ? [p([t(label, { bold: true, color: C.accent })], { after: 60 })] : []), ...paras],
    })] })],
  });
}
// 見出し＋行の「カード」（表の1行を縦に並べたもの）
const card = (title, lines, result) => [box(null, C.card, [
  p([t(title, { bold: true })], { after: 40 }),
  ...lines.map(l => p(l, { after: 30 })),
  ...(result ? [p([t(result, { bold: true, color: C.accent })], { before: 60, after: 0 })] : []),
]), gap()];
// 「ラベル：値」の行
const kv = (k, v) => p([t(k + '：', { bold: true }), t(v)], { after: 40 });

const body = [];
const add = (...xs) => body.push(...xs);

// ===== Markdown（見出し・箇条書き・段落・区切り線のみ）をスマホ版 Word にする =====
// 使い方: node tools/build_md_mobile.js <入力.md> <出力.docx>
const lines = fs.readFileSync(process.argv[2], 'utf-8').split('\n');
let title = '';
const inline = s => {
  // `code` と **太字** を処理
  const runs = [];
  s.replace(/`([^`]+)`/g, '$1').split(/(\*\*[^*]+\*\*)/).forEach(part => {
    if (!part) return;
    if (part.startsWith('**')) runs.push(t(part.slice(2, -2), { bold: true }));
    else runs.push(t(part));
  });
  return runs;
};
let firstH1 = true;
// 表は1行ずつ「カード」にする（1列目を見出し、他の列は「見出し：値」）。<br> は改行
const cells = l => l.replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim());
let tblHead = null;
for (const raw of lines) {
  const line = raw.replace(/\s+$/, '');
  if (line.startsWith('|')) {
    const c = cells(line);
    if (c.every(x => /^:?-+:?$/.test(x))) continue;
    if (!tblHead) { tblHead = c; continue; }
    const paras = [];
    if (c[0]) paras.push(p(inline(c[0]).map(r => r), { after: 40 }));
    c.slice(1).forEach((v, k) => {
      if (!v) return;
      const parts = v.split(/<br\s*\/?>/);
      const hd = tblHead[k + 1] || '';
      const last = k === c.length - 2;
      if (parts.length === 1 && last) paras.push(p([t(hd ? hd + '：' : '', { bold: true }), t(v.replace(/\*\*/g, ''), { bold: true, color: C.accent })], { after: 20 }));
      else { if (hd) paras.push(p([t(hd, { bold: true })], { after: 20 })); parts.forEach(x => paras.push(p(inline(x), { after: 20 }))); }
    });
    if (!c[0] && paras.length) paras.unshift(p([t('合計', { bold: true })], { after: 20 }));
    add(box(null, C.card, paras, true), gap());
    continue;
  }
  tblHead = null;
  if (!line) continue;
  if (line.startsWith('>')) { const x = line.replace(/^>\s?/, ''); if (x) add(box(null, C.pt, [p(inline(x), { after: 0 })], true), gap()); continue; }
  if (line === '---') { add(brk()); continue; }
  let m;
  if ((m = line.match(/^# (.+)/))) {
    if (!title) { title = m[1];
      add(p('', { before: 1000 }), p([t(m[1], { size: 36, bold: true, color: C.accent })], { align: AlignmentType.CENTER, after: 400 }));
      continue; }
    add(h1(m[1])); continue;
  }
  if ((m = line.match(/^## (.+)/))) { add(h2(m[1])); continue; }
  if ((m = line.match(/^### (.+)/))) { add(h3(m[1])); continue; }
  if ((m = line.match(/^(\s*)- (.+)/))) {
    add(new Paragraph({ children: inline(m[2]), numbering: { reference: 'bul', level: m[1].length >= 2 ? 1 : 0 }, spacing: { after: 80, line: 320 } }));
    continue;
  }
  if ((m = line.match(/^(\d+)\. (.+)/))) { add(p([t(m[1] + '. ', { bold: true }), ...inline(m[2])])); continue; }
  add(p(inline(line)));
}

const doc = new Document({
  creator: 'alc-study', title,
  styles: {
    default: { document: { run: { font: FONT, size: BASE } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: C.accent },
        paragraph: { spacing: { before: 120, after: 140 }, keepNext: true, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.accent, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 26, bold: true, color: C.accent },
        paragraph: { spacing: { before: 200, after: 100 }, keepNext: true, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 24, bold: true },
        paragraph: { spacing: { before: 160, after: 80 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [{ reference: 'bul', levels: [
    { level: 0, format: LevelFormat.BULLET, text: '●', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 300, hanging: 240 } }, run: { size: 14 } } },
    { level: 1, format: LevelFormat.BULLET, text: '・', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 560, hanging: 240 } } } },
  ] }] },
  sections: [{
    properties: { page: { size: PAGE, margin: { top: M, bottom: M + 100, left: M, right: M, footer: 200 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ children: [PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], font: FONT, size: 14, color: '808080' }),
    ] })] }) },
    children: body,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(process.argv[3], buf); console.log('written', process.argv[3]); });

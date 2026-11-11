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

// ===== 大問1（種類・品目の判定）学習用 =====
// 使い方: node tools/build_dai1.js <extract_dai1.py が出力したJSON> <出力docx>
const data = JSON.parse(fs.readFileSync(process.argv[2], 'utf-8'));

// 各酒類の判定の決め手（復習用）
const POINTS = {
  'R4イ': ['発芽玄米は清酒の原料（米）に当たらない → 清酒ではない。', '穀類・糖類を発酵、20度未満・エキス2度以上 → その他の醸造酒。', '酒類（焼酎）を加えるとその他の醸造酒でなくなる → エキス2度以上でリキュール。', '発泡性があり10度未満 → 種類は発泡性酒類。'],
  'R4ロ': ['政令物品（焼酎・水あめ・有機酸）の合計 580kg ＝ 米（こうじ米を含む）の50%。', '「超えない」なので、ちょうど50%は清酒になる。'],
  'R4ハ': ['濃縮果汁も「果実」。果実が原料なら連続式蒸留機でも焼酎にならない。', '原酒のアルコール分 270ℓ ≧ 総量 670ℓ の10%（67ℓ）→ ブランデーのまま。'],
  'R4ニ': ['砂糖は転化糖に換算（×1.05）してから果実の糖類と比べる：89.25kg ≦ 90kg。', 'ブランデー 28ℓ ≦ 総量 324ℓ の10%（32.4ℓ）→ 果実酒のまま。'],
  'R4ホ': ['麦芽 1,000kg ≧ ホップ・水以外 1,200kg の50%（600kg）。', '果実・香味料（りんご果汁＋コリアンダー）50kg ≦ 麦芽の5%（50kg）。'],
  'R5イ': ['政令物品（焼酎225＋ぶどう糖35＋有機酸30）290kg ＝ 米（300＋こうじ米280）の50%。', 'R4ロと同じ型。ちょうど50%は清酒。'],
  'R5ロ': ['最初の酒類：麦芽 1,000kg ≧ 1,700kg の50%、果実・香味料なし → ビール。', '加えて発酵した後も同じ基準で判定（法3条12号ハ）：果実・香味料 50kg ≦ 麦芽の5%（50kg）。'],
  'R5ハ': ['果実酒を蒸留したものもブランデー（法3条16号イ）。', '原酒 150ℓ ≧ 総量 1,350ℓ の10%（135ℓ）。カラメルはブランデーに加えてよい色素。'],
  'R5ニ': ['発芽玄米 → 清酒ではない → その他の醸造酒（R4イと同じ）。', 'スピリッツ（アルコール以外の酒類）を加えるとその他の醸造酒でなくなる（令8条1号）。', 'エキス分ちょうど2度は「2度以上」なのでリキュール（2度未満ならスピリッツ）。'],
  'R5ホ': ['果糖・ぶどう糖は換算係数1（砂糖だけ1.05）：100kg ≦ 100kg。', 'ブランデー 12ℓ ≦ 総量 142ℓ の10%（14.2ℓ）。'],
  'R6イ': ['発芽させた穀類と水で穀類（ライ麦）を糖化 → ウイスキー（法3条15号ロ）。連続式蒸留機でも焼酎にならない。', '原料用アルコールは「アルコール」として加えられる：原酒 360ℓ ≧ 総量 2,760ℓ の10%（276ℓ）。'],
  'R6ロ': ['政令物品 360kg ≦ 米の50%（580kg）→ 清酒。', '清酒に清酒かすを加えて、こしたものも清酒（法3条7号ハ）。'],
  'R6ハ': ['砂糖80×1.05＝84kg ≦ 85kg → まず果実酒。', 'ブランデー 350ℓ が総量 560ℓ の10%（56ℓ）を超える → 果実酒でない。', '90%（504ℓ）以下 → 甘味果実酒（混成酒類）。'],
  'R6ニ': ['小麦粉は清酒の原料として認められていない → 清酒ではない。', 'アルコールを原料にして15度以上 → その他の醸造酒から除かれる（令8条2号）。', '酒類＋物品でエキス2度以上 → リキュール。比重はアルコール重量30%基準の計算用（今回は基準未満）。'],
  'R6ホ': ['加えて発酵した後の麦芽比率がちょうど50%（700kg／1,400kg）→ ビールのまま。', '果実の合計 30kg ≦ 麦芽の5%（35kg）。オレンジピールは果実として扱った。'],
  'R7イ': ['糖蜜（含糖質物）原料でも留出時95度ちょうどなら焼酎から除かれない → 連続式蒸留焼酎。', '芋の酒は48度（36度以上）→ 原料用アルコール。', '原料用アルコール＋連続式蒸留焼酎で36度未満 → 連続式蒸留焼酎を製造したとみなす（法43条6項）。'],
  'R7ロ': ['砂糖90×1.05＝94.5kg ≦ 100kg → 果実酒。', 'ブランデー 48ℓ ≦ 総量 483ℓ の10%（48.3ℓ）。僅差に注意。'],
  'R7ハ': ['発芽玄米 → その他の醸造酒。', '炭酸水でエキス分が2度未満（1.8度）になる → スピリッツを製造したとみなす（令50条11項）。', '発泡性・10度未満 → 種類は発泡性酒類。'],
  'R7ニ': ['果汁も果実。単式蒸留でも留出95度未満ならブランデー。', '原酒 2,000ℓ ≧ 総量 5,600ℓ の10%（560ℓ）。'],
  'R7ホ': ['蜂蜜（味付け目的）は香味料＝政令物品。', '加えた後：麦芽 1,400kg ＜ 2,804kg の50%（1,402kg）、果実・香味料 74kg ＞ 麦芽の5%（70kg）→ ビールでない。', '麦芽を原料の一部とする発泡性の酒類 → 発泡酒。'],
};

// ===== 表紙・使い方 =====
add(
  p('', { before: 1000 }),
  p([t('酒税法', { size: 40, bold: true })], { align: AlignmentType.CENTER, after: 0 }),
  p([t('大問1 種類・品目の判定', { size: 32, bold: true, color: C.accent })], { align: AlignmentType.CENTER, after: 160 }),
  p([t('R4〜R7 本科二次試験', { size: 24 })], { align: AlignmentType.CENTER, after: 0 }),
  p([t('年度別・スマホ版', { size: 24 })], { align: AlignmentType.CENTER, after: 600 }),
  brk(),
  h1('この資料の使い方'),
  bullet('R4〜R7の大問1（酒類の種類・品目の判定）を年度ごとにまとめました。'),
  bullet('各酒類は「問題」→「答え」→「判定理由」→「決め手」の順です。まず問題だけを見て、種類・品目と理由を自分で考えてから確認してください。'),
  bullet('【公式】はR4の公式解答例、【自作】は講本・条文・解釈通達だけを根拠に作成したもの（公式解答例がないため正解かどうかは未確認）です。'),
  bullet('最初の「判定の基準まとめ」で、毎年使う数値基準を確認できます。'),
  brk(),
);

// ===== 判定の基準まとめ =====
add(
  h1('判定の基準まとめ'),
  p('毎年の大問1で繰り返し使う基準です（法＝酒税法、令＝施行令）。'),
  box('清酒（法3条7号）', C.card, [
    bullet('米・米こうじ・水を原料に発酵させて、こしたもの。22度未満。'),
    bullet('政令物品（令2条：アルコール、焼酎、ぶどう糖その他の糖類、有機酸、アミノ酸塩、清酒）の重量合計が、米（こうじ米を含む）の重量の50%を超えないこと。ちょうど50%は可。'),
    bullet('発芽玄米・小麦粉は清酒の原料にならない（R4公式解答例）。'),
    bullet('清酒に清酒かすを加えて、こしたものも清酒。'),
  ], true), gap(),
  box('果実酒・甘味果実酒（法3条13号・14号、令7条）', C.card, [
    bullet('加えた糖類（転化糖換算：砂糖×1.05、ぶどう糖・果糖×1）が果実に含まれる糖類を超えないこと。'),
    bullet('糖類やブランデー等を加えたものは15度未満であること。'),
    bullet('ブランデー等のアルコール分が、加えた後の総量の10%以下なら果実酒。10%を超え90%以下なら甘味果実酒。'),
  ], true), gap(),
  box('ウイスキー・ブランデー（法3条15号・16号）', C.card, [
    bullet('留出時のアルコール分95度未満。'),
    bullet('原酒のアルコール分が、アルコール・スピリッツ・香味料を加えた後の総量の10%以上。'),
    bullet('加えられる色素はカラメルだけ（通達）。'),
    bullet('発芽させた穀類・果実（濃縮果汁を含む）を原料にしたものは連続式蒸留焼酎にならない（法3条9号イ）。'),
  ], true), gap(),
  box('焼酎・原料用アルコール（法3条9号・10号・17号）', C.card, [
    bullet('連続式蒸留焼酎：36度未満。含糖質物（糖蜜など）原料で留出95度未満のものは除く。'),
    bullet('単式蒸留焼酎：45度以下。'),
    bullet('原料用アルコール：焼酎の要件（度数以外）を満たし45度超。'),
    bullet('連続式の原料用アルコール＋連続式蒸留焼酎で36度未満 → 連続式蒸留焼酎とみなす（法43条6項）。'),
  ], true), gap(),
  box('ビール・発泡酒（法3条12号・18号、令6条）', C.card, [
    bullet('ビール：麦芽がホップ・水以外の原料の50%以上、果実・香味料（令6条1項2号）が麦芽の5%以下、20度未満。'),
    bullet('ビールに物品を加えて発酵させた場合も、加えた後の原料全体で同じ基準を満たすこと（法3条12号ハ）。'),
    bullet('ビールに当たらない麦芽・麦原料の発泡性の酒類は発泡酒（R6以降は定義が拡大）。'),
  ], true), gap(),
  box('その他の醸造酒・リキュール・スピリッツ', C.card, [
    bullet('その他の醸造酒：穀類・糖類等を発酵、20度未満、エキス2度以上（法3条19号）。'),
    bullet('アルコール以外の酒類を原料にしたもの、アルコールを原料にして15度以上又は重量30%以上のものは除く（令8条）。'),
    bullet('その他の醸造酒に水・炭酸水を加えてエキス2度未満 → スピリッツ（令50条11項）。'),
    bullet('リキュール：酒類＋糖類その他の物品でエキス2度以上。スピリッツ：エキス2度未満。'),
  ], true), gap(),
  box('種類の判定と炭酸ガス', C.card, [
    bullet('ビール・発泡酒以外で発泡性があり、低アルコール（条文は11度未満。R4公式解答例は「10度未満」と記述）→ その他の発泡性酒類（種類は発泡性酒類）。'),
    bullet('炭酸ガス（炭酸水）を加えても、品目は原則として混和前のまま（法43条2項）。'),
  ], true), gap(),
  box('答案の書き方（R4公式解答例の型）', C.pt, [
    bullet('原料・製法を書く → 要件を数値で示す（重量・アルコール総量の計算式）→「〇〇に該当する」。'),
    bullet('混和した段階ごとに判定を書く。'),
    bullet('最後に「〇〇の種類は、△△である」で締める。'),
  ], true),
  brk(),
);

// ===== 年度別 =====
data.forEach((y, yi) => {
  const tag = y.src === 'official' ? '公式' : '自作';
  add(h1(`${y.key}（${y.label}）`));
  add(p(y.lead));
  add(h3('答え一覧'));
  y.items.forEach(it => add(...card(`${it.kana}　${it.kind}／${it.item}`, [], null)));
  add(brk());
  y.items.forEach((it, ii) => {
    add(
      h2(`${y.key}　${it.kana}`),
      box('問題', C.q, [p(it.problem)]),
      gap(),
      box(`答え【${tag}】`, C.a, [p([t(`種類：${it.kind}`, { bold: true })], { after: 20 }), p([t(`品目：${it.item}`, { bold: true })], { after: 0 })], true),
      gap(),
      box('判定理由', C.a, it.reasons.map(r => p(r))),
      gap(),
      box('決め手', C.pt, [...(POINTS[y.key + it.kana] || []).map(s => bullet(s)), ...it.notes.map(n => bullet('注：' + n))], true),
    );
    if (!(yi === data.length - 1 && ii === y.items.length - 1)) add(brk());
  });
});

const doc = new Document({
  creator: 'alc-study',
  title: '酒税法 大問1 種類・品目の判定（スマホ版）',
  styles: {
    default: { document: { run: { font: FONT, size: BASE } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: C.accent },
        paragraph: { spacing: { before: 120, after: 140 }, keepNext: true, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.accent, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 26, bold: true, color: C.accent },
        paragraph: { spacing: { before: 120, after: 100 }, keepNext: true, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 24, bold: true },
        paragraph: { spacing: { before: 160, after: 80 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: 'bul', levels: [
        { level: 0, format: LevelFormat.BULLET, text: '●', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 300, hanging: 240 } }, run: { size: 14 } } },
        { level: 1, format: LevelFormat.BULLET, text: '・', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 560, hanging: 240 } } } },
      ] },
    ],
  },
  sections: [{
    properties: { page: { size: PAGE, margin: { top: M, bottom: M + 100, left: M, right: M, footer: 200 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ children: [PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], font: FONT, size: 14, color: '808080' }),
    ] })] }) },
    children: body,
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(process.argv[3], buf); console.log('written', process.argv[3]); });

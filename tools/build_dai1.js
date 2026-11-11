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
], true), gap()];
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

// ===== 品目別の並び =====
// 品目ごとに、判定基準と、その品目が答えになった問題（年度順）をまとめる
const GROUPS = [
  { item: '清酒', cat: '醸造酒類', rules: [
    '米・米こうじ・水を原料に発酵させて、こしたもの。22度未満（法3条7号）。',
    '政令物品（令2条：アルコール、焼酎、ぶどう糖その他の糖類、有機酸、アミノ酸塩、清酒）の重量合計が、米（こうじ米を含む）の重量の50%を超えないこと。ちょうど50%は可。',
    '発芽玄米・小麦粉は清酒の原料にならない（R4公式解答例）。',
    '清酒に清酒かすを加えて、こしたものも清酒（法3条7号ハ）。'] },
  { item: '果実酒', cat: '醸造酒類', rules: [
    '果実（濃縮果汁・果汁を含む）と水に糖類を加えて発酵させたもの（法3条13号）。',
    '加えた糖類（転化糖換算：砂糖×1.05、ぶどう糖・果糖×1）が果実に含まれる糖類を超えないこと（令7条1項1号）。',
    '糖類やブランデー等を加えたものは15度未満であること。',
    'ブランデー等のアルコール分が、加えた後の総量の10%以下であること（法3条13号ニ）。'] },
  { item: '甘味果実酒', cat: '混成酒類', rules: [
    '果実酒の要件を満たさないもののうち、果実酒等にブランデー等を加えたもので、ブランデー等のアルコール分が加えた後の総量の90%以下のもの（法3条14号ハ）。',
    '果実酒との境目：ブランデー等が10%以下なら果実酒、10%を超え90%以下なら甘味果実酒。'] },
  { item: '連続式蒸留焼酎', cat: '蒸留酒類', rules: [
    'アルコール含有物を連続式蒸留機で蒸留したもの。36度未満（法3条9号）。',
    '発芽させた穀類・果実を原料にしたもの、含糖質物（糖蜜など）原料で留出95度未満のものは除く。',
    '45度を超えると原料用アルコール（法3条17号）。',
    '連続式の原料用アルコール＋連続式蒸留焼酎で36度未満 → 連続式蒸留焼酎とみなす（法43条6項）。'] },
  { item: 'ウイスキー', cat: '蒸留酒類', rules: [
    '発芽させた穀類と水（で穀類を糖化させて）発酵させたアルコール含有物を蒸留。留出時95度未満（法3条15号イ・ロ）。',
    '原酒のアルコール分が、アルコール・スピリッツ・香味料を加えた後の総量の10%以上（法3条15号ハ）。',
    '発芽させた穀類が原料なので、連続式蒸留機でも焼酎にならない。'] },
  { item: 'ブランデー', cat: '蒸留酒類', rules: [
    '果実（濃縮果汁・果汁を含む）を原料として発酵させたアルコール含有物又は果実酒を蒸留。留出時95度未満（法3条16号イ）。',
    '原酒のアルコール分が、アルコール・スピリッツ・香味料を加えた後の総量の10%以上（法3条16号ロ）。',
    '加えられる色素はカラメルだけ（通達）。',
    '果実が原料なので、連続式蒸留機でも焼酎にならない（法3条9号イ）。'] },
  { item: 'リキュール', cat: '混成酒類（発泡性があれば発泡性酒類）', rules: [
    '酒類と糖類その他の物品（酒類を含む）を原料とした酒類で、エキス分2度以上（法3条21号）。',
    'その他の醸造酒に酒類（焼酎・スピリッツ）を加えると、その他の醸造酒でなくなる（令8条1号）。アルコールを原料にして15度以上のものも除かれる（令8条2号）。',
    'エキス分ちょうど2度はリキュール（2度未満ならスピリッツ）。',
    '発泡性があり低アルコールなら、種類は発泡性酒類（その他の発泡性酒類）。'] },
  { item: 'スピリッツ', cat: '蒸留酒類（発泡性があれば発泡性酒類）', rules: [
    '他の品目に当たらない酒類で、エキス分2度未満（法3条20号）。',
    'その他の醸造酒に水・炭酸水を加えてエキス分2度未満 → スピリッツを製造したとみなす（令50条11項）。'] },
  { item: 'ビール', cat: '発泡性酒類', rules: [
    '麦芽・ホップ・水（＋麦、米、とうもろこし、ばれいしょ、でん粉、糖類、果実、香味料など）を原料に発酵。20度未満（法3条12号、令6条）。',
    '麦芽がホップ・水以外の原料の50%以上。',
    '果実・香味料（令6条1項2号：果実、コリアンダー、蜂蜜など）の合計が麦芽の5%以下。',
    'ビールに物品を加えて発酵させた場合も、加えた後の原料全体で同じ基準を満たすこと（法3条12号ハ）。'] },
  { item: '発泡酒', cat: '発泡性酒類', rules: [
    '麦芽又は麦を原料の一部とした発泡性の酒類で、ビール等に当たらないもの。20度未満（法3条18号）。',
    'ビールの基準（麦芽50%以上・果実香味料5%以下）を一つでも満たさなければ、発泡酒になる。',
    'R6（令和5年10月改正）以降は、ホップ等を原料としたものなどにも定義が拡大。'] },
];

const ORDER = ['R4', 'R5', 'R6', 'R7'];
const all = [];
data.forEach(y => y.items.forEach(it => all.push({ ...it, year: y.key, label: y.label, src: y.src })));

// ===== 表紙・使い方 =====
add(
  p('', { before: 1000 }),
  p([t('酒税法', { size: 40, bold: true })], { align: AlignmentType.CENTER, after: 0 }),
  p([t('大問1 種類・品目の判定', { size: 32, bold: true, color: C.accent })], { align: AlignmentType.CENTER, after: 160 }),
  p([t('R4〜R7 本科二次試験', { size: 24 })], { align: AlignmentType.CENTER, after: 0 }),
  p([t('品目別・スマホ版', { size: 24 })], { align: AlignmentType.CENTER, after: 600 }),
  brk(),
  h1('この資料の使い方'),
  bullet('R4〜R7の大問1（酒類の種類・品目の判定）20問を、答えの品目ごとに並べ直しました。'),
  bullet('各品目の最初に、その品目の判定基準をまとめています。続けて、その品目が答えになった問題を年度順に並べています。'),
  bullet('各問は「問題」→「答え」→「判定理由」→「決め手」の順です。まず問題だけを見て、自分で考えてから確認してください。'),
  bullet('【公式】はR4の公式解答例、【自作】は講本・条文・解釈通達だけを根拠に作成したもの（公式解答例がないため正解かどうかは未確認）です。'),
  bullet('条文番号の「法」は酒税法、「令」は酒税法施行令です。'),
  brk(),
);

// ===== 目次（品目別の出題一覧） =====
add(h1('品目別の出題一覧'), p('どの品目が何回答えになったかの一覧です。'));
GROUPS.forEach(g => {
  const hits = all.filter(x => x.item === g.item).sort((a, b) => ORDER.indexOf(a.year) - ORDER.indexOf(b.year));
  add(...card(`${g.item}（${hits.length}問）`, hits.map(x => `${x.year} ${x.kana}　種類：${x.kind}`)));
});
const known = new Set(GROUPS.map(g => g.item));
const missing = all.filter(x => !known.has(x.item));
if (missing.length) throw new Error('品目グループ未定義: ' + missing.map(x => x.year + x.kana + x.item).join(', '));

add(
  h2('種類の判定（全品目共通）'),
  box(null, C.card, [
    bullet('醸造酒類：清酒、果実酒、その他の醸造酒。蒸留酒類：焼酎、ウイスキー、ブランデー、原料用アルコール、スピリッツ。混成酒類：合成清酒、みりん、甘味果実酒、リキュールなど。発泡性酒類：ビール、発泡酒、その他の発泡性酒類（法3条3号〜6号）。'),
    bullet('ビール・発泡酒以外で発泡性があり低アルコール（条文は11度未満。R4公式解答例は「10度未満」と記述）→ その他の発泡性酒類として、種類は発泡性酒類。品目はそのまま。'),
    bullet('炭酸ガス（炭酸水）を加えても、品目は原則として混和前のまま（法43条2項）。'),
  ], true), gap(),
  box('答案の書き方（R4公式解答例の型）', C.pt, [
    bullet('原料・製法を書く → 要件を数値で示す（重量・アルコール総量の計算式）→「〇〇に該当する」。'),
    bullet('混和した段階ごとに判定を書く。'),
    bullet('最後に「〇〇の種類は、△△である」で締める。'),
  ], true),
  brk(),
);

// ===== 品目ごとの本編 =====
GROUPS.forEach((g, gi) => {
  const hits = all.filter(x => x.item === g.item).sort((a, b) => ORDER.indexOf(a.year) - ORDER.indexOf(b.year));
  add(
    h1(`${g.item}（${hits.length}問）`),
    p([t('種類：', { bold: true }), t(g.cat)]),
    box(`${g.item}の判定基準`, C.card, g.rules.map(r => bullet(r)), true),
    brk(),
  );
  hits.forEach((it, ii) => {
    const tag = it.src === 'official' ? '公式' : '自作';
    add(
      h2(`${it.year} ${it.kana}　${g.item}`),
      box(`問題（${it.label}）`, C.q, [p(it.problem)]),
      gap(),
      box(`答え【${tag}】`, C.a, [p([t(`種類：${it.kind}`, { bold: true })], { after: 20 }), p([t(`品目：${it.item}`, { bold: true })], { after: 0 })], true),
      gap(),
      box('判定理由', C.a, it.reasons.map(r => p(r))),
      gap(),
      box('決め手', C.pt, [...(POINTS[it.year + it.kana] || []).map(s => bullet(s)), ...it.notes.map(n => bullet('注：' + n))], true),
    );
    if (!(gi === GROUPS.length - 1 && ii === hits.length - 1)) add(brk());
  });
});

const doc = new Document({
  creator: 'alc-study',
  title: '酒税法 大問1 種類・品目の判定（品目別・スマホ版）',
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

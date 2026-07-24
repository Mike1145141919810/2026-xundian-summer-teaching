/**
 * Build Li Fuzhou's teaching plan docx — 《把大象塞进冰箱》
 *
 * Usage:
 *   node build_03_doc.js
 * Outputs: 03-把大象塞进冰箱-李福洲.docx in the same directory
 */

'use strict';

const fs   = require('fs');
const path = require('path');

const {
  Document, Packer,
  Paragraph, TextRun,
  Table, TableRow, TableCell,
  Header, Footer,
  PageNumber, AlignmentType, LineRuleType, HeadingLevel,
  LevelFormat, LevelSuffix, BorderStyle, WidthType, ShadingType,
  TableOfContents, PageBreak,
} = require('docx');

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const OUTPUT_PATH   = path.join(__dirname, '03-把大象塞进冰箱-李福洲.docx');

// Page / margin
const PAGE_W        = 11906;   // A4
const PAGE_H        = 16838;
const MARGIN        = 1418;    // 2.5 cm
const CONTENT_W     = PAGE_W - 2 * MARGIN;  // 9070 DXA

// Table border presets
const THICK = { style: BorderStyle.SINGLE, size: 12, color: '000000' };
const THIN  = { style: BorderStyle.SINGLE, size: 6,  color: '000000' };
const NONE  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };

// Chapter counter
let _chapter = 0;

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function body(text) {
  return new Paragraph({
    children: [new TextRun(text)],
  });
}

function bodyBold(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true })],
  });
}

function bodyMulti(runs) {
  return new Paragraph({ children: runs });
}

function blank() {
  return new Paragraph({ children: [] });
}

/** H1 — Chapter heading with Chinese numeral */
function h1Chinese(text) {
  _chapter++;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

/** H2 — Section heading with Word auto-numbering */
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    numbering: { reference: `sections_c${_chapter}`, level: 0 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

/** H3 — Subsection heading with Word auto-numbering */
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    numbering: { reference: `sections_c${_chapter}`, level: 1 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

// Code / monospace block paragraph
function codeLine(text) {
  return new Paragraph({
    indent: { left: 480, firstLine: 0 },
    spacing: { line: 240, lineRule: LineRuleType.AUTO },
    children: [new TextRun({ text, font: { ascii: 'Consolas', eastAsia: 'SimSun', hAnsi: 'Consolas' }, size: 21 })],
  });
}

// Indented note paragraph
function indentBody(text) {
  return new Paragraph({
    indent: { left: 480, firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function indentBodyBold(text) {
  return new Paragraph({
    indent: { left: 480, firstLine: 0 },
    children: [new TextRun({ text, bold: true })],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Three-line table
// ─────────────────────────────────────────────────────────────────────────────

function threeLineTable(headers, rows, colWidths) {
  const n = headers.length;
  if (!colWidths) {
    const w = Math.floor(CONTENT_W / n);
    colWidths = Array(n).fill(w);
    colWidths[n - 1] = CONTENT_W - w * (n - 1);
  }

  const cellOf = (text, w, borders, bold = false) => {
    return new TableCell({
      width:   { size: w, type: WidthType.DXA },
      borders,
      shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        indent:    { firstLine: 0 },
        children:  [new TextRun({ text, bold })],
      })],
    });
  };

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      cellOf(h, colWidths[i], { top: THICK, bottom: THIN, left: NONE, right: NONE }, true)
    ),
  });

  const bodyRows = rows.map((row, ri) => {
    const isLast = ri === rows.length - 1;
    return new TableRow({
      children: row.map((cell, i) =>
        cellOf(String(cell), colWidths[i], {
          top: NONE, bottom: isLast ? THICK : NONE, left: NONE, right: NONE,
        })
      ),
    });
  });

  return new Table({
    width:        { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: colWidths,
    rows:         [headerRow, ...bodyRows],
  });
}

// Borderless info table (for course info block)
function borderlessTable(rows, colWidths) {
  const cellOf = (text, w, bold = false) => {
    return new TableCell({
      width:   { size: w, type: WidthType.DXA },
      borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
      children: [new Paragraph({
        indent: { firstLine: 0 },
        children: [new TextRun({ text, bold, size: 22 })],
      })],
    });
  };

  return new Table({
    width:        { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: rows.map((row, ri) =>
      new TableRow({
        children: row.map((cell, i) => cellOf(cell.text, colWidths[i], cell.bold || false)),
      })
    ),
  });
}

// Bullet list item
function bulletItem(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    children: [new TextRun(text)],
  });
}

function bulletItemBoldPrefix(prefix, text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    children: [
      new TextRun({ text: prefix, bold: true }),
      new TextRun(text),
    ],
  });
}

// Numbered list item (for teaching suggestions, etc.)
function numberItem(text) {
  return new Paragraph({
    numbering: { reference: 'numbers', level: 0 },
    children: [new TextRun(text)],
  });
}

function numberItemBoldPrefix(prefix, text) {
  return new Paragraph({
    numbering: { reference: 'numbers', level: 0 },
    children: [
      new TextRun({ text: prefix, bold: true }),
      new TextRun(text),
    ],
  });
}

// page break
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ─────────────────────────────────────────────────────────────────────────────
// Numbering setup — 7 chapters for this teaching plan
// ─────────────────────────────────────────────────────────────────────────────

function buildNumberingConfig(chapterCount) {
  const configs = [
    {
      reference: 'references',
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '[%1]',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 480 } } } },
      ],
    },
    {
      reference: 'bullets',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      ],
    },
    {
      reference: 'numbers',
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '%1.',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      ],
    },
  ];

  for (let c = 1; c <= chapterCount; c++) {
    configs.push({
      reference: `sections_c${c}`,
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: `${c}.%1`,
          suffix: LevelSuffix.SPACE, alignment: AlignmentType.LEFT },
        { level: 1, format: LevelFormat.DECIMAL, text: `${c}.%1.%2`,
          suffix: LevelSuffix.SPACE, alignment: AlignmentType.LEFT },
      ],
    });
  }

  return { config: configs };
}

const NUMBERING = buildNumberingConfig(7);

// ─────────────────────────────────────────────────────────────────────────────
// Styles
// ─────────────────────────────────────────────────────────────────────────────

const STYLES = {
  default: {
    document: {
      run: {
        font: { ascii: 'Cambria Math', hAnsi: 'Cambria Math', eastAsia: 'SimSun' },
        size: 24,  // 12pt
      },
      paragraph: {
        spacing: { line: 300, lineRule: LineRuleType.AUTO },  // ~1.25x for readability
        indent:  { firstLine: 480 },
      },
    },
  },
  paragraphStyles: [
    {
      id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' }, size: 32, bold: true },
      paragraph: {
        alignment: AlignmentType.LEFT,
        indent:    { firstLine: 0 },
        spacing:   { before: 240, after: 120, line: 360, lineRule: LineRuleType.AUTO },
        outlineLevel: 0,
      },
    },
    {
      id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' }, size: 28, bold: true },
      paragraph: {
        alignment: AlignmentType.LEFT,
        indent:    { firstLine: 0 },
        spacing:   { before: 180, after: 60, line: 360, lineRule: LineRuleType.AUTO },
        outlineLevel: 1,
      },
    },
    {
      id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' }, size: 24, bold: true },
      paragraph: {
        alignment: AlignmentType.LEFT,
        indent:    { firstLine: 0 },
        spacing:   { before: 120, after: 60, line: 300, lineRule: LineRuleType.AUTO },
        outlineLevel: 2,
      },
    },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// ██  CONTENT — Teaching Plan  ████████████████████████████████████████████████
// ─────────────────────────────────────────────────────────────────────────────

const CONTENT = [

  // ── Title ──────────────────────────────────────────────────────────────────
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing:   { before: 0, after: 360 },
    indent:    { firstLine: 0 },
    children:  [new TextRun({
      text: '教案：《把大象塞进冰箱》',
      bold: true, size: 36,
      font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' },
    })],
  }),

  // ── Course Info ────────────────────────────────────────────────────────────
  borderlessTable([
    [{ text: '课程系列', bold: true }, { text: '数字魔法师 · 第3讲' }],
    [{ text: '授课教师', bold: true }, { text: '李福洲（计算机学院）' }],
    [{ text: '授课时间', bold: true }, { text: '7月30日 9:00-9:50（50分钟）' }],
    [{ text: '授课对象', bold: true }, { text: '寻甸县雷锋希望小学学生（3-6年级混龄）' }],
    [{ text: '课程类型', bold: true }, { text: '信息与编码启蒙 · 互动游戏课' }],
  ], [1500, 7570]),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 一、教学目标
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('一、教学目标'),
  blank(),
  threeLineTable(
    ['维度', '目标'],
    [
      ['认知', '理解数据压缩的基本思想：去掉冗余、用更少的符号表达同样的信息'],
      ['技能', '能对重复字符序列使用游程编码进行压缩；能判断一段信息"压缩后是变大还是变小"'],
      ['情感', '感受"用聪明方法节省空间"的乐趣；理解"压缩"在日常生活中的普遍性'],
    ],
    [1500, 7570]
  ),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 二、教学重难点
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('二、教学重难点'),
  blank(),
  bulletItemBoldPrefix('重点：', '游程编码（RLE）的基本操作——数连续重复字符的个数'),
  bulletItemBoldPrefix('难点：', '理解"不是所有东西都能压缩"——没有规律的随机数据反而更"大"'),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 三、教学准备
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('三、教学准备'),
  blank(),
  threeLineTable(
    ['物品', '数量', '说明'],
    [
      ['大白纸/黑板', '1块', '演示用'],
      ['空白纸条', '每人3-4张', '压缩练习用'],
      ['大号塑料袋', '1个', '演示"物理压缩"'],
      ['羽绒服/蓬松衣物', '1件', '演示用——塞进塑料袋的过程'],
      ['预制的"待压缩消息"', '4条', '见附录'],
    ],
    [2500, 1500, 5070]
  ),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 四、教学过程
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('四、教学过程'),
  blank(),

  // ── 环节一 ──
  h2('环节一：情景导入——"大象怎么塞进冰箱？"（5分钟）'),
  blank(),

  bodyBold('教师活动：'),
  numberItem('先不讲压缩，拿起一件羽绒服，问学生："这件衣服体积很大，怎么让它变小？"'),
  indentBody('学生自然会说："塞进袋子！用力压！"'),
  numberItem('现场演示：把羽绒服塞进塑料袋，挤出空气。'),
  numberItem('提问："变小的衣服和原来有什么不同？还能穿吗？"'),
  indentBody('引导：衣服还是那件衣服，只是去掉了"空气"。'),
  numberItem('揭题："计算机里存东西也一样——有些文件里面很多\'空气\'，我们可以去掉它们让文件变小。今天我们就来学怎么把\'大象塞进冰箱\'！"'),
  numberItem('板书课题：《把大象塞进冰箱》'),
  blank(),
  bodyBold('设计意图：'),
  body('用生活化的物理形象建立"压缩"的直觉。'),
  blank(),

  // ── 环节二 ──
  h2('环节二：游程编码——"AAAAA = 5个A"（20分钟）'),
  blank(),

  // Step 1
  h3('Step 1：发现冗余（5分钟）'),
  blank(),
  numberItem('在黑板上写两行"加密消息"：'),
  blank(),
  codeLine('第一行：A A A A A A A A A A A A A A A A A A A A'),
  codeLine('第二行：W E R T Y U I O P A S D F G H J K L Z X'),
  blank(),
  numberItem('提问：'),
  indentBody('"第一行有什么特点？" → 全是同一个字母A！'),
  indentBody('"如果让你传纸条，第一行你会怎么写？" → "20个A"'),
  indentBody('"同样的思路，\'BBBBBBBBBB\'你怎么写？" → "10个B"'),
  blank(),
  numberItem('引出核心思想："用\'几个什么\'来代替\'一个一个列出来\'，这就是压缩！"'),
  blank(),

  // Step 2
  h3('Step 2：动手压缩（10分钟）'),
  blank(),
  numberItem('给出第一条消息：'),
  blank(),
  codeLine('原始：B B B B B B B W W W W W W W A A A'),
  blank(),
  numberItem('教学生用游程编码规则：'),
  indentBody('写"7B" 代替"BBBBBBB"'),
  indentBody('写"7W" 代替"WWWWWWW"'),
  indentBody('写"3A" 代替"AAA"'),
  indentBody('压缩结果：7B 7W 3A'),
  blank(),
  numberItem('对比：原来需要17个字符，现在只需要7个字符（含空格）！'),
  blank(),
  numberItem('让学生自己练习：'),
  blank(),
  codeLine('练习题1：G G G G G → 5G'),
  codeLine('练习题2：Z Z Z Z Z Z Z Z Z Z → 10Z'),
  codeLine('练习题3：X X X X X Y Y Y Y Y Z Z Z Z Z → 5X 5Y 5Z'),
  codeLine('练习题4：A A A B B B B B C C C C C C C → 3A 5B 7C'),
  blank(),

  // Step 3
  h3('Step 3：进阶挑战——黑白图像（5分钟）'),
  blank(),
  numberItem('在黑板上画一个8×8的简单黑白格子图（比如一个黑色方块），用0表示白、1表示黑。'),
  numberItem('逐行写出原始编码（64个0和1），让学生感受"好多！"。'),
  numberItem('用游程编码压缩——"23个0、2个1、22个0……"。'),
  numberItem('对比："原来要写64个数字，现在只要写几个\'几个什么\'！这就是图片压缩的原理！"'),
  blank(),
  bodyBold('设计意图：'),
  body('从一维字母到二维图像，让学生感知压缩的普遍性和威力。'),
  blank(),

  // ── 环节三 ──
  h2('环节三：压缩的"边界"——不是什么东西都能变小（10分钟）'),
  blank(),

  h3('Step 1：反例演示（5分钟）'),
  blank(),
  numberItem('给出一条"随机"消息：'),
  blank(),
  codeLine('W E R T Y U I O P'),
  blank(),
  numberItem('让学生尝试用游程编码压缩——发现每个字母都不连续重复，写出来变成"1W 1E 1R 1T 1Y 1U 1I 1O 1P"，反而更长了！'),
  numberItem('结论："压缩只能去掉重复的东西，没有重复就没办法压缩。就像你没法把一块石头再变小——它里面没有空气。"'),
  blank(),

  h3('Step 2：生活例子（5分钟）'),
  blank(),
  numberItem('提问互动：'),
  indentBody('"什么样的衣服容易压缩？" → 羽绒服（很多空气）'),
  indentBody('"什么样的不容易？" → 石头（实的）'),
  indentBody('"照片为什么能压缩成更小的文件？" → 大片同色区域'),
  indentBody('"考试卷子扫描成PDF为什么能变小？" → 很多空白区域'),
  blank(),
  numberItem('总结：压缩 = 找规律 + 去掉重复'),
  blank(),

  // ── 环节四 ──
  h2('环节四：压缩接力赛（10分钟）'),
  blank(),
  bodyBold('游戏规则：'),
  numberItem('全班分成4组，每组拿到一条"原始消息"（见附录）。'),
  numberItem('比赛：哪组能压缩到最短。'),
  numberItem('每组选代表把压缩结果写在黑板上。'),
  numberItem('评选"最佳压缩大师"——压缩率最高的小组获胜。'),
  blank(),
  bodyBold('注意：'),
  body('压缩结果必须能还原！不能丢信息。（引入"无损压缩"概念——压缩了还能一模一样变回来。）'),
  blank(),

  // ── 环节五 ──
  h2('环节五：总结（5分钟）'),
  blank(),
  numberItem('回顾三个收获：'),
  indentBody('"压缩就是去掉重复、用更少的字符表达同样的信息。"'),
  indentBody('"\'几个什么\'（游程编码）是最简单的压缩方法。"'),
  indentBody('"没有规律的东西不好压缩——压缩需要找到重复。"'),
  numberItem('联系生活："你手机里的照片、视频、音乐，全都是压缩过的！不然手机根本存不下。"'),
  numberItem('预告下节课："学会了压缩，下节课我们来玩一个更有趣的游戏——模拟信息在网络上怎么传输，有\'丢包\'和\'黑客\'哦！"'),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 五、板书设计
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('五、板书设计'),
  blank(),
  codeLine('《把大象塞进冰箱》——数据压缩'),
  blank(),
  codeLine('  羽绒服 🧥 → 挤出空气 → 变小了！（还是那件衣服）'),
  blank(),
  codeLine('  文字压缩：'),
  codeLine('  AAAAAAAAAAAAAAAAAAAA  →  "20个A"'),
  codeLine('  BBBBBBBWWWWWWWAAA      →  7B 7W 3A'),
  blank(),
  codeLine('  规则："几个 + 什么" = 游程编码（RLE）'),
  blank(),
  codeLine('  ⚠ 没有规律 → 没法压缩！'),
  codeLine('  "WERTYUOP" → 1W 1E 1R… 反而更长了！'),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 六、附录：压缩挑战赛题目
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('六、附录：压缩挑战赛题目'),
  blank(),
  threeLineTable(
    ['编号', '原始消息', '参考答案'],
    [
      ['1', 'AAAAA BBBBB CCCCC DDDDD', '5A 5B 5C 5D'],
      ['2', 'XXXXXXXXXX YY ZZZZZZZZZZ', '10X 2Y 10Z'],
      ['3', 'KKK KKK KKK KKK', '3K(空格)3K(空格)3K(空格)3K → 12K（去掉空格）'],
      ['4', '0000000000 11111 0000000000 11111', '10个0 5个1 10个0 5个1 → 10W 5B 10W 5B'],
    ],
    [800, 4270, 4000]
  ),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 七、教学建议与注意事项
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('七、教学建议与注意事项'),
  blank(),
  numberItemBoldPrefix('游程编码名称：', '不必强制学生记"游程编码"或"RLE"这个词，他们只要能理解"几个什么"就够了。名词提一下即可。'),
  numberItemBoldPrefix('压缩率计算：', '如果有高年级学生，可以引入压缩率概念——压缩后大小÷原始大小。低年级跳过。'),
  numberItemBoldPrefix('羽绒服演示：', '如果现场没有羽绒服，用塑料袋+棉花/废纸团也可以，核心是展示"除去空隙"的直观感受。'),
  numberItemBoldPrefix('道具备选：', '没有实物的话，用手比划"大"→ 挤压 → "小"也行，配合夸张表情。'),
  numberItemBoldPrefix('趣味加分：', '如果有时间，可以让学生自己想一句"最难压缩"的话（每个字都不一样），互相挑战。'),
  blank(),
];

// ─────────────────────────────────────────────────────────────────────────────
// Build & write document
// ─────────────────────────────────────────────────────────────────────────────

const doc = new Document({
  styles:    STYLES,
  numbering: NUMBERING,
  sections: [{
    properties: {
      page: {
        size:   { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            indent:    { firstLine: 0 },
            children:  [new TextRun({ text: '数字魔法师 · 第3讲', size: 18, font: { ascii: 'Cambria Math', eastAsia: 'SimSun', hAnsi: 'Cambria Math' }, color: '888888' })],
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            indent:    { firstLine: 0 },
            children:  [
              new TextRun({ text: '第 ', size: 18 }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18 }),
              new TextRun({ text: ' 页', size: 18 }),
            ],
          }),
        ],
      }),
    },
    children: CONTENT,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUTPUT_PATH, buf);
  console.log(`✓  Written: ${OUTPUT_PATH}`);
}).catch(err => {
  console.error('Error building document:', err.message);
  process.exit(1);
});

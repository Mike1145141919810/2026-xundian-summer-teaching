/**
 * Build Li Fuzhou's teaching plan docx — 《信息大冒险》
 *
 * Usage:
 *   node build_04_doc.js
 * Outputs: 04-信息大冒险-李福洲.docx in the same directory
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
  PageBreak,
} = require('docx');

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const OUTPUT_PATH   = path.join(__dirname, '04-信息大冒险-李福洲.docx');

const PAGE_W        = 11906;
const PAGE_H        = 16838;
const MARGIN        = 1418;
const CONTENT_W     = PAGE_W - 2 * MARGIN;

const THICK = { style: BorderStyle.SINGLE, size: 12, color: '000000' };
const THIN  = { style: BorderStyle.SINGLE, size: 6,  color: '000000' };
const NONE  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };

let _chapter = 0;

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function body(text) {
  return new Paragraph({ children: [new TextRun(text)] });
}

function bodyBold(text) {
  return new Paragraph({ children: [new TextRun({ text, bold: true })] });
}

function blank() {
  return new Paragraph({ children: [] });
}

function h1Chinese(text) {
  _chapter++;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    numbering: { reference: `sections_c${_chapter}`, level: 0 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    numbering: { reference: `sections_c${_chapter}`, level: 1 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function codeLine(text) {
  return new Paragraph({
    indent: { left: 480, firstLine: 0 },
    spacing: { line: 240, lineRule: LineRuleType.AUTO },
    children: [new TextRun({ text, font: { ascii: 'Consolas', eastAsia: 'SimSun', hAnsi: 'Consolas' }, size: 21 })],
  });
}

function indentBody(text) {
  return new Paragraph({
    indent: { left: 480, firstLine: 0 },
    children: [new TextRun(text)],
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
    rows: rows.map(row =>
      new TableRow({
        children: row.map((cell, i) => cellOf(cell.text, colWidths[i], cell.bold || false)),
      })
    ),
  });
}

function bulletItemBoldPrefix(prefix, text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    children: [new TextRun({ text: prefix, bold: true }), new TextRun(text)],
  });
}

function numberItem(text) {
  return new Paragraph({
    numbering: { reference: 'numbers', level: 0 },
    children: [new TextRun(text)],
  });
}

function numberItemBoldPrefix(prefix, text) {
  return new Paragraph({
    numbering: { reference: 'numbers', level: 0 },
    children: [new TextRun({ text: prefix, bold: true }), new TextRun(text)],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Numbering — 6 chapters
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

const NUMBERING = buildNumberingConfig(6);

// ─────────────────────────────────────────────────────────────────────────────
// Styles
// ─────────────────────────────────────────────────────────────────────────────

const STYLES = {
  default: {
    document: {
      run: {
        font: { ascii: 'Cambria Math', hAnsi: 'Cambria Math', eastAsia: 'SimSun' },
        size: 24,
      },
      paragraph: {
        spacing: { line: 300, lineRule: LineRuleType.AUTO },
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
// ██  CONTENT  ████████████████████████████████████████████████████████████████
// ─────────────────────────────────────────────────────────────────────────────

const CONTENT = [

  // ── Title ──
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing:   { before: 0, after: 360 },
    indent:    { firstLine: 0 },
    children:  [new TextRun({
      text: '教案：《信息大冒险》',
      bold: true, size: 36,
      font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' },
    })],
  }),

  // ── Course Info ──
  borderlessTable([
    [{ text: '课程系列', bold: true }, { text: '数字魔法师 · 第4讲' }],
    [{ text: '授课教师', bold: true }, { text: '李福洲（计算机学院）' }],
    [{ text: '授课时间', bold: true }, { text: '7月30日 11:10-12:00（50分钟）' }],
    [{ text: '授课对象', bold: true }, { text: '寻甸县雷锋希望小学学生（3-6年级混龄）' }],
    [{ text: '课程类型', bold: true }, { text: '信息与编码启蒙 · 角色扮演游戏课' }],
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
      ['认知', '理解信息在网络中不是"一整段"传过去的，而是分成小包转发的；理解丢包和重传的基本概念'],
      ['技能', '能描述一次完整的信息传输过程（发送→分包→路由→组装）；能解释为什么需要"编号"和"确认"'],
      ['情感', '体验"协作完成任务"的乐趣；对"发一条微信背后发生了什么"产生好奇'],
    ],
    [1500, 7570]
  ),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 二、教学重难点
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('二、教学重难点'),
  blank(),
  bulletItemBoldPrefix('重点：', '信息分包传输、路由转发的过程'),
  bulletItemBoldPrefix('难点：', '理解为什么需要"编号"——数据包可能乱序到达，需要编号来重新排列'),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 三、教学准备
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('三、教学准备'),
  blank(),
  threeLineTable(
    ['物品', '数量', '说明'],
    [
      ['大信封（或纸袋）', '4-6个', '模拟"数据包"，贴上编号标签'],
      ['小纸条（消息内容）', '若干', '写有不同的消息，剪成小段'],
      ['角色标识牌', '4种各若干', '"发送者""路由器1""路由器2""接收者"'],
      ['胶带/绳子', '2条', '在地上贴出/拉出"网络路线"'],
      ['计时器（手机即可）', '1个', '记录传输时间'],
    ],
    [2800, 1500, 4770]
  ),
  blank(),
  indentBody('所有道具零经费：信封用纸折，标识牌用纸片写，路线用粉笔画或胶带贴。'),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 四、教学过程
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('四、教学过程'),
  blank(),

  // ── 环节一 ──
  h2('环节一：情景导入——"你的微信是怎么发出去的？"（5分钟）'),
  blank(),
  bodyBold('教师活动：'),
  numberItem('提问："你给朋友发一条微信，点击发送，对面马上就收到了。这中间发生了什么？"'),
  numberItem('学生自由猜测（"从天上飞的""走网线""WiFi信号"等）。'),
  numberItem('揭题："信息不是一整条飞过去的，而是被切成一小块一小块，像快递一样一站一站转过去的。今天我们来当快递员——把一条消息从教室这头传到那头！"'),
  numberItem('板书课题：《信息大冒险》'),
  blank(),

  // ── 环节二 ──
  h2('环节二：角色扮演第一轮——"直接传递"（10分钟）'),
  blank(),

  h3('布景准备（2分钟）'),
  blank(),
  numberItem('在地上画出路线：'),
  blank(),
  codeLine('发送者 ──── 路由器1 ──── 路由器2 ──── 接收者'),
  blank(),
  numberItem('分配角色：1个发送者、2个路由器、1个接收者，其余同学当"观察员"。'),
  blank(),

  h3('规则说明（2分钟）'),
  blank(),
  numberItem('发送者拿到一条完整消息（写在小纸条上），直接走过去递给路由器1。'),
  numberItem('路由器1再递给路由器2。'),
  numberItem('路由器2递给接收者。'),
  numberItem('接收者念出消息内容。'),
  blank(),

  h3('第一轮游戏（6分钟）'),
  blank(),
  numberItem('选一组学生上场，传递一条消息（如"下午三点操场集合"）。'),
  numberItem('观察员记录：用了多久？有没有出错？'),
  numberItem('讨论问题："如果路由器1\'死机\'了（退出游戏），消息怎么办？"'),
  indentBody('学生可能会说"换一条路"——引出"网络有备用路径"的概念。'),
  blank(),

  // ── 环节三 ──
  h2('环节三：角色扮演第二轮——"分包传输"（15分钟）'),
  blank(),

  h3('引入问题（3分钟）'),
  blank(),
  numberItem('教师写一条很长的消息在黑板上（40个字左右）。'),
  numberItem('提问："如果一条消息很长，一个\'快递员\'拿着纸条走，万一丢了怎么办？全都没了！"'),
  numberItem('引出方案："把大消息切成小块，每块加上编号，分开送。丢了也只丢一小块，重发那一块就行。"'),
  blank(),

  h3('分包演示（5分钟）'),
  blank(),
  numberItem('把长消息剪成4段，分别装进4个信封，信封上标编号：1/4、2/4、3/4、4/4。'),
  numberItem('发送者把4个信封分别交给4个不同的"数据包快递员"（新增角色）。'),
  numberItem('4个快递员同时出发，可以走不同路径（如果有多个路由器）。'),
  numberItem('接收者按编号重新排列，还原完整消息。'),
  blank(),

  h3('制造故障（7分钟）'),
  blank(),
  numberItemBoldPrefix('场景一：丢包——', '教师偷偷拿走其中一个信封（比如3/4）。'),
  indentBody('接收者发现少了3号包，怎么办？'),
  indentBody('引导："接收者可以喊\'3号包没到，请重发！\'"'),
  indentBody('发送者重新发送3号包。'),
  blank(),
  numberItemBoldPrefix('场景二：乱序到达——', '4个快递员故意以不同顺序到达。'),
  indentBody('接收者手中收到：1/4、4/4、2/4、3/4（乱序）。'),
  indentBody('引导学生发现："有编号就可以重新排好！1→2→3→4。"'),
  blank(),
  numberItemBoldPrefix('场景三：重复到达——', '同一个包被送了两次。'),
  indentBody('接收者发现两个3/4号包，怎么办？→ "扔掉重复的！"'),
  blank(),
  bodyBold('设计意图：'),
  body('通过"制造事故"的方式让学生理解网络协议的三大核心机制——丢包重传、编号排序、重复去重。'),
  blank(),

  // ── 环节四 ──
  h2('环节四：全班大游戏——"网络大冒险"（15分钟）'),
  blank(),

  h3('游戏规则（3分钟）'),
  blank(),
  numberItem('全班分成4个组，分别扮演：'),
  indentBody('发送组：负责拆分消息+编号+发出'),
  indentBody('路由组（2个组）：负责转发，但可以制造"故障"'),
  indentBody('接收组：负责接收+排序+还原+检查'),
  numberItem('教室布置成网络拓扑：'),
  blank(),
  codeLine('发送组 ──→ 路由组A ──→ 路由组B ──→ 接收组'),
  codeLine('            ↘          ↗'),
  codeLine('              路由组C'),
  blank(),
  numberItem('规则：'),
  indentBody('发送组把消息分成最少5个包，编号后发出。'),
  indentBody('路由组可以随机"丢包"（扣下不传），但不能丢超过2个。'),
  indentBody('路由组可以"延迟"某些包（最后再传）。'),
  indentBody('接收组还原消息后，检查是否完整，不完整就喊"重传编号X！"'),
  blank(),

  h3('游戏进行（10分钟）'),
  blank(),
  numberItem('第一轮：用一条20字的简单消息。'),
  numberItem('教师计时，记录各环节用时。'),
  numberItem('第二轮：换一组消息，难度升级——路由组增加干扰（故意传错编号、多传一个别人的包等）。'),
  blank(),

  h3('游戏后讨论（2分钟）'),
  blank(),
  bulletItemBoldPrefix('', '"接收组是怎么知道消息\'完整\'了？" → "看编号，连号就是完整。"'),
  bulletItemBoldPrefix('', '"如果你的微信消息在路上丢了，你的手机会怎么做？" → "自动重发，你根本不知道。"'),
  blank(),

  // ── 环节五 ──
  h2('环节五：总结（5分钟）'),
  blank(),
  numberItem('回顾三个核心概念（板书）：'),
  blank(),
  codeLine('信息传输三要素：'),
  codeLine('① 分包 —— 把大消息切成小块'),
  codeLine('② 编号 —— 每块标上"第几块/共几块"'),
  codeLine('③ 重传 —— 丢了就喊"再来一次！"'),
  blank(),
  numberItem('联系生活："你每发一条微信，背后就有无数个\'快递员\'在帮你传这些小包——这就是互联网的运作方式。"'),
  numberItem('预告下节课："明天就是我们的最后一节课——《画里的数字》，我们来看看一张照片是怎么变成数字的！"'),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 五、板书设计
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('五、板书设计'),
  blank(),
  codeLine('《信息大冒险》——网络是怎么传信息的'),
  blank(),
  codeLine('  发送者 ──→ 路由器1 ──→ 路由器2 ──→ 接收者'),
  blank(),
  codeLine('  📦 分包：大消息 → 切成小块'),
  codeLine('  🔢 编号：每块标号 1/4, 2/4, 3/4, 4/4'),
  codeLine('  🔄 重传：丢了就喊"X号包请重发！"'),
  blank(),
  codeLine('  核心三要素：分包 | 编号 | 重传'),
  blank(),
  codeLine('  你的微信、视频、游戏，都是这么传的！ ✨'),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 六、教学建议与注意事项
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('六、教学建议与注意事项'),
  blank(),
  numberItemBoldPrefix('空间需求：', '需要教室前面留出足够的走动空间。如果教室太小，可以改为"传纸条"形式，学生坐着传递，通过多列传递来模拟路由。'),
  numberItemBoldPrefix('角色轮换：', '确保让尽量多的学生都有机会当一次"发送者"或"路由器"，而不是全程旁观。可以轮换角色多玩几轮。'),
  numberItemBoldPrefix('混龄应对：', '低年级多参与简单的"传纸条"角色；高年级可以担任"接收组"负责还原和检查。'),
  numberItemBoldPrefix('故障设计：', '丢包、乱序、重复等"故障"是本节课最出彩的地方，教师要在游戏中自然制造——不要提前预告，让学生在体验中自己发现问题和解决方案。'),
  numberItemBoldPrefix('课堂纪律：', '走动环节容易混乱，需要提前讲清楚"不可跑、不可推挤"的安全规则。教师喊"开始"和"停"作为信号。'),
  numberItemBoldPrefix('趣味加分：', '可以给每个角色配一个小道具——发送者拿"手机"（纸板做的），路由器拿"文件夹"，接收者拿"放大镜"，增强代入感。'),
  blank(),
];

// ─────────────────────────────────────────────────────────────────────────────
// Build & write
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
            children:  [new TextRun({ text: '数字魔法师 · 第4讲', size: 18, font: { ascii: 'Cambria Math', eastAsia: 'SimSun', hAnsi: 'Cambria Math' }, color: '888888' })],
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

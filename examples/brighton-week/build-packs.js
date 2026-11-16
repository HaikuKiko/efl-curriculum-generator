// Builds the Brighton Skills Week student and teacher packs as Word files.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, PageOrientation, LevelFormat,
  TabStopType, PageNumber, VerticalAlign, ExternalHyperlink, HeightRule,
} = require('docx');

const OUT = process.argv[2];
const LOGO = fs.readFileSync(process.argv[3]);

const FONT = 'Arial';
const PURPLE = '7B3F8C';
const LIGHT = 'F3ECF6';
const A4_W = 11906, A4_H = 16838;
const MARGIN = 1000;
const PORTRAIT_W = A4_W - 2 * MARGIN;   // usable width, portrait
const LANDSCAPE_W = A4_H - 2 * MARGIN;  // usable width, landscape

// ---------- text helpers ----------
// Supports **bold** and *italic* inside a string.
function runs(text, base = {}) {
  const parts = String(text).split(/(\*\*[^*]+\*\*|\*[^*]+\*)/).filter(Boolean);
  return parts.map((p) => {
    if (p.startsWith('**')) return new TextRun({ text: p.slice(2, -2), bold: true, ...base });
    if (p.startsWith('*')) return new TextRun({ text: p.slice(1, -1), italics: true, ...base });
    return new TextRun({ text: p, ...base });
  });
}
const P = (text, opts = {}) =>
  new Paragraph({ children: runs(text, opts.run || {}), spacing: { after: 120, ...(opts.spacing || {}) }, alignment: opts.align, ...(opts.para || {}) });
const H1 = (text, pageBreak = true) =>
  new Paragraph({ heading: 'Heading1', pageBreakBefore: pageBreak, children: [new TextRun(text)] });
const H2 = (text) => new Paragraph({ heading: 'Heading2', children: [new TextRun(text)] });
const H3 = (text) => new Paragraph({ heading: 'Heading3', children: [new TextRun(text)] });
const Small = (text) => P(text, { run: { size: 18, color: '666666' } });
const NameLine = () =>
  new Paragraph({
    spacing: { after: 200 },
    tabStops: [{ type: TabStopType.LEFT, position: 5200 }],
    children: [new TextRun({ text: 'Name: ____________________________', size: 20 }), new TextRun({ text: '\tDate: ________________', size: 20 })],
  });
const WriteLine = () =>
  new Paragraph({ spacing: { before: 280, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '999999', space: 1 } }, children: [] });
const lines = (n) => Array.from({ length: n }, WriteLine);

let listInstance = 0;
function numbered(items) {
  const inst = ++listInstance;
  return items.map((t) => new Paragraph({ numbering: { reference: 'num', level: 0, instance: inst }, spacing: { after: 80 }, children: runs(t) }));
}
const bullets = (items) => items.map((t) => new Paragraph({ numbering: { reference: 'bullet', level: 0 }, spacing: { after: 60 }, children: runs(t) }));
const checks = (items) => items.map((t) => new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: '☐  ', size: 26 }), ...runs(t)] }));

// ---------- table helpers ----------
const thin = { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB' };
const dashed = { style: BorderStyle.DASHED, size: 8, color: '888888' };
const borders = (b) => ({ top: b, bottom: b, left: b, right: b });

function cell(content, width, opts = {}) {
  const children = Array.isArray(content) ? content : [P(content, { run: opts.run || {}, align: opts.align })];
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: borders(opts.border || thin),
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: opts.vAlign || VerticalAlign.TOP,
    children,
  });
}
function table(rows, widths, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: opts.header && i === 0,
      height: opts.rowHeight && !(opts.header && i === 0) ? { value: opts.rowHeight, rule: HeightRule.ATLEAST } : undefined,
      cantSplit: true,
      children: r.map((c, j) => cell(c, widths[j], {
        fill: opts.header && i === 0 ? PURPLE : undefined,
        run: opts.header && i === 0 ? { bold: true, color: 'FFFFFF' } : {},
        border: opts.border,
      })),
    })),
  });
}
// Cut-up cards: n columns, dashed borders.
function cards(texts, cols = 2, height = 1500, size = 26) {
  const w = Math.floor(PORTRAIT_W / cols);
  const widths = Array(cols).fill(w);
  const rows = [];
  for (let i = 0; i < texts.length; i += cols) {
    const slice = texts.slice(i, i + cols);
    while (slice.length < cols) slice.push('');
    rows.push(new TableRow({
      height: { value: height, rule: HeightRule.ATLEAST },
      cantSplit: true,
      children: slice.map((t) => cell([P(t, { run: { size }, align: AlignmentType.CENTER })], w, { border: dashed, vAlign: VerticalAlign.CENTER })),
    }));
  }
  return new Table({ width: { size: w * cols, type: WidthType.DXA }, columnWidths: widths, rows });
}
function box(children, width = PORTRAIT_W, fill = LIGHT) {
  return new Table({
    width: { size: width, type: WidthType.DXA },
    columnWidths: [width],
    rows: [new TableRow({ children: [cell(children, width, { fill, border: { style: BorderStyle.SINGLE, size: 8, color: PURPLE } })] })],
  });
}
const spacer = () => new Paragraph({ spacing: { after: 120 }, children: [] });
const link = (text, url) => new Paragraph({ spacing: { after: 80 }, children: [new ExternalHyperlink({ link: url, children: [new TextRun({ text, style: 'Hyperlink' })] })] });

// ---------- page furniture ----------
function header(subtitle) {
  return new Header({
    children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: PORTRAIT_W }],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: PURPLE, space: 4 } },
      children: [
        new ImageRun({ type: 'png', data: LOGO, transformation: { width: 40, height: 40 } }),
        new TextRun({ text: '  Kings Brighton', bold: true, color: PURPLE, size: 22 }),
        new TextRun({ text: `\t${subtitle}`, color: '666666', size: 18 }),
      ],
    })],
  });
}
function landscapeHeader(subtitle) {
  return new Header({
    children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: LANDSCAPE_W }],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: PURPLE, space: 4 } },
      children: [
        new ImageRun({ type: 'png', data: LOGO, transformation: { width: 40, height: 40 } }),
        new TextRun({ text: '  Kings Brighton', bold: true, color: PURPLE, size: 22 }),
        new TextRun({ text: `\t${subtitle}`, color: '666666', size: 18 }),
      ],
    })],
  });
}
const footer = () => new Footer({
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '888888' })] })],
});

const styles = {
  default: { document: { run: { font: FONT, size: 22 } } },
  paragraphStyles: [
    { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: FONT, size: 34, bold: true, color: PURPLE }, paragraph: { spacing: { before: 0, after: 160 }, keepNext: true, outlineLevel: 0 } },
    { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: FONT, size: 26, bold: true, color: '333333' }, paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 1 } },
    { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: FONT, size: 22, bold: true, color: PURPLE }, paragraph: { spacing: { before: 200, after: 80 }, keepNext: true, outlineLevel: 2 } },
  ],
};
const numbering = {
  config: [
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 360 } } } }] },
    { reference: 'bullet', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] },
  ],
};
const portrait = { page: { size: { width: A4_W, height: A4_H }, margin: { top: 1200, bottom: 900, left: MARGIN, right: MARGIN, header: 400 } } };
const landscape = { page: { size: { width: A4_W, height: A4_H, orientation: PageOrientation.LANDSCAPE }, margin: { top: 1100, bottom: 700, left: MARGIN, right: MARGIN, header: 400 } } };

// ================= CONTENT =================
const TF_DAY1 = [
  'The Royal Pavilion has domes like a palace in India.',
  'Brighton is the capital of England.',
  'The Lanes are wide, modern streets.',
  'Brighton has two piers, but one is in ruins.',
  'You can swim in the sea in Brighton.',
  'Hove is in a different country.',
  "There's a big arts festival in Brighton every May.",
  'It never rains in Brighton.',
];
const FSW = [
  'has been on Palace Pier', 'has swum in the sea in Brighton', 'likes fish and chips', 'lives in Hove',
  'took a photo of a seagull this week', 'can name a street in Brighton', 'has been inside the Royal Pavilion',
  'takes the bus to school', 'bought something in the Lanes', 'is going to go to London soon',
];
const INTERVIEW = [
  'Where do you live in Brighton? Who do you live with?', 'How do you get to school? How long does it take?',
  'What did you do last weekend?', "What's your favourite place in Brighton so far? Why?",
  'What can you see from your bedroom window?', "Is Brighton bigger or smaller than your home town? What's different?",
  'What English food have you tried? Did you like it?', 'What are you going to do this weekend?',
];
const POSTCARD = [
  'Dear Grandma,',
  "I'm in Brighton! I'm staying with a nice host family in Hove. On Saturday I went to Palace Pier with my friends and we ate fish and chips. A seagull tried to steal my food! Today it's sunny and I'm sitting on the beach. I can see the sea and the i360. Next weekend I'm going to visit the Royal Pavilion. I miss you!",
  'Love,',
  'Sofia',
];
const CHECKLIST = [
  'Greeting (*Dear... / Hi...*)', "Where I am and who I'm staying with", 'One thing I did (past simple)',
  "What I'm doing or can see now", 'One plan (*going to*)', 'Goodbye (*Love, / See you soon,*)', '60–80 words',
];
const PLACES = ['The Royal Pavilion', 'Palace Pier', 'the i360', 'the Lanes', 'North Laine', 'the Clock Tower',
  'Churchill Square', 'Brighton station', 'the Sea Life Centre', 'the West Pier', 'Brighton Marina', 'the beach'];
const TOURISTS = [
  'You love **history** and old buildings. You have one afternoon.',
  'You have **£10** and you love shopping for unusual things.',
  "It's **raining**. You're with a **7-year-old child**.",
  'You want the **best view** of Brighton for photos.',
  'You love **nature and walking**. You have a whole day.',
  'You want to try **typical English food**.',
];
const LISTEN_PLACES = ['the Lanes', 'the beach', 'the seafront', 'the Royal Pavilion', 'the pier', 'the Marina', 'the i360', 'the Booth Museum'];
const LISTEN_TF = [
  'Lucas went to the city centre with Yuki.', 'Yuki bought a ring for her mother.', 'The ring cost £20.',
  'Brighton beach has a lot of sand.', "A seagull took one of Lucas's chips.", 'Yuki won a toy on the pier.',
  'Lucas thinks the i360 was the best thing all day.', 'They went home by train.',
];
const GAPS = [
  'It cost ______________ pounds.', 'Then we went to the beach, at about ______________ o’clock.',
  'For lunch we had fish and ______________ on the seafront.', 'We could see all of Brighton and the ______________.',
  'We went home on the bus at ______________ o’clock.', "Next weekend I'm going to visit the Royal ______________.",
];
const BINGO = ['pier', 'pebbles', 'seagull', 'seafront', 'beach', 'tower', 'view', 'lane', 'palace', 'queue', 'shop',
  'fish and chips', 'bus', 'waves', 'sunny', 'windy'];
const JAM = ['The beach in Brighton', 'My host family or my flat', 'English food', 'My home town', 'The weather this week',
  'My best day in Brighton', 'Seagulls', 'Shopping in Brighton', 'My school', 'Public transport',
  'A place I want to visit in England', 'My weekend plans'];
const CORNERS = ['AGREE', 'DISAGREE', 'NOT SURE', "IT'S THE SAME IN MY COUNTRY"];
const CULTURE = [
  'English people queue for everything.', 'English people say "sorry" all the time.',
  'Tea with milk is the best drink in the world.', "It's rude to ask someone's age or salary.",
  'People talk about the weather all the time.', 'Dinner at 6 pm is too early.',
  'School students should wear a uniform.', "You should take your shoes off when you visit someone's house.",
];
const BOARD = [
  'START', 'Say three things you can do in Brighton.', 'What did you do yesterday evening?',
  'Compare Brighton and your home town.', 'Describe your bedroom: There is... There are...',
  'What are you going to do next weekend?', 'Go forward 2!', "Describe the best meal you've had in England.",
  "What can't you do well? What can you do well?", 'What should a new student do in their first week?',
  "Tell us about a place in Brighton you don't like.", 'Ask the player on your left 3 questions.',
  'What did you do last Saturday?', "What's the most beautiful place in Brighton?", 'Go back 3!',
  'Describe a typical day at school.', 'Which is better: the beach or the city? Why?',
  "What's in your bag? (There is... There are...)", 'Tell us about your host family or flatmates.',
  'What are you going to buy before you go home?', 'Miss a turn!', 'What was the weather like last weekend?',
  'Give directions from school to the beach.', "What's the strangest English habit?",
  'Say 5 things that are cheaper in your country.', 'What did you eat for breakfast today?', 'Go forward 1!',
  'Tell a story about a seagull (true or invented!).', "What's the best thing you learned this week?",
  'FINISH! Say "Thank you, Brighton!"',
];
const SCRIPT = [
  "Hi Ana, it's Lucas! I want to tell you about my Saturday in Brighton. I went into the city centre with my friend Yuki. First we walked around the Lanes. There are a lot of small shops there, and Yuki bought a ring for her sister. It cost twelve pounds.",
  "Then we went to the beach, at about one o'clock. But there isn't any sand! There are only pebbles. It wasn't very comfortable. For lunch we had fish and chips on the seafront. But a seagull took one of my chips! I was so surprised.",
  'After lunch we went on the pier. We played games, and Yuki won a small toy. Then in the afternoon we went up the i360. The view was amazing. We could see all of Brighton and the sea. It was a bit expensive, but it was the best thing all day.',
  "We were really tired, so we went home on the bus at seven o'clock. Next weekend I'm going to visit the Royal Pavilion. It's near my school, but I haven't been inside yet. OK, bye! Call me soon!",
];
const KAHOOTS = [
  ['1 Tue', 'Brighton basics', 'https://create.kahoot.it/details/739bef19-5282-47db-9c2d-305e02e691e5'],
  ['2 Wed', 'English culture', 'https://create.kahoot.it/details/7d6971c3-87f7-42f0-8808-e6dd0b2d4467'],
  ['3 Thu', 'Places in Brighton', 'https://create.kahoot.it/details/75dc390b-2070-49ef-965b-2cafad5249b4'],
  ['4 Fri', 'A Saturday in Brighton', 'https://create.kahoot.it/details/190ddd16-7748-4bfc-a4ab-1f8f8a6b5771'],
  ['5 Mon', 'Big week review', 'https://create.kahoot.it/details/2a9b54b3-568b-46df-8316-0f868e6fe796'],
];

// ================= STUDENT PACK =================
function studentPack() {
  const s = [];
  // Contents page
  s.push(new Paragraph({ heading: 'Heading1', children: [new TextRun('Brighton Skills Week: Student worksheets')] }));
  s.push(P('2A Elementary (A1+ to A2). Print one copy per student unless the page says otherwise.'));
  s.push(table([
    ['Page', 'Worksheet', 'Day', 'Copies'],
    ['2', 'Find someone who...', 'Tue', '1 per student'],
    ['3', 'Brighton: true or false?', 'Tue', '1 per team'],
    ['4', 'My first week: interview', 'Tue', '1 per student'],
    ['5', 'A postcard from Brighton (model)', 'Wed', '1 per student'],
    ['6', 'My postcard (writing frame)', 'Wed', '1 per student'],
    ['7', 'Secret places (or show on screen)', 'Thu', '1 per pair'],
    ['8', 'Tourist role cards (cut up)', 'Thu', '1 set per 6 students'],
    ['9', 'Listening: A Saturday in Brighton', 'Fri', '1 per student'],
    ['10', 'Brighton bingo', 'Fri', '1 per 2 students (cut in half)'],
    ['11', 'Just a minute topic cards (cut up)', 'Fri', '1 set per group'],
    ['12–15', 'Four corners signs', 'Mon', '1 set for the room'],
    ['16', 'Brighton board game (A4, or enlarge to A3)', 'Mon', '1 per group'],
  ], [1100, 4700, 1100, 3006], { header: true }));

  // 1 Find someone who
  s.push(H1('Find someone who...'));
  s.push(NameLine());
  s.push(P('Ask *"Have you...?"* or *"Do you...?"*. Write a name, then ask one more question: *When? Who with? Did you like it?*'));
  s.push(table([['Find someone who...', 'Name', 'Extra information'], ...FSW.map((f) => [f, '', ''])], [4200, 2200, 3506], { header: true, rowHeight: 700 }));

  // 2 True or false
  s.push(H1('Brighton: true or false?'));
  s.push(P('Team name: ______________________'));
  s.push(P('Listen to your teacher. Write **T** (true) or **F** (false). If it’s false, correct it for a bonus point!'));
  s.push(table([['#', 'Sentence', 'T / F', 'Correction'], ...TF_DAY1.map((t, i) => [String(i + 1), t, '', ''])], [500, 4900, 900, 3606], { header: true, rowHeight: 650 }));

  // 3 Interview
  s.push(H1('My first week in Brighton: interview'));
  s.push(NameLine());
  s.push(P("Ask your partner the questions. Write short notes. Then tell a new group about your partner: *\"Maria lives in Hove. She takes the bus...\"*"));
  s.push(P("My partner's name: ______________________"));
  s.push(table([['Question', 'Notes'], ...INTERVIEW.map((q) => [q, ''])], [5000, 4906], { header: true, rowHeight: 900 }));

  // 4 Postcard model
  s.push(H1('A postcard from Brighton'));
  s.push(NameLine());
  s.push(H3('1. Read the postcard'));
  s.push(box(POSTCARD.map((l, i) => P(l, { run: { size: 24, italics: i !== 1 ? false : false }, spacing: { after: i === 1 ? 160 : 60 } }))));
  s.push(H3('2. Answer the questions'));
  ['Who is the postcard to?', 'What did Sofia do on Saturday?', "What's the weather like today?", 'What is she going to do next weekend?']
    .forEach((q, i) => { s.push(P(`${i + 1}. ${q}`, { spacing: { after: 0 } })); s.push(WriteLine()); s.push(spacer()); });
  s.push(H3('3. Find the parts'));
  s.push(P('Number the parts of the postcard in order:'));
  s.push(table([
    ['___ what I can see now', '___ goodbye', '___ greeting'],
    ['___ what I did', '___ plans', '___ where I am'],
  ], [3302, 3302, 3302]));
  s.push(spacer());
  s.push(P('Underline one example of: **past simple** · **present continuous** · **can** · **going to**'));

  // 5 Writing frame
  s.push(H1('My postcard'));
  s.push(P('**Plan.** Make notes for each part.'));
  s.push(table([
    ['Where I am / who with', ''], ['One thing I did', ''], ['What I can see now', ''], ['My plan', ''],
  ], [3000, 6906], { rowHeight: 600 }));
  s.push(spacer());
  s.push(P('**Write** 60–80 words. Write your name on the back only!'));
  s.push(box([P('Dear ______________________,', { spacing: { after: 0 } }), ...lines(11)], PORTRAIT_W, 'FFFFFF'));
  s.push(H3('Check with a partner'));
  s.push(...checks(CHECKLIST));

  // 6 Secret places
  s.push(H1('Secret places'));
  s.push(P('Choose a secret place. Your partner asks up to 10 yes/no questions to guess it.'));
  s.push(cards(PLACES, 3, 1000, 28));
  s.push(spacer());
  s.push(box([
    P('**Question starters**'),
    ...bullets(['Is it near...? Is it next to...? Is it opposite...?', 'Can you... there?', 'Is it bigger / older than...?', 'Is there a... there?']),
  ]));
  s.push(spacer());
  s.push(box([
    P('**Directions**'),
    P('go straight on · turn left / right · take the second left · it’s on your right · next to · opposite · between · on the corner of · at the end of'),
  ]));

  // 7 Tourist cards
  s.push(H1('Tourist role cards'));
  s.push(Small('Cut along the dashed lines. One card per tourist.'));
  s.push(cards(TOURISTS.map((t) => `**TOURIST**\n${t}`).map((t) => t.replace('\n', ' — ')), 2, 1700, 26));
  s.push(spacer());
  s.push(box([
    P('**Locals: useful language**'),
    ...bullets(['You should go to...', "Why don't you try...?", "It's the best place for...", "It's cheaper / nicer / more interesting than...", "Don't go to... because..."]),
  ]));

  // 8 Listening
  s.push(H1('Listening: A Saturday in Brighton'));
  s.push(NameLine());
  s.push(P('Lucas leaves a voice message for his sister, Ana.'));
  s.push(H3('Task 1. Tick (✓) the places Lucas visited.'));
  s.push(table([LISTEN_PLACES.slice(0, 4).map((p) => `☐  ${p}`), LISTEN_PLACES.slice(4).map((p) => `☐  ${p}`)], [2476, 2476, 2477, 2477]));
  s.push(H3('Task 2. True or false?'));
  s.push(table([['#', 'Sentence', 'T / F'], ...LISTEN_TF.map((t, i) => [String(i + 1), t, ''])], [500, 8006, 1400], { header: true }));
  s.push(H3('Task 3. Listen again and complete.'));
  s.push(...numbered(GAPS));
  s.push(H3('Task 4. Talk to your partner.'));
  s.push(P('Tell your partner about your Saturday. Where did you go? What did you eat? Was it better or worse than Lucas’s day?'));

  // 9 Bingo
  s.push(H1('Brighton bingo'));
  s.push(P('Choose 9 words from the box and write them in your grid. Listen to your teacher. Cross out the words. Get a full line? Shout **"Brighton!"**'));
  s.push(box([P(BINGO.join('  ·  '), { align: AlignmentType.CENTER, run: { size: 24 } })]));
  const grid = () => new Table({
    width: { size: 6300, type: WidthType.DXA }, columnWidths: [2100, 2100, 2100], alignment: AlignmentType.CENTER,
    rows: [0, 1, 2].map(() => new TableRow({ height: { value: 1300, rule: HeightRule.EXACT }, children: [0, 1, 2].map(() => cell('', 2100, { border: { style: BorderStyle.SINGLE, size: 12, color: PURPLE } })) })),
  });
  s.push(spacer()); s.push(grid());
  s.push(new Paragraph({ spacing: { before: 300, after: 300 }, border: { bottom: dashed }, children: [new TextRun({ text: '✂  cut here', size: 16, color: '888888' })] }));
  s.push(P('Choose 9 words. Write them in your grid.'));
  s.push(grid());

  // 10 Just a minute
  s.push(H1('Just a minute: topic cards'));
  s.push(Small('Cut up one set per group. Speak for 30 seconds without stopping (60 seconds for a challenge!).'));
  s.push(cards(JAM, 3, 1600, 26));

  // 11 Four corners
  CORNERS.forEach((c, i) => {
    s.push(new Paragraph({ pageBreakBefore: true, spacing: { before: 3600 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: c, bold: true, size: c.length > 12 ? 96 : 140, color: PURPLE })] }));
  });

  return s;
}

function boardGame() {
  const COLS = 6, ROWS = 5;
  const w = Math.floor(LANDSCAPE_W / COLS);
  const rows = [];
  for (let r = 0; r < ROWS; r++) {
    const idx = Array.from({ length: COLS }, (_, c) => r * COLS + c);
    if (r % 2 === 1) idx.reverse();
    rows.push(new TableRow({
      height: { value: 1700, rule: HeightRule.EXACT },
      cantSplit: true,
      children: idx.map((i) => {
        const special = i === 0 || i === 29;
        const fun = /^(Go|Miss)/.test(BOARD[i]);
        return cell([
          new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: String(i + 1), bold: true, size: 28, color: special ? 'FFFFFF' : PURPLE })] }),
          new Paragraph({ children: [new TextRun({ text: BOARD[i], size: special ? 26 : 18, bold: special || fun, color: special ? 'FFFFFF' : '222222' })] }),
        ], w, { fill: special ? PURPLE : fun ? 'FBE9F1' : (r + idx.indexOf(i)) % 2 ? 'FFFFFF' : LIGHT, border: { style: BorderStyle.SINGLE, size: 8, color: PURPLE } });
      }),
    }));
  }
  return [
    new Paragraph({ heading: 'Heading1', children: [new TextRun('Brighton board game')] }),
    P('Roll the dice and move. Do the task and speak for at least 20 seconds. The group asks you one question. Can’t do it? Go back 2. Follow the numbers: the path snakes back and forth.', { run: { size: 18 } }),
    new Table({ width: { size: w * COLS, type: WidthType.DXA }, columnWidths: Array(COLS).fill(w), rows }),
  ];
}

// ================= TEACHER PACK =================
function teacherPack() {
  const t = [];
  t.push(new Paragraph({ heading: 'Heading1', children: [new TextRun('Brighton Skills Week: Teacher pack')] }));
  t.push(P('**Level:** 2A Elementary (A1+ to A2) · **Length:** one 90-minute skills lesson a day, taught as two halves (Lesson 1 and Lesson 2) · **Grammar review:** present simple, past simple, *can*, *there is / there are*, comparatives and superlatives, *going to*.'));
  t.push(P('Mostly structured speaking, plus one writing lesson (Wednesday), one listening lesson (Friday) and a Kahoot every day. Student worksheets are in the separate student pack.'));
  t.push(H2('Week at a glance'));
  t.push(table([
    ['Day', 'First half', 'Second half', 'Kahoot'],
    ['Tue', 'Find someone who + Brighton true or false (S)', 'My first week: interview and report back (S)', '1 Brighton basics'],
    ['Wed', 'Writing: a postcard from Brighton (R)(W)', 'Postcard gallery (R)(S)', '2 English culture'],
    ['Thu', 'Guess the place + big-screen directions (S)(L). Overflow for writing', 'Recommend a place: tourist and local role-play (S)', '3 Places in Brighton'],
    ['Fri', 'Listening: A Saturday in Brighton (L)(S)', 'Brighton bingo + Just a minute (S)', '4 A Saturday in Brighton'],
    ['Mon', '"Is it true in your country?" four corners (S)', 'Brighton board game: grammar review (S)', '5 Big week review'],
  ], [900, 3600, 3400, 2006], { header: true }));
  t.push(H2('Kahoots'));
  t.push(P('These links open each kahoot\u2019s page, where you press **Start** to host it or **Duplicate** to copy it into your own account. They work once the kahoots are set to public.'));
  t.push(new Table({
    width: { size: PORTRAIT_W, type: WidthType.DXA }, columnWidths: [1200, 3400, 5306],
    rows: [
      new TableRow({ tableHeader: true, children: ['Day', 'Kahoot', 'Link'].map((h, i) => cell(h, [1200, 3400, 5306][i], { fill: PURPLE, run: { bold: true, color: 'FFFFFF' } })) }),
      ...KAHOOTS.map(([d, n, u]) => new TableRow({ children: [cell(d, 1200), cell(n, 3400), cell([new Paragraph({ children: [new ExternalHyperlink({ link: u, children: [new TextRun({ text: 'Open the kahoot', style: 'Hyperlink' })] })] })], 5306)] })),
    ],
  }));
  t.push(H2('Prep checklist'));
  t.push(...checks([
    'Print the student pack (copies listed on its first page).',
    'Cut up: tourist role cards, Just a minute cards, bingo sheets.',
    'Print the board game, one per group (A3 if you can). Bring dice and counters.',
    'Thursday: open Google Maps on central Brighton on the classroom screen. Nothing to print.',
    'Friday: read the listening script aloud, or paste it into Word and use Review → Read Aloud. Nothing to record.',
    'Monday: stick the four corners signs in the corners of the room.',
  ]));

  // Day plans
  const day = (title, intro, lessons) => {
    t.push(H1(title));
    t.push(P(intro));
    lessons.forEach(([h, aim, steps, extra]) => {
      t.push(H2(h));
      t.push(P(`**Aim:** ${aim}`));
      t.push(...numbered(steps));
      if (extra) t.push(P(extra));
    });
  };
  day('Tuesday: Welcome to Brighton', 'Students find out what they already know about Brighton and talk about their first days here.', [
    ['First half: Find someone who + Brighton true or false (S)', 'students ask and answer simple questions about life in Brighton, with follow-up questions.', [
      '**Kahoot 1: Brighton basics (10 min).** Run it cold as a warmer. Don’t teach anything yet; note which facts surprise them.',
      '**Find someone who (15 min).** Students mingle and ask *"Have you been to the pier?"*. When someone says yes, they write the name and ask one follow-up question.',
      '**Feedback (5 min).** "Who found someone who has swum in the sea? Tell us about it." Board useful words.',
      '**Brighton true or false (10 min).** Teams of three. Read the 8 statements; teams write T or F. A point per right answer, a bonus point for correcting a false one.',
      '**Say it right (5 min).** Drill *Pavilion, pier, the Lanes, Hove, seafront, i360*. Students mark the stress.',
    ]],
    ['Second half: My first week, interview and report back (S)', 'present simple, past simple, comparatives and *going to* in a personal interview, then reporting in the third person.', [
      '**Lead-in (5 min).** Board: *Brighton: one word*. Each student says one word about Brighton so far.',
      '**Interview (15 min).** Pairs use the interview sheet. A asks all 8 questions and makes notes; then swap.',
      '**Report back (15 min).** New groups of four. Each student tells the group about their partner: *"Maria lives in Hove. She takes the 7 bus..."* Listeners ask one question each.',
      '**Class survey (5 min).** "What’s the best place in Brighton so far?" Tally on the board.',
      '**Error clinic (5 min).** Board 5 errors you heard (no names). Teams correct them.',
    ]],
  ]);
  day('Wednesday: A postcard from Brighton', 'Students write a 60–80 word postcard home, then read each other’s and guess the writer. If you run out of time, finish the writing in Thursday’s first half and do the gallery straight after; move Thursday’s places lesson to Friday’s second half.', [
    ['First half: Writing a postcard from Brighton (R)(W)', 'students write a short informal postcard using past simple, present continuous, *can* and *going to*.', [
      '**Lead-in (5 min).** "Who do you send messages to at home? What do you tell them?"',
      '**Read the model (10 min).** Students read Sofia’s postcard and answer the four questions.',
      '**Notice the structure (5 min).** Number the parts; underline one example of each tense.',
      '**Plan (5 min).** Students make notes in the plan boxes.',
      '**Write (15 min).** 60–80 words in the postcard frame. Name on the back only.',
      '**Check (5 min).** Swap with a partner and use the checklist.',
    ]],
    ['Second half: Postcard gallery + Kahoot 2 (R)(S)', 'students read for specific information and talk about their week.', [
      '**Gallery (15 min).** Number the postcards and stick them on the walls, names hidden. Students read them all and write who they think wrote each one.',
      '**Guess who (10 min).** Go through the numbers. *"I think it’s Ahmed because he went to the football."* The writer reveals.',
      '**Hot seat (10 min).** 3 or 4 authors answer questions about their week.',
      '**Kahoot 2: English culture (10 min).** A closer that sets up Monday’s culture lesson.',
    ], '**Homework:** write a real email or message to someone at home about Brighton, 80–100 words.'],
  ]);
  day('Thursday: Brighton places', 'Students describe where places are, give directions and recommend places to a visitor. **You need:** the classroom screen with Google Maps open on central Brighton. Nothing to print except the role cards.', [
    ['First half: Guess the place + big-screen directions (S)(L)', 'prepositions of place and simple directions.', [
      '**Kahoot 3: Places in Brighton (10 min).** As a warmer.',
      '**Language (5 min).** Board the prepositions and directions phrases from the *Secret places* sheet.',
      '**Guess the place (15 min).** Show the *Secret places* list. In pairs, A secretly chooses a place; B asks up to 10 yes/no questions. Then swap.',
      '**Big-screen directions (10 min).** Mark the school on Google Maps. A team chooses a secret place and gives directions from the school, step by step. You follow with the mouse; the other teams guess. Three or four rounds.',
      '**Real-world check (5 min).** "Which of these places have you been to? How did you get there?"',
    ]],
    ['Second half: Recommend a place, tourist and local role-play (S)', 'recommendations with *should*, *why don’t you...?*, comparatives and superlatives.', [
      '**Model (5 min).** You are a tourist; a strong student is a local. *"I love history. Where should I go?" "You should go to the Royal Pavilion. It’s the most beautiful building in Brighton."*',
      '**Language (5 min).** Board the locals’ useful language.',
      '**Role-play (20 min).** Half are tourists (role cards), half are locals. Tourists talk to three locals and note each recommendation. Swap halfway.',
      '**Decide and report (10 min).** *"I’m going to go to the i360 because the view is better than from the pier."*',
      '**Class awards (5 min).** Vote: best place for a rainy day, cheapest day out, most romantic place, best food.',
    ]],
  ]);
  day('Friday: Listening and Brighton words', 'The listening is a one-speaker voice message (script below). Read it aloud yourself, or paste it into Word and use **Review → Read Aloud**.', [
    ['First half: Listening, A Saturday in Brighton (L)(S)', 'listening for gist and detail; past simple and *there is / there are*.', [
      '**Lead-in (5 min).** "What can you do in Brighton on a Saturday?" Pairs list 5 things in 2 minutes.',
      '**Pre-teach (5 min).** *pebbles, seafront, seagull, pier, lane, view, tower, queue*.',
      '**First listen: gist (5 min).** Task 1: tick the places Lucas visits.',
      '**Second listen: detail (10 min).** Task 2: true or false. Pairs compare, then listen again for disagreements.',
      '**Third listen (5 min).** Task 3: gap fill.',
      '**Speaking (15 min).** Task 4: pairs talk about their own Saturday; each pair tells the class one surprising thing.',
    ]],
    ['Second half: Brighton bingo + Just a minute (S)', 'recycle the week’s vocabulary and speak for 30–60 seconds.', [
      '**Kahoot 4: A Saturday in Brighton (10 min).** Checks the listening and the new words.',
      '**Brighton bingo (10 min).** Students fill their grids from the word box. Read the definitions (answer key). First full line shouts *Brighton!*',
      '**Just a minute (20 min).** Groups of four with the topic cards. Speak for 30 seconds (60 for strong students); the others ask one question each. Rotate.',
      '**Feedback (5 min).** Each group’s best speaker talks in front of the class.',
    ]],
  ]);
  day('Monday: English culture and review', 'If Monday is a test day, move these to any free slot.', [
    ['First half: "Is it true in your country?" four corners (S)', 'giving opinions, agreeing and disagreeing, comparing cultures.', [
      '**Recall (5 min).** "What do you remember from Wednesday’s Kahoot about English culture?"',
      '**Language (5 min).** *I agree / I don’t agree because... In my country we... It’s the same / different in...*',
      '**Four corners (15 min).** Read a culture statement (answer key section). Students walk to a corner sign, then explain their choice.',
      '**Culture tips (15 min).** Groups write 5 tips for a new student: *"You should queue at the bus stop. Don’t feed the seagulls!"* Choose the top 10 for a wall poster.',
      '**Wrap-up (5 min).** "What’s the most surprising thing about England for you?"',
    ]],
    ['Second half: Brighton board game + Kahoot 5 (S)', 'review the week’s grammar while speaking.', [
      '**Set up (5 min).** Groups of 3–4 with a board, a dice and counters.',
      '**Play (25 min).** Do the task on your square and speak for 20+ seconds; the group asks one question. Can’t do it? Go back 2. Monitor and note errors.',
      '**Error clinic (5 min).** Board 5 errors; teams correct them.',
      '**Kahoot 5: Big week review (10 min).** Announce the week’s champion.',
    ]],
  ]);

  // Listening script
  t.push(H1('Listening script: A Saturday in Brighton'));
  t.push(P('Lucas leaves a voice message for his sister, Ana. About 90 seconds. Read it three times (gist, detail, gap fill).'));
  t.push(box(SCRIPT.map((p) => P(p, { run: { size: 24 } }))));
  t.push(H3('Online alternative (not checked: preview first)'));
  t.push(P('A Brighton seaside video from an ESL Printables A1–A2 worksheet:'));
  t.push(link('With subtitles: youtube.com/watch?v=rXf4_1Eex2c', 'https://www.youtube.com/watch?v=rXf4_1Eex2c'));
  t.push(link('Without subtitles: youtube.com/watch?v=CSUvZ9sJfYc', 'https://www.youtube.com/watch?v=CSUvZ9sJfYc'));

  // Answer keys
  t.push(H1('Answer keys'));
  t.push(H2('Tuesday: Brighton true or false'));
  t.push(table(TF_DAY1.map((s, i) => [String(i + 1), s, ['T', 'F: London is the capital.', "F: they're narrow and old.", 'T: Palace Pier and the West Pier.', 'T', "F: it's part of the city of Brighton and Hove.", 'T: the Brighton Festival.', 'F'][i]]), [500, 5400, 4006]));
  t.push(H2('Wednesday: postcard questions'));
  t.push(...bullets(['To her grandma.', 'She went to Palace Pier and ate fish and chips.', "It's sunny.", "She's going to visit the Royal Pavilion."]));
  t.push(P('**Order of parts:** greeting → where I am → what I did → what I can see now → plans → goodbye.'));
  t.push(H2('Friday: listening'));
  t.push(P('**Task 1:** the Lanes, the beach, the seafront, the pier, the i360. (Not: the Royal Pavilion, the Marina, the Booth Museum.)'));
  t.push(P('**Task 2:** 1 T · 2 F (her sister) · 3 F (£12) · 4 F (pebbles) · 5 T · 6 T · 7 T · 8 F (by bus)'));
  t.push(P('**Task 3:** 1 twelve · 2 one · 3 chips · 4 sea · 5 seven · 6 Pavilion'));
  t.push(H2('Friday: bingo definitions to read out'));
  t.push(table([
    ['Word', 'Definition'],
    ['pier', 'A long structure that goes out over the sea.'], ['pebbles', 'Small round stones on the beach.'],
    ['seagull', 'A white bird that steals your food.'], ['seafront', 'The road and path next to the beach.'],
    ['beach', 'The place by the sea where you sit and swim.'], ['tower', 'A very tall, thin building.'],
    ['view', 'What you can see from a high place.'], ['lane', 'A small, narrow street.'],
    ['palace', 'A very big, beautiful house for a king or queen.'], ['queue', 'A line of people waiting.'],
    ['shop', 'A place where you buy things.'], ['fish and chips', 'A typical English meal, often eaten by the sea.'],
    ['bus', 'Big public transport on the road.'], ['waves', 'The sea moves up and down in these.'],
    ['sunny', 'The weather when you need sunglasses.'], ['windy', 'The weather when your hat flies away.'],
  ], [2400, 7506], { header: true }));
  t.push(H2('Monday: culture statements (read aloud)'));
  t.push(...numbered(CULTURE));
  return t;
}

// ================= BUILD =================
async function build() {
  const studentDoc = new Document({
    creator: 'Kings Brighton', title: 'Brighton Skills Week: Student worksheets', styles, numbering,
    sections: [
      { properties: portrait, headers: { default: header('Brighton Skills Week · 2A Elementary · Student pack') }, footers: { default: footer() }, children: studentPack() },
      { properties: landscape, headers: { default: landscapeHeader('Brighton Skills Week · Monday board game') }, footers: { default: footer() }, children: boardGame() },
    ],
  });
  const teacherDoc = new Document({
    creator: 'Kings Brighton', title: 'Brighton Skills Week: Teacher pack', styles, numbering,
    sections: [{ properties: portrait, headers: { default: header('Brighton Skills Week · 2A Elementary · Teacher pack') }, footers: { default: footer() }, children: teacherPack() }],
  });
  fs.writeFileSync(path.join(OUT, 'Brighton Week - Student Pack.docx'), await Packer.toBuffer(studentDoc));
  fs.writeFileSync(path.join(OUT, 'Brighton Week - Teacher Pack.docx'), await Packer.toBuffer(teacherDoc));
  console.log('built');
}
build().catch((e) => { console.error(e); process.exit(1); });

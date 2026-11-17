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

// ================= BRIGHTON WEEK: LEVEL CONTENT =================
// Course 2 = A1, Course 3 = A2+. Shared: secret places, directions, Kahoots.
const KAHOOTS = [
  ['Day 1', 'Brighton basics', '739bef19-5282-47db-9c2d-305e02e691e5'],
  ['Day 2', 'English culture', '7d6971c3-87f7-42f0-8808-e6dd0b2d4467'],
  ['Day 3', 'Places in Brighton', '75dc390b-2070-49ef-965b-2cafad5249b4'],
  ['Day 4', 'A Saturday in Brighton', '190ddd16-7748-4bfc-a4ab-1f8f8a6b5771'],
  ['Day 5', 'Big week review', '2a9b54b3-568b-46df-8316-0f868e6fe796'],
];
const PLACES = [`The Royal Pavilion`, `Palace Pier`, `the i360`, `the Lanes`, `North Laine`, `the Clock Tower`,
  `Churchill Square`, `Brighton station`, `the Sea Life Centre`, `the West Pier`, `Brighton Marina`, `the beach`];
const DIRECTIONS = `go straight on · turn left / right · take the second left · it’s on your right · next to · opposite · between · on the corner of · at the end of`;
const CORNERS = ['AGREE', 'DISAGREE', 'NOT SURE', "IT'S THE SAME IN MY COUNTRY"];

const LEVELS = {
  c2: {
    course: 'Course 2', level: 'A1', file: 'Course 2 (A1)',
    grammar: `*be*, present simple, *can / can’t*, *there is / there are*, past simple (common verbs), *want to*`,
    fswIntro: `Ask *"Do you...?"*, *"Can you...?"* or *"Did you...?"*. Write a name. Then ask one more question: *Where? When? Who with?*`,
    fsw: [`likes fish and chips`, `lives in Hove`, `takes the bus to school`, `can swim`, `went to the beach this week`,
      `likes tea with milk`, `has a brother or a sister`, `can name a street in Brighton`, `took a photo of a seagull`, `wants to visit London`],
    tf: [
      [`The Royal Pavilion looks like a palace from India.`, `T`],
      [`Brighton is the capital of England.`, `F: London is.`],
      [`The Lanes are big, new streets.`, `F: they are small and old.`],
      [`Brighton has two piers. One is very old and broken.`, `T: Palace Pier and the West Pier.`],
      [`You can swim in the sea in Brighton.`, `T`],
      [`Brighton beach has a lot of sand.`, `F: it has pebbles (small stones).`],
      [`Brighton’s football team is called the Seagulls.`, `T`],
      [`It never rains in Brighton.`, `F`],
    ],
    interviewIntro: `Ask your partner. Write short notes. Then tell a new partner: *"Maria lives in Hove. She takes the bus..."*`,
    interview: [`Where do you live in Brighton?`, `Who do you live with?`, `How do you come to school?`,
      `What is your favourite place in Brighton?`, `What can you see from your window?`, `Do you like English food?`,
      `What did you do last weekend?`, `What do you want to do this weekend?`],
    writing: { genre: 'postcard', title: 'A postcard from Brighton', words: '40–60',
      model: [`Dear Grandma,`, `I’m in Brighton! I live with a nice family in Hove. On Saturday I went to the pier and I ate fish and chips. Today it’s sunny and I’m on the beach. I can see the sea. I love it here!`, `Love,`, `Sofia`],
      questions: [[`Who is the postcard to?`, `Her grandma.`], [`Where does Sofia live?`, `With a family in Hove.`], [`What did she eat?`, `Fish and chips.`], [`What’s the weather like?`, `It’s sunny.`]],
      parts: [`___ today`, `___ goodbye`, `___ greeting`, `___ what I did`, `___ where I live`],
      partsKey: `greeting → where I live → what I did → today → goodbye`,
      underline: `**past simple** (went) · **present continuous** (I’m...ing) · **can**`,
      plan: [`Where I live / who with`, `One thing I did`, `Today: weather / what I can see`],
      opener: `Dear ______________________,`, lines: 9,
      checklist: [`Greeting (*Dear... / Hi...*)`, `Where I live`, `One thing I did (*I went... I ate...*)`, `Today (*It’s sunny... I can see...*)`, `Goodbye (*Love, / See you soon,*)`, `40–60 words`],
    },
    tourists: [`You like **old buildings**.`, `You like **shopping**. You have **£10**.`, `It’s **raining**. You have a **child**.`,
      `You want to take **photos** of Brighton.`, `You like **walking** and **nature**.`, `You are **hungry**. You want **English food**.`],
    locals: [`Go to...`, `It’s good for...`, `It’s near...`, `It’s cheap / big / beautiful / fun.`, `You can... there.`],
    questionStarters: [`Is it near...? Is it next to...?`, `Can you... there?`, `Is it big / old?`, `Is there a... there?`],
    listening: {
      script: [
        `Hi Ana! It’s Lucas. On Saturday I went to the city centre with my friend Yuki. We walked in the Lanes. There are lots of small shops. Yuki bought a ring for her sister. It was twelve pounds.`,
        `At one o’clock we went to the beach. There is no sand! There are only pebbles. We ate fish and chips. A seagull took one of my chips!`,
        `Then we went on the pier. We played games and Yuki won a small toy. In the afternoon we went up the i360. The view was great. We saw all of Brighton and the sea. It was the best thing all day.`,
        `At seven o’clock we went home on the bus. Next weekend I want to visit the Royal Pavilion. Bye! Call me soon!`,
      ],
      length: 'about 60 seconds',
      places: [`the Lanes`, `the beach`, `the Royal Pavilion`, `the pier`, `the Marina`, `the i360`, `the Booth Museum`, `the station`],
      placesKey: `the Lanes, the beach, the pier, the i360. (Not: the Royal Pavilion, the Marina, the Booth Museum, the station.)`,
      tf: [[`Lucas went to the city centre with Yuki.`, `T`], [`Yuki bought a ring for her mother.`, `F (her sister)`],
        [`The ring was £20.`, `F (£12)`], [`Brighton beach has a lot of sand.`, `F (pebbles)`],
        [`A seagull took one of Lucas’s chips.`, `T`], [`Yuki won a toy on the pier.`, `T`],
        [`The i360 was the best thing all day.`, `T`], [`They went home by train.`, `F (by bus)`]],
      task3title: 'Listen again and complete.',
      task3: [[`It was ______________ pounds.`, `twelve`], [`At ______________ o’clock we went to the beach.`, `one`],
        [`We ate fish and ______________.`, `chips`], [`We saw all of Brighton and the ______________.`, `sea`],
        [`At ______________ o’clock we went home on the bus.`, `seven`], [`Next weekend I want to visit the Royal ______________.`, `Pavilion`]],
      task3lines: false,
      talk: `Tell your partner about your Saturday. Where did you go? What did you eat?`,
    },
    bingo: [[`pier`, `A long road that goes out over the sea.`], [`pebbles`, `Small round stones on the beach.`],
      [`seagull`, `A white bird. It takes your food!`], [`beach`, `You sit here next to the sea.`],
      [`sea`, `Lots of salty water.`], [`tower`, `A very tall, thin building.`], [`shop`, `You buy things here.`],
      [`bus`, `A big red or blue car for lots of people.`], [`ring`, `You wear it on your finger.`],
      [`palace`, `A very big house for a king or queen.`], [`fish and chips`, `A typical English meal.`],
      [`sunny`, `The weather when you need sunglasses.`], [`windy`, `The weather when your hat flies away.`],
      [`cold`, `Not hot.`], [`photo`, `A picture you take with your phone.`], [`queue`, `A line of people waiting.`]],
    jamSecs: 30, jamChallenge: 45,
    jam: [`My family`, `My favourite food`, `My home town`, `The weather today`, `My school`, `My bedroom`,
      `My best friend`, `Brighton beach`, `My hobbies`, `My phone`, `My weekend`, `My favourite place in Brighton`],
    culture: [`People in England drink a lot of tea.`, `People in England say "sorry" a lot.`,
      `People in England talk about the weather a lot.`, `In England, people eat dinner early.`,
      `Students in England wear a uniform.`, `English food is good.`,
      `People in England queue (wait in a line).`, `Fish and chips is the best English food.`],
    board: [`START`, `Say 3 things you can do in Brighton.`, `What’s your name? Spell it.`, `What did you eat today?`,
      `Describe your bedroom: There is... There are...`, `Name 5 colours you can see.`, `Go forward 2!`,
      `What’s your favourite food? Why?`, `What can you do well?`, `Count from 20 to 30.`, `What did you do yesterday?`,
      `Ask the player on your left 2 questions.`, `What’s the weather like today?`, `Go back 3!`, `What time do you get up?`,
      `Name 5 things on the beach.`, `Describe your best friend.`, `What’s in your bag?`, `Miss a turn!`,
      `Talk about your family.`, `What do you want to buy in Brighton?`, `Name 3 places in Brighton.`, `Go forward 1!`,
      `What’s your favourite place in Brighton?`, `Do you like fish and chips? Why?`, `What’s your favourite day of the week?`,
      `Say 3 things you did this week.`, `What do you do at the weekend?`, `What’s your favourite English word?`,
      `FINISH! Say "Thank you, Brighton!"`],
    boardSecs: 10,
    cultureLang: `*I agree. / I don’t agree. / In my country... / It’s the same. / It’s different.*`,
    tipsExample: `*"Queue at the bus stop. Don’t feed the seagulls!"*`,
  },

  c3: {
    course: 'Course 3', level: 'A2+', file: 'Course 3 (A2+)',
    grammar: `present perfect (*ever / never / for / since*), past simple and past continuous, comparatives and superlatives, *going to* and *will*, *should*, first and second conditional`,
    fswIntro: `Ask *"Have you ever...?"*, *"Were you...?"* or *"Would you...?"*. Write a name, then ask two follow-up questions: *When? Why? What was it like?*`,
    fsw: [`has been on Palace Pier`, `has swum in the sea in Brighton`, `has tried a Sunday roast`, `has been inside the Royal Pavilion`,
      `has lost something since arriving in Brighton`, `was born in a capital city`, `has had a conversation with a stranger here`,
      `is going to travel somewhere else in the UK`, `would like to live in Brighton`, `has noticed something strange about English people`],
    tf: [
      [`The Royal Pavilion was built for a prince who later became king.`, `T: King George IV.`],
      [`Brighton has always been a city.`, `F: Brighton and Hove became a city in 2000.`],
      [`The Lanes are wide, modern streets.`, `F: they’re narrow and old.`],
      [`Brighton has two piers, but one is in ruins.`, `T: Palace Pier and the West Pier.`],
      [`Palace Pier opened in the 1990s.`, `F: it opened in 1899.`],
      [`Hove is part of the same city as Brighton.`, `T: the city of Brighton and Hove.`],
      [`There’s a big arts festival in Brighton every May.`, `T: the Brighton Festival.`],
      [`Brighton’s football team is called the Pigeons.`, `F: the Seagulls.`],
    ],
    interviewIntro: `Interview your partner and make notes. Then report to a new group, using the third person and at least one comparison: *"Maria has been here for two weeks. Her journey is longer than at home..."*`,
    interview: [`How long have you been in Brighton? Where are you staying?`, `How is your journey to school different from at home?`,
      `What did you do last weekend? Would you recommend it?`, `What’s the most surprising thing you’ve noticed about Brighton?`,
      `What have you found difficult so far?`, `If you could change one thing about Brighton, what would it be?`,
      `What English food have you tried? Would you eat it again?`, `What are you going to do before you go home?`],
    writing: { genre: 'email', title: 'An email home', words: '100–130',
      model: [`Hi Grandma,`,
        `I’ve been in Brighton for a week now and I’m having a great time. I’m staying with a friendly host family in Hove, and they’ve already taught me how to make a proper cup of tea!`,
        `Last Saturday I went to Palace Pier with some friends from school. We were eating fish and chips when a seagull stole my food. Everyone thought it was hilarious. The weather has been better than I expected, although it’s quite windy.`,
        `Right now I’m sitting in a café in North Laine, which is much more relaxed than the city centre. Next weekend I’m going to visit the Royal Pavilion, and if it’s sunny, we might walk along the cliffs at the Seven Sisters.`,
        `I miss you all! Write back soon.`, `Love,`, `Sofia`],
      questions: [[`How long has Sofia been in Brighton?`, `A week.`], [`What happened at the pier?`, `A seagull stole her food while she was eating fish and chips.`],
        [`How is North Laine different from the city centre?`, `It’s much more relaxed.`], [`What will Sofia do if it’s sunny next weekend?`, `Walk along the cliffs at the Seven Sisters.`]],
      parts: [`___ a story from last weekend`, `___ plans`, `___ greeting`, `___ how long I’ve been here and where I’m staying`, `___ where I am now`, `___ goodbye`],
      partsKey: `greeting → how long I’ve been here and where I’m staying → a story from last weekend → where I am now → plans → goodbye`,
      underline: `**present perfect** · **past continuous + past simple** (*were eating... when... stole*) · **a comparative** · **an if-sentence** · linkers: **although, which**`,
      plan: [`How long I’ve been here / where I’m staying`, `A story from last weekend`, `Where I am now / a comparison`, `My plans (going to / if...)`],
      opener: `Hi ______________________,`, lines: 10,
      checklist: [`Greeting`, `How long I’ve been here (present perfect)`, `A short story (past continuous + past simple)`,
        `A comparison (*better than... / more relaxed than...*)`, `A plan (*going to* / *if...*)`, `At least 2 linkers (*although, because, which, so*)`,
        `Closing (*Write back soon! / Love,*)`, `100–130 words`],
    },
    tourists: [`You love **history**, but you hate crowds. You have one afternoon.`,
      `You have **£10** for the whole day and want something unusual to take home.`,
      `It’s **raining**, you’re with a **7-year-old**, and you need to be back by 4 pm.`,
      `You want **amazing photos**, but you’re scared of heights.`,
      `You love **nature and walking**. You have a whole day and a bus pass.`,
      `You want **typical English food**, but you’re vegetarian.`],
    locals: [`If I were you, I’d...`, `You could... or... ; ... is better because...`, `It’s the best place for..., although...`,
      `Have you thought about...?`, `Make sure you... / Don’t forget to...`],
    questionStarters: [`Is it near...? Is it opposite...?`, `Can you... there? Have you been there?`, `Was it built before / after...?`, `Is it bigger / more famous than...?`],
    listening: {
      script: [
        `Hi Ana, it’s Lucas! I’ve had such a busy weekend, so I wanted to tell you about it. On Saturday morning I went into the city centre with my friend Yuki, who’s from Japan. First we walked around the Lanes, which are full of tiny old streets and jewellery shops. Yuki bought a silver ring for her sister. It only cost twelve pounds, which was cheaper than she expected.`,
        `We got to the beach at about one o’clock. I was expecting sand, but there isn’t any, just pebbles! It was really uncomfortable to sit on. For lunch we had fish and chips on the seafront. While I was talking to Yuki, a seagull flew down and took one of my chips! Everyone around us laughed.`,
        `After lunch we went on the pier. We played games in the arcade and Yuki won a small toy. Then in the afternoon we went up the i360. The view was amazing: we could see the whole of Brighton and the sea, and even the hills behind the city. It was a bit expensive, but it was definitely the best thing we did all day.`,
        `By the evening we were exhausted, so we caught the bus home at seven o’clock. Next weekend I’m going to visit the Royal Pavilion. I walk past it every day on my way to school, but I haven’t been inside yet. Anyway, call me soon and tell me your news!`,
      ],
      length: 'about 2 minutes',
      places: [`the Lanes`, `the beach`, `the seafront`, `the Royal Pavilion`, `the pier`, `the Marina`, `the i360`, `the Booth Museum`],
      placesKey: `the Lanes, the beach, the seafront, the pier, the i360. (Not: the Royal Pavilion, the Marina, the Booth Museum.)`,
      tf: [[`Yuki is from Japan.`, `T`], [`Yuki bought a ring for her mother.`, `F (her sister)`],
        [`The ring was more expensive than Yuki expected.`, `F (cheaper)`], [`Lucas expected the beach to be sandy.`, `T`],
        [`Lucas was talking to Yuki when the seagull took his chip.`, `T`], [`Lucas won a toy on the pier.`, `F (Yuki won it)`],
        [`From the i360 they could see hills behind the city.`, `T`], [`Lucas has been inside the Royal Pavilion.`, `F (not yet)`]],
      task3title: 'Listen again and answer the questions.',
      task3: [[`Why was the beach uncomfortable?`, `Because there are only pebbles, no sand.`],
        [`What did people do when the seagull took the chip?`, `They laughed.`],
        [`What did Lucas think about the price of the i360?`, `It was a bit expensive, but worth it (the best thing all day).`],
        [`How did Lucas and Yuki feel in the evening?`, `Exhausted (very tired).`],
        [`Why does Lucas walk past the Royal Pavilion every day?`, `It’s on his way to school.`]],
      task3lines: true,
      talk: `Tell your partner about the best day you’ve had in Brighton so far. Was it better or worse than Lucas’s Saturday? Why?`,
    },
    bingo: [[`promenade`, `A wide path next to the sea for walking.`], [`pebbles`, `Small round stones on the beach.`],
      [`seagull`, `A noisy white bird that steals your food.`], [`seafront`, `The road and path along the beach.`],
      [`pier`, `A long structure that goes out over the sea.`], [`arcade`, `A place with games machines.`],
      [`view`, `What you can see from a high place.`], [`lane`, `A small, narrow street.`],
      [`palace`, `A very big, beautiful house for a king or queen.`], [`queue`, `A line of people waiting.`],
      [`independent shop`, `A small shop that isn’t part of a big chain.`], [`tide`, `The sea coming in and going out.`],
      [`waves`, `The sea moves up and down in these.`], [`breezy`, `A little windy.`],
      [`crowded`, `Full of people.`], [`souvenir`, `Something you buy to remember a place.`]],
    jamSecs: 60, jamChallenge: 90,
    jam: [`The best and worst things about Brighton`, `A time I got lost`, `My home town vs Brighton`, `A tradition from my country`,
      `The perfect weekend in Brighton`, `Something I’ve learned this week`, `Tourists: good or bad for a city?`,
      `Living with a host family`, `Social media`, `A person I admire`, `My plans for the future`, `Why I’m learning English`],
    culture: [`Queuing is a sign of a fair society.`, `Saying "sorry" all the time makes it less meaningful.`,
      `It’s rude to ask someone’s age or salary.`, `Small talk about the weather is a waste of time.`,
      `School uniforms are a good idea.`, `You should always take your shoes off in someone’s home.`,
      `Tourists should learn some of the local language.`, `Fast food is ruining traditional food.`],
    board: [`START`, `What have you done this week that you’ve never done before?`, `If you won £1,000 tomorrow, what would you do in Brighton?`,
      `Give 3 differences between Brighton and your home town.`, `What were you doing at 8 pm yesterday?`,
      `What’s the most beautiful place you’ve ever visited?`, `Go forward 2!`, `What should a new student do in their first week? Why?`,
      `Describe an English habit and give your opinion.`, `What are you going to do this weekend? Give details.`,
      `Tell a story about a time something went wrong.`, `Ask the player on your left 3 questions about their weekend.`,
      `How long have you been learning English?`, `Go back 3!`, `What’s the best and worst thing about living with other people?`,
      `Explain the rules of a game from your country.`, `Would you rather live by the sea or in the mountains? Why?`,
      `Describe the best meal you’ve had in England.`, `What will you miss about Brighton?`, `Miss a turn!`,
      `If you could meet anyone, who would it be? Why?`, `Give directions from school to the beach.`,
      `What’s the strangest thing you’ve seen in Brighton?`, `Go forward 1!`, `Persuade the group to visit your country.`,
      `Talk for 30 seconds about your favourite film.`, `Tell a story about a seagull (true or invented!).`,
      `What’s the most useful thing you’ve learned this week?`, `How has your English improved since you arrived?`,
      `FINISH! Say "Thank you, Brighton!"`],
    boardSecs: 30,
    cultureLang: `*I completely agree because... / I’m not sure about that... / In my country, it’s more common to... / That depends on...*`,
    tipsExample: `*"You should queue at the bus stop, because people get annoyed if you don’t. Never feed the seagulls!"*`,
  },
};
// ================= BRIGHTON WEEK: PACKS (one per level) =================
function studentPack(L) {
  const s = [];
  const W = L.writing;
  s.push(new Paragraph({ heading: 'Heading1', children: [new TextRun(`Brighton Skills Week: Student worksheets`)] }));
  s.push(P(`**${L.course} (${L.level}).** Print one copy per student unless the table says otherwise.`));
  s.push(table([
    ['Page', 'Worksheet', 'Day', 'Copies'],
    ['2', 'Find someone who...', 'Day 1', '1 per student'],
    ['3', 'Brighton: true or false?', 'Day 1', '1 per team'],
    ['4', 'My first week: interview', 'Day 1', '1 per student'],
    ['5', `${W.title} (model)`, 'Day 2', '1 per student'],
    ['6', `My ${W.genre} (writing frame)`, 'Day 2', '1 per student'],
    ['7', 'Secret places (or show on screen)', 'Day 3', '1 per pair'],
    ['8', 'Tourist role cards (cut up)', 'Day 3', '1 set per 6 students'],
    ['9', 'Listening: A Saturday in Brighton', 'Day 4', '1 per student'],
    ['10', 'Brighton bingo', 'Day 4', '1 per 2 students (cut in half)'],
    ['11', 'Just a minute topic cards (cut up)', 'Day 4', '1 set per group'],
    ['12–15', 'Four corners signs', 'Day 5', '1 set for the room'],
    ['16', 'Brighton board game (A4, or enlarge to A3)', 'Day 5', '1 per group'],
  ], [1100, 4900, 1100, 2806], { header: true }));

  s.push(H1('Find someone who...'));
  s.push(NameLine());
  s.push(P(L.fswIntro));
  s.push(table([['Find someone who...', 'Name', 'Extra information'], ...L.fsw.map((f) => [f, '', ''])], [4200, 2200, 3506], { header: true, rowHeight: 700 }));

  s.push(H1('Brighton: true or false?'));
  s.push(P('Team name: ______________________'));
  s.push(P('Listen to your teacher. Write **T** (true) or **F** (false). If it’s false, correct it for a bonus point!'));
  s.push(table([['#', 'Sentence', 'T / F', 'Correction'], ...L.tf.map(([t], i) => [String(i + 1), t, '', ''])], [500, 4900, 900, 3606], { header: true, rowHeight: 650 }));

  s.push(H1('My first week in Brighton: interview'));
  s.push(NameLine());
  s.push(P(L.interviewIntro));
  s.push(P("My partner's name: ______________________"));
  s.push(table([['Question', 'Notes'], ...L.interview.map((q) => [q, ''])], [5000, 4906], { header: true, rowHeight: 900 }));

  s.push(H1(W.title));
  s.push(NameLine());
  s.push(H3(`1. Read the ${W.genre}`));
  s.push(box(W.model.map((l, i) => P(l, { run: { size: 22 }, spacing: { after: 100 } }))));
  s.push(H3('2. Answer the questions'));
  W.questions.forEach(([q], i) => { s.push(P(`${i + 1}. ${q}`, { spacing: { after: 0 } })); s.push(WriteLine()); s.push(spacer()); });
  s.push(H3('3. Find the parts'));
  s.push(P(`Number the parts of the ${W.genre} in order:`));
  const parts = W.parts.slice();
  const rows = [];
  for (let i = 0; i < parts.length; i += 3) { const r = parts.slice(i, i + 3); while (r.length < 3) r.push(''); rows.push(r); }
  s.push(table(rows, [3302, 3302, 3302]));
  s.push(spacer());
  s.push(P(`Underline one example of: ${W.underline}`));

  s.push(H1(`My ${W.genre}`));
  s.push(P('**Plan.** Make notes for each part.'));
  s.push(table(W.plan.map((p) => [p, '']), [3400, 6506], { rowHeight: 600 }));
  s.push(spacer());
  s.push(P(`**Write** ${W.words} words.${W.genre === 'postcard' ? ' Write your name on the back only!' : ' Continue on the back if you need more space. Write your name at the end.'}`));
  s.push(box([P(W.opener, { spacing: { after: 0 } }), ...lines(W.lines)], PORTRAIT_W, 'FFFFFF'));
  s.push(H3('Check with a partner'));
  s.push(...checks(W.checklist));

  s.push(H1('Secret places'));
  s.push(P('Choose a secret place. Your partner asks up to 10 yes/no questions to guess it.'));
  s.push(cards(PLACES, 3, 1000, 28));
  s.push(spacer());
  s.push(box([P('**Question starters**'), ...bullets(L.questionStarters)]));
  s.push(spacer());
  s.push(box([P('**Directions**'), P(DIRECTIONS)]));

  s.push(H1('Tourist role cards'));
  s.push(Small('Cut along the dashed lines. One card per tourist.'));
  s.push(cards(L.tourists.map((t) => `**TOURIST** — ${t}`), 2, 1700, 24));
  s.push(spacer());
  s.push(box([P('**Locals: useful language**'), ...bullets(L.locals)]));

  const LS = L.listening;
  s.push(H1('Listening: A Saturday in Brighton'));
  s.push(NameLine());
  s.push(P('Lucas leaves a voice message for his sister, Ana.'));
  s.push(H3('Task 1. Tick (✓) the places Lucas visited.'));
  s.push(table([LS.places.slice(0, 4).map((p) => `☐  ${p}`), LS.places.slice(4).map((p) => `☐  ${p}`)], [2476, 2476, 2477, 2477]));
  s.push(H3('Task 2. True or false?'));
  s.push(table([['#', 'Sentence', 'T / F'], ...LS.tf.map(([t], i) => [String(i + 1), t, ''])], [500, 8006, 1400], { header: true }));
  s.push(H3(`Task 3. ${LS.task3title}`));
  if (LS.task3lines) LS.task3.forEach(([q], i) => { s.push(P(`${i + 1}. ${q}`, { spacing: { after: 0 } })); s.push(WriteLine()); });
  else s.push(...numbered(LS.task3.map(([q]) => q)));
  s.push(H3('Task 4. Talk to your partner.'));
  s.push(P(LS.talk));

  s.push(H1('Brighton bingo'));
  s.push(P('Choose 9 words from the box and write them in your grid. Listen to your teacher. Cross out the words. Get a full line? Shout **"Brighton!"**'));
  s.push(box([P(L.bingo.map((b) => b[0]).join('  ·  '), { align: AlignmentType.CENTER, run: { size: 24 } })]));
  const grid = () => new Table({
    width: { size: 6300, type: WidthType.DXA }, columnWidths: [2100, 2100, 2100], alignment: AlignmentType.CENTER,
    rows: [0, 1, 2].map(() => new TableRow({ height: { value: 1300, rule: HeightRule.EXACT }, children: [0, 1, 2].map(() => cell('', 2100, { border: { style: BorderStyle.SINGLE, size: 12, color: PURPLE } })) })),
  });
  s.push(spacer()); s.push(grid());
  s.push(new Paragraph({ spacing: { before: 300, after: 300 }, border: { bottom: dashed }, children: [new TextRun({ text: '✂  cut here', size: 16, color: '888888' })] }));
  s.push(P('Choose 9 words. Write them in your grid.'));
  s.push(grid());

  s.push(H1('Just a minute: topic cards'));
  s.push(Small(`Cut up one set per group. Speak for ${L.jamSecs} seconds without stopping (${L.jamChallenge} seconds for a challenge!).`));
  s.push(cards(L.jam, 3, 1600, 24));

  CORNERS.forEach((c) => {
    s.push(new Paragraph({ pageBreakBefore: true, spacing: { before: 3600 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: c, bold: true, size: c.length > 12 ? 96 : 140, color: PURPLE })] }));
  });
  return s;
}

function boardGame(L) {
  const COLS = 6, ROWS = 5;
  const w = Math.floor(LANDSCAPE_W / COLS);
  const rows = [];
  for (let r = 0; r < ROWS; r++) {
    const idx = Array.from({ length: COLS }, (_, c) => r * COLS + c);
    if (r % 2 === 1) idx.reverse();
    rows.push(new TableRow({
      height: { value: 1700, rule: HeightRule.EXACT }, cantSplit: true,
      children: idx.map((i) => {
        const special = i === 0 || i === 29;
        const fun = /^(Go|Miss)/.test(L.board[i]);
        return cell([
          new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: String(i + 1), bold: true, size: 28, color: special ? 'FFFFFF' : PURPLE })] }),
          new Paragraph({ children: [new TextRun({ text: L.board[i], size: special ? 24 : 17, bold: special || fun, color: special ? 'FFFFFF' : '222222' })] }),
        ], w, { fill: special ? PURPLE : fun ? 'FBE9F1' : (r + idx.indexOf(i)) % 2 ? 'FFFFFF' : LIGHT, border: { style: BorderStyle.SINGLE, size: 8, color: PURPLE } });
      }),
    }));
  }
  return [
    new Paragraph({ heading: 'Heading1', children: [new TextRun('Brighton board game')] }),
    P(`Roll the dice and move. Do the task and speak for at least ${L.boardSecs} seconds. The group asks you one question. Can’t do it? Go back 2. Follow the numbers: the path snakes back and forth.`, { run: { size: 18 } }),
    new Table({ width: { size: w * COLS, type: WidthType.DXA }, columnWidths: Array(COLS).fill(w), rows }),
  ];
}

function teacherPack(L) {
  const t = [];
  const W = L.writing, LS = L.listening;
  t.push(new Paragraph({ heading: 'Heading1', children: [new TextRun(`Brighton Skills Week: Teacher pack`)] }));
  t.push(P(`**${L.course} (${L.level})** · **Length:** five days, one 90-minute skills lesson a day, taught as two halves · **Grammar review:** ${L.grammar}.`));
  t.push(P(`Mostly structured speaking, plus one writing lesson (Day 2), one listening lesson (Day 4) and a Kahoot every day. Student worksheets are in the separate student pack. The other course has its own pack with the same plan at a different level; the Kahoots are shared.`));
  t.push(H2('The week at a glance'));
  t.push(table([
    ['Day', 'First half', 'Second half', 'Kahoot'],
    ['Day 1', 'Find someone who + Brighton true or false (S)', 'My first week: interview and report back (S)', 'Brighton basics'],
    ['Day 2', `Writing: ${W.title.toLowerCase()} (R)(W)`, `${W.genre === 'postcard' ? 'Postcard' : 'Email'} gallery (R)(S)`, 'English culture'],
    ['Day 3', 'Guess the place + big-screen directions (S)(L). Overflow for writing', 'Recommend a place: tourist and local role-play (S)', 'Places in Brighton'],
    ['Day 4', 'Listening: A Saturday in Brighton (L)(S)', 'Brighton bingo + Just a minute (S)', 'A Saturday in Brighton'],
    ['Day 5', '"Is it true in your country?" four corners (S)', 'Brighton board game: grammar review (S)', 'Big week review'],
  ], [900, 3600, 3400, 2006], { header: true }));
  t.push(H2('Kahoots (shared by both courses)'));
  t.push(P('These links open each kahoot’s page, where you press **Start** to host it or **Duplicate** to copy it into your own account. They work once the kahoots are set to public.'));
  t.push(new Table({
    width: { size: PORTRAIT_W, type: WidthType.DXA }, columnWidths: [1200, 3400, 5306],
    rows: [
      new TableRow({ tableHeader: true, children: ['Day', 'Kahoot', 'Link'].map((h, i) => cell(h, [1200, 3400, 5306][i], { fill: PURPLE, run: { bold: true, color: 'FFFFFF' } })) }),
      ...KAHOOTS.map(([d, n, u]) => new TableRow({ children: [cell(d, 1200), cell(`Brighton ${d}: ${n}`, 3400), cell([new Paragraph({ children: [new ExternalHyperlink({ link: `https://create.kahoot.it/details/${u}`, children: [new TextRun({ text: 'Open the kahoot', style: 'Hyperlink' })] })] })], 5306)] })),
    ],
  }));
  t.push(H2('Prep checklist'));
  t.push(...checks([
    'Print the student pack (copies listed on its first page).',
    'Cut up: tourist role cards, Just a minute cards, bingo sheets.',
    'Print the board game, one per group (A3 if you can). Bring dice and counters.',
    'Day 3: open Google Maps on central Brighton on the classroom screen. Nothing to print.',
    'Day 4: read the listening script aloud, or paste it into Word and use Review → Read Aloud. Nothing to record.',
    'Day 5: stick the four corners signs in the corners of the room.',
  ]));

  const day = (title, intro, lessons) => {
    t.push(H1(title));
    t.push(P(intro));
    lessons.forEach(([h, aim, steps, extra]) => {
      t.push(H2(h)); t.push(P(`**Aim:** ${aim}`)); t.push(...numbered(steps)); if (extra) t.push(P(extra));
    });
  };
  const a1 = L.level === 'A1';
  day('Day 1: Welcome to Brighton', 'Students find out what they already know about Brighton and talk about their first days here.', [
    ['First half: Find someone who + Brighton true or false (S)', a1 ? 'simple questions and answers about life in Brighton.' : 'questions with the present perfect and follow-up questions.', [
      '**Kahoot: Brighton basics (10 min).** Run it cold as a warmer. Note which facts surprise them.',
      `**Find someone who (15 min).** Students mingle and ask the questions. ${a1 ? 'Drill the question forms first: *Do you...? Can you...? Did you...?*' : 'Insist on two follow-up questions each time.'}`,
      '**Feedback (5 min).** "Who found someone who...? Tell us about it." Board useful words.',
      '**Brighton true or false (10 min).** Teams of three. Read the 8 statements; teams write T or F. A point per right answer, a bonus point for correcting a false one.',
      '**Say it right (5 min).** Drill *Pavilion, pier, the Lanes, Hove, seafront, i360*. Students mark the stress.',
    ]],
    ['Second half: My first week, interview and report back (S)', a1 ? 'present simple and past simple about real life; *he / she* + *-s*.' : 'present perfect, comparatives and conditionals in a personal interview, then reporting in the third person.', [
      '**Lead-in (5 min).** Board: *Brighton: one word*. Each student says one word about Brighton so far.',
      '**Interview (15 min).** Pairs use the interview sheet. A asks all 8 questions and makes notes; then swap.',
      `**Report back (15 min).** New groups of four. Each student tells the group about their partner. ${a1 ? 'Model first on the board: *She lives... She likes...*' : 'Each report must include one comparison.'} Listeners ask one question each.`,
      '**Class survey (5 min).** "What’s the best place in Brighton so far?" Tally on the board.',
      '**Error clinic (5 min).** Board 5 errors you heard (no names). Teams correct them.',
    ]],
  ]);
  day(`Day 2: ${W.title}`, `Students write a ${W.words} word ${W.genre} home, then read each other’s and guess the writer. If you run out of time, finish the writing in Day 3’s first half and do the gallery straight after; move Day 3’s places lesson to Day 4’s second half.`, [
    [`First half: Writing ${W.title.toLowerCase()} (R)(W)`, a1 ? 'a short postcard with past simple, present continuous and *can*.' : 'an informal email with the present perfect, past continuous, a comparison, a plan and linkers.', [
      '**Lead-in (5 min).** "Who do you send messages to at home? What do you tell them?"',
      `**Read the model (10 min).** Students read Sofia’s ${W.genre} and answer the four questions.`,
      '**Notice the structure (5 min).** Number the parts; underline the language examples.',
      '**Plan (5 min).** Students make notes in the plan boxes.',
      `**Write (15 min).** ${W.words} words in the writing frame. Name on the back only.`,
      '**Check (5 min).** Swap with a partner and use the checklist.',
    ]],
    [`Second half: ${W.genre === 'postcard' ? 'Postcard' : 'Email'} gallery + Kahoot (R)(S)`, 'reading for specific information and talking about the week.', [
      `**Gallery (15 min).** Number the ${W.genre}s and stick them on the walls, names hidden. Students read them all and write who they think wrote each one.`,
      '**Guess who (10 min).** Go through the numbers. *"I think it’s Ahmed because he went to the football."* The writer reveals.',
      '**Hot seat (10 min).** 3 or 4 authors answer questions about their week.',
      '**Kahoot: English culture (10 min).** A closer that sets up Day 5’s culture lesson.',
    ], a1 ? '**Homework:** write a message to a friend at home about Brighton, 40–60 words.' : '**Homework:** write a real email to someone at home about Brighton, 120–150 words.'],
  ]);
  day('Day 3: Brighton places', 'Students describe where places are, give directions and recommend places to a visitor. **You need:** the classroom screen with Google Maps open on central Brighton. Nothing to print except the role cards.', [
    ['First half: Guess the place + big-screen directions (S)(L)', 'prepositions of place and simple directions.', [
      '**Kahoot: Places in Brighton (10 min).** As a warmer.',
      '**Language (5 min).** Board the directions phrases from the *Secret places* sheet.',
      '**Guess the place (15 min).** Show the *Secret places* list. In pairs, A secretly chooses a place; B asks up to 10 yes/no questions. Then swap.',
      '**Big-screen directions (10 min).** Mark the school on Google Maps. A team chooses a secret place and gives directions from the school, step by step. You follow with the mouse; the other teams guess.',
      '**Real-world check (5 min).** "Which of these places have you been to? How did you get there?"',
    ]],
    ['Second half: Recommend a place, tourist and local role-play (S)', a1 ? 'simple recommendations: *Go to... It’s good for... You can...*' : 'recommendations with *if I were you*, comparatives and superlatives, and dealing with a problem.', [
      `**Model (5 min).** You are a tourist; a strong student is a local. ${a1 ? '*"I like old buildings." "Go to the Royal Pavilion. It’s beautiful."*' : '*"I love history but hate crowds." "If I were you, I’d go to the Royal Pavilion early in the morning."*'}`,
      '**Language (5 min).** Board the locals’ useful language.',
      '**Role-play (20 min).** Half are tourists (role cards), half are locals. Tourists talk to three locals and note each recommendation. Swap halfway.',
      `**Decide and report (10 min).** ${a1 ? '*"I want to go to the i360. It’s fun."*' : '*"I’m going to go to the Booth Museum because it’s quieter than the Pavilion."*'}`,
      '**Class awards (5 min).** Vote: best place for a rainy day, cheapest day out, most romantic place, best food.',
    ]],
  ]);
  day('Day 4: Listening and Brighton words', `The listening is a one-speaker voice message (script below, ${LS.length}). Read it aloud yourself, or paste it into Word and use **Review → Read Aloud**.`, [
    ['First half: Listening, A Saturday in Brighton (L)(S)', a1 ? 'listening for gist and detail; past simple and *there is / there are*.' : 'listening for gist, detail and inference; narrative tenses.', [
      '**Lead-in (5 min).** "What can you do in Brighton on a Saturday?" Pairs list 5 things in 2 minutes.',
      `**Pre-teach (5 min).** ${a1 ? '*pebbles, seagull, pier, ring, view, toy*' : '*pebbles, seafront, arcade, exhausted, view, expect*'}.`,
      '**First listen: gist (5 min).** Task 1: tick the places Lucas visits.',
      '**Second listen: detail (10 min).** Task 2: true or false. Pairs compare, then listen again for disagreements.',
      `**Third listen (5 min).** Task 3: ${a1 ? 'gap fill' : 'answer the questions'}.`,
      '**Speaking (15 min).** Task 4 in pairs; each pair tells the class one surprising thing.',
    ]],
    ['Second half: Brighton bingo + Just a minute (S)', `recycle the week’s vocabulary and speak for ${L.jamSecs} seconds.`, [
      '**Kahoot: A Saturday in Brighton (10 min).** Checks the listening and the new words.',
      '**Brighton bingo (10 min).** Students fill their grids from the word box. Read the definitions (answer key). First full line shouts *Brighton!*',
      `**Just a minute (20 min).** Groups of four with the topic cards. Speak for ${L.jamSecs} seconds (${L.jamChallenge} for strong students); the others ask one question each. Rotate.`,
      '**Feedback (5 min).** Each group’s best speaker talks in front of the class.',
    ]],
  ]);
  day('Day 5: English culture and review', 'If Day 5 is a test day, move these to any free slot.', [
    ['First half: "Is it true in your country?" four corners (S)', a1 ? 'simple opinions: agree, disagree, same, different.' : 'giving and justifying opinions, agreeing and disagreeing politely, comparing cultures.', [
      '**Recall (5 min).** "What do you remember from Day 2’s Kahoot about English culture?"',
      `**Language (5 min).** ${L.cultureLang}`,
      '**Four corners (15 min).** Read a culture statement (answer key section). Students walk to a corner sign, then explain their choice.',
      `**Culture tips (15 min).** Groups write 5 tips for a new student: ${L.tipsExample} Choose the top 10 for a wall poster.`,
      '**Wrap-up (5 min).** "What’s the most surprising thing about England for you?"',
    ]],
    ['Second half: Brighton board game + Kahoot (S)', 'review the week’s grammar while speaking.', [
      '**Set up (5 min).** Groups of 3–4 with a board, a dice and counters.',
      `**Play (25 min).** Do the task on your square and speak for ${L.boardSecs}+ seconds; the group asks one question. Can’t do it? Go back 2. Monitor and note errors.`,
      '**Error clinic (5 min).** Board 5 errors; teams correct them.',
      '**Kahoot: Big week review (10 min).** Announce the week’s champion.',
    ]],
  ]);

  t.push(H1('Listening script: A Saturday in Brighton'));
  t.push(P(`Lucas leaves a voice message for his sister, Ana. ${LS.length[0].toUpperCase()}${LS.length.slice(1)}. Read it three times (gist, detail, Task 3).`));
  t.push(box(LS.script.map((p) => P(p, { run: { size: 23 } }))));

  t.push(H1('Answer keys'));
  t.push(H2('Day 1: Brighton true or false'));
  t.push(table(L.tf.map(([s, a], i) => [String(i + 1), s, a]), [500, 5400, 4006]));
  t.push(H2(`Day 2: ${W.genre} questions`));
  t.push(...numbered(W.questions.map(([q, a]) => `${q} **${a}**`)));
  t.push(P(`**Order of parts:** ${W.partsKey}.`));
  t.push(H2('Day 4: listening'));
  t.push(P(`**Task 1:** ${LS.placesKey}`));
  t.push(P(`**Task 2:** ${LS.tf.map(([, a], i) => `${i + 1} ${a}`).join(' · ')}`));
  t.push(P(`**Task 3:**`));
  t.push(...numbered(LS.task3.map(([q, a]) => LS.task3lines ? `${q} **${a}**` : `**${a}**`)));
  t.push(H2('Day 4: bingo definitions to read out'));
  t.push(table([['Word', 'Definition'], ...L.bingo], [2400, 7506], { header: true }));
  t.push(H2('Day 5: culture statements (read aloud)'));
  t.push(...numbered(L.culture));
  return t;
}

async function build() {
  const outDir = process.argv[2];
  for (const key of ['c2', 'c3']) {
    const L = LEVELS[key];
    const sub = `Brighton Skills Week · ${L.course} (${L.level})`;
    const sd = new Document({
      creator: 'Kings Brighton', title: `Brighton Week ${L.file}: Student Pack`, styles, numbering,
      sections: [
        { properties: portrait, headers: { default: header(`${sub} · Student pack`) }, footers: { default: footer() }, children: studentPack(L) },
        { properties: landscape, headers: { default: landscapeHeader(`${sub} · Board game`) }, footers: { default: footer() }, children: boardGame(L) },
      ],
    });
    const td = new Document({
      creator: 'Kings Brighton', title: `Brighton Week ${L.file}: Teacher Pack`, styles, numbering,
      sections: [{ properties: portrait, headers: { default: header(`${sub} · Teacher pack`) }, footers: { default: footer() }, children: teacherPack(L) }],
    });
    fs.writeFileSync(path.join(outDir, `Brighton Week ${L.file} - Student Pack.docx`), await Packer.toBuffer(sd));
    fs.writeFileSync(path.join(outDir, `Brighton Week ${L.file} - Teacher Pack.docx`), await Packer.toBuffer(td));
  }
  console.log('built');
}
build().catch((e) => { console.error(e); process.exit(1); });

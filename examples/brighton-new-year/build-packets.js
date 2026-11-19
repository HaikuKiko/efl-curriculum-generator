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

// ================= BRIGHTON & NEW YEAR: CONTENT (two levels) =================
const RED = 'C00000';
const KAHOOTS = [
  ['Brighton: basics', 'Brighton facts', '739bef19-5282-47db-9c2d-305e02e691e5'],
  ['Brighton: English culture', 'before Four corners', '7d6971c3-87f7-42f0-8808-e6dd0b2d4467'],
  ['Brighton: places', 'before Secret places', '75dc390b-2070-49ef-965b-2cafad5249b4'],
  ['New Year plans: going to, will, want to', 'future forms', '190ddd16-7748-4bfc-a4ab-1f8f8a6b5771'],
  ['Brighton & New Year: big review', 'end of the week', '2a9b54b3-568b-46df-8316-0f868e6fe796'],
];
const PLACES = [`The Royal Pavilion`, `Palace Pier`, `the i360`, `the Lanes`, `North Laine`, `the Clock Tower`,
  `Churchill Square`, `Brighton station`, `the Sea Life Centre`, `the West Pier`, `Brighton Marina`, `the beach`];
const DIRECTIONS = `go straight on · turn left / right · take the second left · it’s on your right · next to · opposite · between · on the corner of · at the end of`;
const CORNERS = ['AGREE', 'DISAGREE', 'NOT SURE', 'IT DEPENDS'];

const LEVELS = {
  c2: {
    course: 'Course 2', level: 'A1', dir: 'course-2-A1',
    future: `*want to*, *would like to* and simple *going to* for plans`,
    fswIntro: `Ask *"Do you want to...?"*, *"Are you going to...?"* or *"Can you...?"*. Write a name. Then ask one more question: *Where? When? Why?*`,
    fswDrill: `*Do you want to...? Are you going to...? Can you...?*`,
    fsw: [`wants to visit London`, `is going to travel next year`, `wants to learn a new hobby`, `likes fish and chips`, `lives in Hove`,
      `wants to swim in the sea`, `is going to see their family soon`, `wants to learn to cook`, `can name a street in Brighton`, `is going to buy a souvenir`],
    tf: [
      [`The Royal Pavilion looks like a palace from India.`, `T`, ``],
      [`Brighton is the capital of England.`, `F`, `London is the capital.`],
      [`The Lanes are big, new streets.`, `F`, `They are small and old.`],
      [`Brighton has two piers. One is very old and broken.`, `T`, `Palace Pier and the West Pier.`],
      [`You can swim in the sea in Brighton.`, `T`, ``],
      [`Brighton beach has a lot of sand.`, `F`, `It has pebbles (small stones).`],
      [`Brighton’s football team is called the Seagulls.`, `T`, ``],
      [`It never rains in Brighton.`, `F`, `It often rains!`],
    ],
    interviewTitle: 'My plans for the new year',
    interviewIntro: `Ask your partner. Write short notes. Then tell a new partner: *"Maria wants to... She is going to..."*`,
    interview: [`Where do you live in Brighton?`, `What do you want to do in Brighton this month?`, `What are you going to do this weekend?`,
      `What do you want to learn next year?`, `Where do you want to travel next year?`, `What is your New Year’s resolution?`,
      `What English food do you want to try?`, `What are you going to do on New Year’s Eve?`],
    reportModel: `*She wants to... She is going to...*`,
    questionStarters: [`Is it near...? Is it next to...?`, `Can you... there?`, `Is it big / old?`, `Is there a... there?`],
    plannerTitle: 'Plan a day out next weekend',
    tourists: [`Next weekend you want to see **old buildings**.`, `You want to go **shopping**. You have **£10**.`, `Next Saturday it’s **raining**. You have a **child**.`,
      `You want to take **photos** of Brighton.`, `You want to go **walking**. You like **nature**.`, `You want to eat **English food**.`],
    touristSay: `*I want to... Where can I go?*`,
    locals: [`Go to...`, `You can... there.`, `It’s good for...`, `It’s near...`, `You’re going to love it!`],
    planFrame: [`On Saturday morning I’m going to...`, `For lunch I want to...`, `In the afternoon I’m going to...`],
    bingo: [[`pier`, `A long road that goes out over the sea.`], [`pebbles`, `Small round stones on the beach.`],
      [`seagull`, `A white bird. It takes your food!`], [`beach`, `You sit here next to the sea.`],
      [`sea`, `Lots of salty water.`], [`tower`, `A very tall, thin building.`], [`shop`, `You buy things here.`],
      [`bus`, `A big red or blue car for lots of people.`], [`fireworks`, `Bright lights in the sky on New Year’s Eve.`],
      [`palace`, `A very big house for a king or queen.`], [`fish and chips`, `A typical English meal.`],
      [`sunny`, `The weather when you need sunglasses.`], [`windy`, `The weather when your hat flies away.`],
      [`party`, `People meet, eat, dance and have fun.`], [`resolution`, `A promise to do something better next year.`], [`queue`, `A line of people waiting.`]],
    jamSecs: 30, jamChallenge: 45,
    jam: [`My plans for this weekend`, `My New Year’s resolution`, `A place I want to visit`, `A hobby I want to try`,
      `My plans for next year`, `My dream job`, `My family`, `My favourite food`, `My home town`, `Brighton beach`,
      `My best friend`, `My favourite place in Brighton`],
    cultureLang: `*I agree. / I don’t agree. / I’m not sure. / It depends. / In my country...*`,
    culture: [`New Year’s resolutions are a good idea.`, `I want to live in England one day.`, `Next year will be a good year.`,
      `I’m going to speak English every day.`, `People in England drink a lot of tea.`, `English food is good.`,
      `New Year’s Eve is the best party of the year.`, `I want to visit London next year.`],
    extension4c: `Class resolutions poster: each student writes one sentence, *"This year I want to..."*, on a sticky note. Make a wall poster.`,
    board: [`START`, `Say 3 things you want to do in Brighton.`, `What’s your name? Spell it.`, `What are you going to do tonight?`,
      `Describe your bedroom: There is... There are...`, `What’s your New Year’s resolution?`, `Go forward 2!`,
      `What’s your favourite food? Why?`, `What do you want to learn next year?`, `Count from 20 to 30.`, `What did you do yesterday?`,
      `Ask the player on your left 2 questions.`, `Where do you want to go on holiday?`, `Go back 3!`, `What time do you get up?`,
      `Name 5 things on the beach.`, `What job do you want?`, `What are you going to eat tomorrow?`, `Miss a turn!`,
      `Talk about your family.`, `What do you want to buy in Brighton?`, `Name 3 places in Brighton.`, `Go forward 1!`,
      `What are you going to do this weekend?`, `Do you like fish and chips? Why?`, `Where do you want to live in the future?`,
      `Say 3 things you did this week.`, `What are you going to do on New Year’s Eve?`, `What’s your favourite English word?`,
      `FINISH! Say "Happy New Year, Brighton!"`],
    boardSecs: 10,
    dialogue: [
      ['A', `Happy New Year! What are your plans for this year?`],
      ['B', `I want to learn to swim. And you?`],
      ['A', `I’m going to visit London in March.`],
      ['B', `Great! What do you want to see?`],
      ['A', `I want to see Big Ben and the London Eye.`],
      ['B', `Are you going to go by train?`],
      ['A', `Yes. It’s only one hour from Brighton.`],
      ['B', `Can I come with you?`],
    ],
    dialogueNames: ['Ana', 'Ben'],
    dialogueFollow: `Change the plans to your own: *"I want to... I’m going to..."* and perform your new dialogue.`,
  },

  c3: {
    course: 'Course 3', level: 'A2+', dir: 'course-3-A2-plus',
    future: `*going to* (plans), *will* and *might* (predictions), present continuous (arrangements), *want to / would like to*, New Year’s resolutions`,
    fswIntro: `Ask *"Are you going to...?"*, *"Have you ever...?"* or *"Do you think you’ll...?"*. Write a name, then ask two follow-up questions: *When? Why? How?*`,
    fswDrill: `*Are you going to...? Have you ever...? Do you think you’ll...? Are you doing anything...?*`,
    fsw: [`has made a New Year’s resolution`, `is going to travel somewhere new next year`, `is meeting a friend this weekend`,
      `thinks they’ll live abroad one day`, `has broken a New Year’s resolution before`, `is going to take an exam soon`,
      `might change job or studies next year`, `has been on Palace Pier`, `would like to live in Brighton`, `has a plan for New Year’s Eve`],
    tf: [
      [`The Royal Pavilion was built for a prince who later became king.`, `T`, `King George IV.`],
      [`Brighton has always been a city.`, `F`, `Brighton and Hove became a city in 2000.`],
      [`The Lanes are wide, modern streets.`, `F`, `They’re narrow and old.`],
      [`Brighton has two piers, but one is in ruins.`, `T`, `Palace Pier and the West Pier.`],
      [`Palace Pier opened in the 1990s.`, `F`, `It opened in 1899.`],
      [`Hove is part of the same city as Brighton.`, `T`, `The city of Brighton and Hove.`],
      [`There’s a big arts festival in Brighton every May.`, `T`, `The Brighton Festival.`],
      [`Brighton’s football team is called the Pigeons.`, `F`, `They’re called the Seagulls.`],
    ],
    interviewTitle: 'My plans for the future',
    interviewIntro: `Interview your partner and make notes. Then report to a new group: *"Maria is going to... She thinks she’ll... She might..."*`,
    interview: [`What are you going to do when you finish this course?`, `What’s your New Year’s resolution? Do you think you’ll keep it?`,
      `What are you doing this weekend?`, `Where do you think you’ll be in five years?`, `What might you do differently next year?`,
      `What would you like to learn that you’ve never tried?`, `How do you think Brighton will change in the future?`,
      `What are you going to miss when you leave Brighton?`],
    reportModel: `*She’s going to... She thinks she’ll... She might...*`,
    questionStarters: [`Is it near...? Is it opposite...?`, `Can you... there? Have you been there?`, `Was it built before / after...?`, `Is it bigger / more famous than...?`],
    plannerTitle: 'Plan a day out next Saturday',
    tourists: [`You love **history**, but you hate crowds. You have one afternoon.`,
      `You have **£10** for the whole day and want something unusual to take home.`,
      `It’s going to **rain**, you’re with a **7-year-old**, and you need to be back by 4 pm.`,
      `You want **amazing photos**, but you’re scared of heights.`,
      `You love **nature and walking**. You have a whole day and a bus pass.`,
      `You want **typical English food**, but you’re vegetarian.`],
    touristSay: `*I’m planning a day out next Saturday. I’d like to... What would you recommend?*`,
    locals: [`If I were you, I’d...`, `You could... or... ; ... is better because...`, `It’s the best place for..., although...`,
      `If it rains, you might want to...`, `Make sure you... / Don’t forget to...`],
    planFrame: [`In the morning I’m going to...`, `I’m meeting... at... (arrangement)`, `If it rains, I might...`, `I think it’ll be... because...`],
    bingo: [[`promenade`, `A wide path next to the sea for walking.`], [`pebbles`, `Small round stones on the beach.`],
      [`seagull`, `A noisy white bird that steals your food.`], [`resolution`, `A promise to change something in the new year.`],
      [`pier`, `A long structure that goes out over the sea.`], [`arcade`, `A place with games machines.`],
      [`view`, `What you can see from a high place.`], [`fireworks`, `Explosions of light in the sky at midnight.`],
      [`countdown`, `Ten, nine, eight... before midnight.`], [`queue`, `A line of people waiting.`],
      [`goal`, `Something you want to achieve.`], [`tide`, `The sea coming in and going out.`],
      [`waves`, `The sea moves up and down in these.`], [`breezy`, `A little windy.`],
      [`crowded`, `Full of people.`], [`souvenir`, `Something you buy to remember a place.`]],
    jamSecs: 60, jamChallenge: 90,
    jam: [`My New Year’s resolutions`, `Where I’ll be in five years`, `How Brighton will change in 20 years`, `Technology in the future`,
      `My dream job`, `Plans I’ve made for this year`, `Something I’ll never do again`, `The perfect New Year’s Eve`,
      `Will people still learn languages in the future?`, `A skill I’m going to learn`, `My home town vs Brighton`, `Tourists: good or bad for a city?`],
    cultureLang: `*I completely agree because... / I’m not sure about that... / It depends on... / In my country, people usually...*`,
    culture: [`New Year’s resolutions are a waste of time.`, `In 20 years, we won’t need to learn languages.`,
      `Next year will be better than this year.`, `People will live on the Moon in our lifetime.`,
      `It’s better to plan everything than to be spontaneous.`, `Queuing is a sign of a fair society.`,
      `Small talk about the weather is a waste of time.`, `Tourists should learn some of the local language.`],
    extension4c: `Predictions poster: groups write 5 predictions for Brighton in 2050 (*"There will be... People won’t... We might..."*) and vote on the most likely.`,
    board: [`START`, `What’s your New Year’s resolution? How will you keep it?`, `If you won £1,000 tomorrow, what would you do in Brighton?`,
      `Give 3 differences between Brighton and your home town.`, `What are you doing this weekend?`,
      `Where will you be this time next year?`, `Go forward 2!`, `What should a new student do in their first week? Why?`,
      `Predict 3 things about technology in 2050.`, `What are you going to do when you finish your course?`,
      `Tell a story about a time something went wrong.`, `Ask the player on your left 3 questions about their plans.`,
      `How long have you been learning English?`, `Go back 3!`, `What might change in Brighton in 50 years?`,
      `Explain a New Year tradition from your country.`, `Would you rather live by the sea or in the mountains? Why?`,
      `What’s something you’ll never do again?`, `What will you miss about Brighton?`, `Miss a turn!`,
      `If you could live anywhere next year, where would it be?`, `Give directions from school to the beach.`,
      `What’s a resolution you’ve broken? Why?`, `Go forward 1!`, `Persuade the group to visit your country.`,
      `What will you do on New Year’s Eve?`, `Do you think you’ll speak perfect English one day? Why?`,
      `What’s the most useful thing you’ve learned this week?`, `What are you going to do differently this year?`,
      `FINISH! Say "Happy New Year, Brighton!"`],
    boardSecs: 30,
    dialogue: [
      ['A', `So, have you made any New Year’s resolutions?`],
      ['B', `A few. I’m going to join a gym, and I’m going to stop eating so much chocolate.`],
      ['A', `Do you think you’ll keep them?`],
      ['B', `Probably not! What about you?`],
      ['A', `I’m taking my English exam in June, so I’m going to study every evening.`],
      ['B', `That’s a good idea. Are you doing anything this weekend?`],
      ['A', `I’m meeting some friends on Saturday. We might go up the i360.`],
      ['B', `Oh, I’ve never been up it. Is it expensive?`],
      ['A', `It’s not cheap, but the view is amazing. Why don’t you come?`],
      ['B', `I’d love to! What time are you meeting?`],
      ['A', `At eleven, by the Clock Tower.`],
      ['B', `Perfect. I’ll see you there!`],
    ],
    dialogueNames: ['Sam', 'Mia'],
    dialogueFollow: `Write your own version with your real resolutions and weekend plans, then perform it for another pair.`,
  },
};
// ================= ACTIVITY PACKETS =================
const redRun = (text) => new TextRun({ text, color: RED, bold: true });
const redCell = (text, width) => cell([new Paragraph({ children: [redRun(text)] })], width);

function teacherNotes(L, a) {
  const t = [];
  t.push(new Paragraph({ heading: 'Heading1', children: [new TextRun(`Teacher notes: ${a.title(L)}`)] }));
  t.push(Small('Print the whole packet. Keep this page; give students the pages after it.'));
  t.push(table([
    ['Level', `${L.course} (${L.level})`],
    ['Time', a.time(L)],
    ['Language', a.focus(L)],
    ['Prepare', a.prep(L)],
  ], [1800, 8106]));
  t.push(H3('Steps'));
  t.push(...numbered(a.steps(L)));
  t.push(H3('Make it shorter or longer'));
  t.push(...bullets(a.tips(L)));
  if (a.extra) t.push(...a.extra(L));
  return t;
}

// Vanishing dialogue helpers
function maskWord(tok, mode) {
  const m = tok.match(/^([^A-Za-z0-9’'£]*)([A-Za-z0-9’'£]+)(.*)$/);
  if (!m) return tok;
  const [, pre, core, post] = m;
  if (mode === 'blank') return pre + '_'.repeat(Math.max(3, core.length)) + post;
  if (mode === 'first') return pre + core[0] + '_'.repeat(Math.max(2, core.length - 1)) + post;
  return tok;
}
function vanish(lineText, lineIdx, stage) {
  const toks = lineText.split(' ');
  return toks.map((tok, i) => {
    const rank = ((lineIdx * 31 + i * 17 + 7) * 37) % 100;
    if (stage === 1) return tok;
    if (stage === 2) return rank < 30 ? maskWord(tok, 'blank') : tok;
    if (stage === 3) return rank < 65 ? maskWord(tok, 'blank') : tok;
    if (stage === 4) return maskWord(tok, 'first');
    return '';
  }).join(' ');
}
function dialoguePage(L, stage) {
  const s = [];
  const labels = { 1: 'Read it together', 2: 'Some words are gone', 3: 'Most words are gone', 4: 'Only the first letters', 5: 'From memory!' };
  s.push(H1(`Vanishing dialogue ${stage}/5: ${labels[stage]}`));
  s.push(Small('Practise with your partner. Then swap roles. Turn the page when you’re ready.'));
  const [na, nb] = L.dialogueNames;
  s.push(table(L.dialogue.map(([who, text], i) => [
    who === 'A' ? na : nb,
    stage === 5 ? '____________________________________________________' : vanish(text, i, stage),
  ]), [1500, 8406], { rowHeight: stage === 5 ? 700 : 560 }));
  return s;
}

const ACTIVITIES = [
  {
    num: '01', file: 'Find someone who', title: () => 'Find someone who',
    time: () => '20–25 minutes', focus: (L) => L.future,
    prep: () => 'One worksheet per student.',
    steps: (L) => [
      `Drill the question forms: ${L.fswDrill}`,
      'Students turn each prompt into a question. Check two or three together.',
      'Students stand up and mingle. When someone says yes, write the name and ask a follow-up question.',
      'Stop after about 12 minutes. Feedback: "Who is going to travel next year? Where?" Board useful language.',
    ],
    tips: () => ['Shorter: students only need 5 names.', 'Longer: students report back in pairs using *he / she*.', 'Warmer: run the *New Year plans* Kahoot first.'],
    student: (L) => [
      H1('Find someone who...', true), NameLine(), P(L.fswIntro),
      table([['Find someone who...', 'Name', 'Extra information'], ...L.fsw.map((f) => [f, '', ''])], [4200, 2200, 3506], { header: true, rowHeight: 700 }),
    ],
  },
  {
    num: '02', file: 'Brighton true or false', title: () => 'Brighton: true or false?', answers: true,
    time: () => '15 minutes', focus: () => 'Brighton facts; correcting false sentences.',
    prep: () => 'One worksheet per team of three. The answer version is a separate file (answers in red).',
    steps: () => [
      'Teams of three. Give each team a worksheet.',
      'Read each sentence aloud twice. Teams write T or F.',
      'Go through the answers. One point for each right answer, a bonus point for correcting a false sentence.',
      'Ask: "What surprised you?"',
    ],
    tips: () => ['Shorter: do the first 4 sentences only.', 'Longer: teams write 2 more true/false sentences about their home town for another team.', 'Pair with the *Brighton: basics* Kahoot.'],
    student: (L, ans) => [
      H1(`Brighton: true or false?${ans ? ' (answers)' : ''}`, true), P('Team name: ______________________'),
      P('Listen to your teacher. Write **T** (true) or **F** (false). If it’s false, correct it for a bonus point!'),
      new Table({
        width: { size: PORTRAIT_W, type: WidthType.DXA }, columnWidths: [500, 4900, 900, 3606],
        rows: [
          new TableRow({ tableHeader: true, children: ['#', 'Sentence', 'T / F', 'Correction'].map((h, i) => cell(h, [500, 4900, 900, 3606][i], { fill: PURPLE, run: { bold: true, color: 'FFFFFF' } })) }),
          ...L.tf.map(([q, a, c], i) => new TableRow({ height: { value: 650, rule: HeightRule.ATLEAST }, children: [
            cell(String(i + 1), 500), cell(q, 4900), ans ? redCell(a, 900) : cell('', 900), ans ? redCell(c, 3606) : cell('', 3606)] })),
        ],
      }),
    ],
  },
  {
    num: '03', file: 'Interview my plans', title: (L) => `Interview: ${L.interviewTitle.toLowerCase()}`,
    time: () => '35–40 minutes', focus: (L) => `${L.future}; reporting in the third person.`,
    prep: () => 'One worksheet per student.',
    steps: (L) => [
      'Lead-in: board *Next year: one word*. Each student says one word about their plans.',
      'Pairs: A asks all 8 questions and makes notes; then swap.',
      `New groups of four: each student tells the group about their partner. Model on the board: ${L.reportModel}`,
      'Listeners ask one question each.',
      'Error clinic: board 5 errors you heard (no names); teams correct them.',
    ],
    tips: () => ['Shorter: choose 4 questions.', 'Longer: class survey; tally the most popular plan or resolution.', 'Follow with the *New Year plans* Kahoot.'],
    student: (L) => [
      H1(`Interview: ${L.interviewTitle.toLowerCase()}`, true), NameLine(), P(L.interviewIntro), P("My partner's name: ______________________"),
      table([['Question', 'Notes'], ...L.interview.map((q) => [q, ''])], [5000, 4906], { header: true, rowHeight: 900 }),
    ],
  },
  {
    num: '04', file: 'Vanishing dialogue', title: () => 'Vanishing dialogue: New Year plans',
    time: () => '25–30 minutes', focus: (L) => `${L.future}; pronunciation and fluency.`,
    prep: () => 'One copy of the 5 dialogue pages per pair (or project them one at a time).',
    steps: (L) => [
      'Read page 1 aloud. Students repeat each line (choral drill), then focus on stress and linking.',
      'Pairs read page 1: A and B, then swap roles.',
      'Pairs turn to page 2 and say the dialogue again, filling the gaps from memory. Swap roles.',
      'Continue to pages 3, 4 and 5. On page 5 there are no words, only the speakers.',
      'Two or three pairs perform it to the class from memory.',
      `Follow-up: ${L.dialogueFollow}`,
    ],
    tips: () => ['Shorter: stop at page 4.', 'Projecting: show one page at a time and let the class say it together before pairs practise.', 'Fun: pairs perform it in a different mood (excited, tired, whispering).'],
    extra: () => [P('**The full dialogue** is on the first student page (1/5).')],
    student: (L) => [1, 2, 3, 4, 5].flatMap((n) => dialoguePage(L, n)),
  },
  {
    num: '05', file: 'Secret places and directions', title: () => 'Secret places and big-screen directions',
    time: () => '30 minutes', focus: () => 'Brighton places; prepositions of place; directions; yes/no questions.',
    prep: () => 'One worksheet per pair (or show it on screen). Google Maps open on central Brighton on the classroom screen.',
    steps: () => [
      'Warmer: the *Brighton: places* Kahoot (10 minutes), optional.',
      'Board the directions phrases from the worksheet.',
      'Guess the place: in pairs, A secretly chooses a place; B asks up to 10 yes/no questions. Swap.',
      'Big-screen directions: mark the school on Google Maps. A team chooses a secret place and gives directions step by step. You follow with the mouse; other teams guess.',
      'Ask: "Which of these places are you going to visit next?"',
    ],
    tips: () => ['Shorter: skip the big-screen directions.', 'Longer: students give directions from school to their home.'],
    student: (L) => [
      H1('Secret places', true), P('Choose a secret place. Your partner asks up to 10 yes/no questions to guess it.'),
      cards(PLACES, 3, 1000, 28), spacer(), box([P('**Question starters**'), ...bullets(L.questionStarters)]), spacer(), box([P('**Directions**'), P(DIRECTIONS)]),
    ],
  },
  {
    num: '06', file: 'Plan a day out role-play', title: (L) => `Role-play: ${L.plannerTitle.toLowerCase()}`,
    time: () => '40–45 minutes', focus: (L) => `Recommendations; ${L.future}.`,
    prep: () => 'Cut up one set of tourist cards per 6 students. One planner sheet per student.',
    steps: (L) => [
      `Model: you are a tourist; a strong student is a local. Tourist: ${L.touristSay}`,
      'Board the locals’ useful language.',
      'Half the class are tourists (cards), half are locals. Tourists talk to three locals and note each idea. Swap halfway.',
      'Tourists complete the planner and tell a partner about their day out.',
      'Class awards: best rainy-day plan, cheapest day out, best food.',
    ],
    tips: () => ['Shorter: tourists talk to two locals only.', 'Longer: pairs present their day out to the class like a TV advert.'],
    student: (L) => [
      H1('Tourist role cards', true), Small('Cut along the dashed lines. One card per tourist.'),
      cards(L.tourists.map((t) => `**TOURIST** — ${t}`), 2, 1600, 24), spacer(),
      box([P('**Locals: useful language**'), ...bullets(L.locals)]),
      H1(`My plan: ${L.plannerTitle.toLowerCase()}`, true), NameLine(),
      table([['Local’s name', 'Their idea'], ['', ''], ['', ''], ['', '']], [3000, 6906], { header: true, rowHeight: 800 }),
      H3('My plan'), ...L.planFrame.flatMap((f) => [P(f, { spacing: { after: 0 } }), WriteLine(), spacer()]),
    ],
  },
  {
    num: '07', file: 'Brighton bingo', title: () => 'Brighton and New Year bingo',
    time: () => '15 minutes', focus: () => 'Seaside and New Year vocabulary.',
    prep: () => 'One sheet per 2 students (cut in half).',
    steps: () => [
      'Students choose 9 words from the box and write them in their grid.',
      'Read the definitions below in any order. Students cross out the word.',
      'First full line shouts *"Brighton!"* and reads the words back to check.',
      'Play again with the second grid.',
    ],
    tips: () => ['Harder: students read the definitions to each other in small groups.', 'Extension: students write a sentence about their plans with 3 of the words.'],
    extra: (L) => [H3('Definitions to read out'), ...L.bingo.map(([w, d]) => P(`**${w}**: ${d}`, { spacing: { after: 30 } }))],
    student: (L) => {
      const grid = () => new Table({
        width: { size: 6300, type: WidthType.DXA }, columnWidths: [2100, 2100, 2100], alignment: AlignmentType.CENTER,
        rows: [0, 1, 2].map(() => new TableRow({ height: { value: 1300, rule: HeightRule.EXACT }, children: [0, 1, 2].map(() => cell('', 2100, { border: { style: BorderStyle.SINGLE, size: 12, color: PURPLE } })) })),
      });
      return [
        H1('Bingo', true), P('Choose 9 words from the box. Write them in your grid. Listen to your teacher and cross out the words. A full line? Shout **"Brighton!"**'),
        box([P(L.bingo.map((b) => b[0]).join('  ·  '), { align: AlignmentType.CENTER, run: { size: 24 } })]),
        spacer(), grid(),
        new Paragraph({ spacing: { before: 300, after: 300 }, border: { bottom: dashed }, children: [new TextRun({ text: '✂  cut here', size: 16, color: '888888' })] }),
        P('Choose 9 words. Write them in your grid.'), grid(),
      ];
    },
  },
  {
    num: '08', file: 'Just a minute', title: () => 'Just a minute',
    time: () => '20–25 minutes', focus: (L) => `Fluency; ${L.future}.`,
    prep: () => 'Cut up one set of topic cards per group of four.',
    steps: (L) => [
      `Model: talk about a card for ${L.jamSecs} seconds without stopping.`,
      `Groups of four. One student takes a card and speaks for ${L.jamSecs} seconds (${L.jamChallenge} for strong students).`,
      'The others listen for long stops, then each asks one question.',
      'Rotate until everyone has spoken 2 or 3 times.',
      'Each group’s best speaker talks in front of the class.',
    ],
    tips: () => ['Easier: give 1 minute of thinking time first.', 'Harder: listeners can "buzz" for repetition and take over the topic.'],
    student: (L) => [
      H1('Just a minute: topic cards', true),
      Small(`Cut up. Speak for ${L.jamSecs} seconds without stopping (${L.jamChallenge} seconds for a challenge!).`),
      cards(L.jam, 3, 1600, 24),
    ],
  },
  {
    num: '09', file: 'Four corners', title: () => 'Four corners: New Year and English culture',
    time: () => '25–40 minutes', focus: () => 'Opinions; agreeing and disagreeing; comparing cultures; the future.',
    prep: () => 'Print the 4 signs and stick one in each corner of the room. The statements are below (read them aloud).',
    steps: (L) => [
      'Optional warmer: the *Brighton: English culture* Kahoot.',
      `Board the language: ${L.cultureLang}`,
      'Read a statement. Students walk to the corner that matches their opinion.',
      'Students explain their choice to the others in their corner, then one person from each corner tells the class.',
      'Repeat with 5–8 statements.',
    ],
    tips: (L) => ['Shorter: use 4 statements.', `Extension: ${L.extension4c}`],
    extra: (L) => [H3('Statements to read aloud'), ...numbered(L.culture)],
    student: () => CORNERS.map((c, i) => new Paragraph({ pageBreakBefore: true, spacing: { before: 3600 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: c, bold: true, size: 150, color: PURPLE })] })),
  },
  {
    num: '10', file: 'Board game', title: () => 'Brighton and New Year board game', landscape: true,
    time: () => '30–35 minutes', focus: (L) => `Review of the week’s language; ${L.future}.`,
    prep: () => 'One board per group of 3–4 (A3 if possible), a dice and counters.',
    steps: (L) => [
      'Groups of 3–4 with a board, a dice and counters.',
      `Roll and move. Do the task on the square and speak for at least ${L.boardSecs} seconds. The group asks one question.`,
      'Can’t do it? Go back 2. Monitor and note errors.',
      'Error clinic: board 5 errors; teams correct them.',
      'Finish with the *Brighton & New Year: big review* Kahoot.',
    ],
    tips: () => ['Shorter: the first player to square 20 wins.', 'No dice? Students can use a dice app on a phone.'],
    student: (L) => {
      const COLS = 6, ROWS = 5;
      const w = Math.floor(LANDSCAPE_W / COLS);
      const rows = [];
      for (let r = 0; r < ROWS; r++) {
        const idx = Array.from({ length: COLS }, (_, c) => r * COLS + c);
        if (r % 2 === 1) idx.reverse();
        rows.push(new TableRow({ height: { value: 1700, rule: HeightRule.EXACT }, cantSplit: true, children: idx.map((i) => {
          const special = i === 0 || i === 29; const fun = /^(Go|Miss)/.test(L.board[i]);
          return cell([
            new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: String(i + 1), bold: true, size: 28, color: special ? 'FFFFFF' : PURPLE })] }),
            new Paragraph({ children: [new TextRun({ text: L.board[i], size: special ? 24 : 17, bold: special || fun, color: special ? 'FFFFFF' : '222222' })] }),
          ], w, { fill: special ? PURPLE : fun ? 'FBE9F1' : (r + idx.indexOf(i)) % 2 ? 'FFFFFF' : LIGHT, border: { style: BorderStyle.SINGLE, size: 8, color: PURPLE } });
        }) }));
      }
      return [
        new Paragraph({ heading: 'Heading1', children: [new TextRun('Brighton and New Year board game')] }),
        P(`Roll the dice and move. Do the task and speak for at least ${L.boardSecs} seconds. The group asks you one question. Can’t do it? Go back 2. Follow the numbers: the path snakes back and forth.`, { run: { size: 18 } }),
        new Table({ width: { size: w * COLS, type: WidthType.DXA }, columnWidths: Array(COLS).fill(w), rows }),
      ];
    },
  },
];

function overview(L) {
  const o = [];
  o.push(new Paragraph({ heading: 'Heading1', children: [new TextRun(`Brighton and New Year: ${L.course} (${L.level})`)] }));
  o.push(P(`Ten speaking-based activities about Brighton and the new year. Use them in any order and fit them into your week where you like.`));
  o.push(P(`**New Year language:** ${L.future}.`));
  o.push(H2('How the packets work'));
  o.push(...bullets([
    'Each activity is its own file. Print the whole file.',
    'The **first page is for you** (teacher notes: time, prep, steps, ideas). The pages after it are for students.',
    'Activities with answers have a second file ending in **ANSWERS**: the same pages with the answers filled in **in red**.',
  ]));
  o.push(H2('The activities'));
  o.push(table([['#', 'Activity', 'Time', 'Answers?'], ...ACTIVITIES.map((a) => [a.num, a.title(L), a.time(L), a.answers ? 'Yes (red)' : '—'])], [700, 5300, 2200, 1706], { header: true }));
  o.push(H2('Kahoots (shared by Course 2 and Course 3)'));
  o.push(P('Each link opens the kahoot’s page: press **Start** to host, or **Duplicate** to copy it into your account. They work once the kahoots are set to public.'));
  o.push(new Table({
    width: { size: PORTRAIT_W, type: WidthType.DXA }, columnWidths: [4300, 2500, 3106],
    rows: [
      new TableRow({ tableHeader: true, children: ['Kahoot', 'Good for', 'Link'].map((h, i) => cell(h, [4300, 2500, 3106][i], { fill: PURPLE, run: { bold: true, color: 'FFFFFF' } })) }),
      ...KAHOOTS.map(([n, g, u]) => new TableRow({ children: [cell(n, 4300), cell(g, 2500), cell([new Paragraph({ children: [new ExternalHyperlink({ link: `https://create.kahoot.it/details/${u}`, children: [new TextRun({ text: 'Open the kahoot', style: 'Hyperlink' })] })] })], 3106)] })),
    ],
  }));
  o.push(H2('Low prep'));
  o.push(...bullets(['No maps to print: Secret places uses Google Maps on the classroom screen.', 'Things to cut: tourist cards, Just a minute cards, bingo sheets.', 'Bring dice and counters for the board game.']));
  return o;
}

async function build() {
  const outDir = process.argv[2];
  for (const key of ['c2', 'c3']) {
    const L = LEVELS[key];
    const dir = path.join(outDir, L.dir);
    fs.mkdirSync(dir, { recursive: true });
    const sub = `Brighton and New Year · ${L.course} (${L.level})`;
    const pSec = (children) => ({ properties: portrait, headers: { default: header(sub) }, footers: { default: footer() }, children });
    const lSec = (children) => ({ properties: landscape, headers: { default: landscapeHeader(sub) }, footers: { default: footer() }, children });
    const save = async (name, sections) => {
      const doc = new Document({ creator: 'Kings Brighton', title: name, styles, numbering, sections });
      fs.writeFileSync(path.join(dir, `${name}.docx`), await Packer.toBuffer(doc));
    };
    await save(`00 Overview`, [pSec(overview(L))]);
    for (const a of ACTIVITIES) {
      const mk = (ans) => a.landscape ? [pSec(teacherNotes(L, a)), lSec(a.student(L, ans))] : [pSec([...teacherNotes(L, a), ...a.student(L, ans)])];
      await save(`${a.num} ${a.file}`, mk(false));
      if (a.answers) await save(`${a.num} ${a.file} - ANSWERS`, mk(true));
    }
  }
  console.log('built');
}
build().catch((e) => { console.error(e); process.exit(1); });

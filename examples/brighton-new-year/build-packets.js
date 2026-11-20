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

// ================= NEW YEAR IN BRIGHTON: CONTENT (two levels) =================
const RED = 'C00000';
const KAHOOTS = [
  ['New Year’s resolutions vocabulary', 'Resolutions bingo, Rank the resolutions', 'd9d9f787-10e2-4115-acb2-c5182605af94'],
  ['New Year plans: going to, will, want to', 'any activity: future forms', '190ddd16-7748-4bfc-a4ab-1f8f8a6b5771'],
  ['Brighton: places', 'My Brighton bucket list', '75dc390b-2070-49ef-965b-2cafad5249b4'],
  ['Brighton: basics', 'a warmer about the city', '739bef19-5282-47db-9c2d-305e02e691e5'],
  ['Brighton: English culture', 'Four corners', '7d6971c3-87f7-42f0-8808-e6dd0b2d4467'],
  ['Brighton & New Year: big review', 'the end of the week', '2a9b54b3-568b-46df-8316-0f868e6fe796'],
];
const PLACES = [`The Royal Pavilion`, `Palace Pier`, `the i360`, `the Lanes`, `North Laine`, `the Clock Tower`,
  `Churchill Square`, `Brighton station`, `the Sea Life Centre`, `the West Pier`, `Brighton Marina`, `the beach`];
const BUCKET_IDEAS = `the i360 · Palace Pier · the Royal Pavilion · the Lanes · North Laine · the Booth Museum · the Sea Life Centre · Brighton Marina · the Seven Sisters · a football match at the Amex · swimming in the sea · fish and chips on the beach`;
const CORNERS = ['AGREE', 'DISAGREE', 'NOT SURE', 'IT DEPENDS'];

const LEVELS = {
  c2: {
    course: 'Course 2', level: 'A1', dir: 'course-2-A1',
    future: `*want to*, *would like to* and *going to* for plans and resolutions`,
    // 01
    fswIntro: `Ask *"Do you want to...?"* or *"Are you going to...?"*. Write a name. Then ask one more question: *When? Where? Why?*`,
    fswDrill: `*Do you want to...? Are you going to...?*`,
    fsw: [`wants to get fit`, `is going to save money`, `wants to swim in the sea this year`, `wants to learn to cook`,
      `is going to visit London`, `wants to make new friends`, `is going to go to bed early`, `wants to try a new sport`,
      `is going to eat more fruit`, `wants to speak English every day`],
    // 02
    ttlModel: [`I want to learn to cook.`, `I’m going to run every day.`, `I want to visit Scotland.`],
    ttlFrames: [`I want to ______________________________`, `I’m going to ______________________________`, `I want to ______________________________`],
    ttlQuestions: [`When?`, `Where?`, `Why?`, `Who with?`, `Do you really want to...?`, `Are you really going to...?`],
    // 03
    interviewTitle: 'My year in Brighton',
    interviewIntro: `Ask your partner. Write short notes. Then tell a new partner: *"Maria wants to... She is going to..."*`,
    interview: [`Where do you live in Brighton?`, `What do you want to do in Brighton this year?`, `What are you going to do this weekend?`,
      `What is your New Year’s resolution?`, `What do you want to learn this year?`, `Where do you want to go this year?`,
      `What English food do you want to try?`, `What are you going to do after the course?`],
    reportModel: `*She wants to... She is going to...*`,
    // 04
    dialogue: [
      ['A', `Happy New Year! What’s your resolution?`],
      ['B', `I want to learn to swim.`],
      ['A', `In the sea? It’s very cold!`],
      ['B', `No, at the swimming pool. And you?`],
      ['A', `I’m going to speak English every day.`],
      ['B', `Good idea! Are you going to watch English films?`],
      ['A', `Yes, and I want to make English friends.`],
      ['B', `Let’s start now. Hello, I’m Ben!`],
    ],
    dialogueNames: ['Ana', 'Ben'],
    dialogueSetting: 'New Year’s Day, on Brighton beach.',
    dialogueFollow: `Change the resolutions to your own (*"I want to... I’m going to..."*) and perform your new dialogue.`,
    // 05
    questionStarters: [`Is it near...? Is it next to...?`, `Can you... there?`, `Is it big / old?`, `Is there a... there?`],
    bucketFrame: `I want to go to... because...`,
    bucketCols: ['Place', 'Why?', 'When?'],
    // 06
    plannerTitle: 'Plan your first weekend of the year',
    tourists: [`It’s your first weekend in Brighton. You want to see **old buildings**.`, `You want to go **shopping**. You have **£10**.`,
      `On Saturday it’s **raining**. You have a **child**.`, `You want to take **photos** of Brighton.`,
      `You want to go **walking**. You like **nature**.`, `You want to eat **English food**.`],
    touristSay: `*I want to... Where can I go?*`,
    locals: [`Go to...`, `You can... there.`, `It’s good for...`, `It’s near...`, `You’re going to love it!`],
    planFrame: [`On Saturday morning I’m going to...`, `For lunch I want to...`, `In the afternoon I’m going to...`, `On Sunday I want to...`],
    // 07
    bingo: [[`get fit`, `Do sport and be healthy.`], [`save money`, `Keep your money. Don’t spend it.`],
      [`eat healthily`, `Eat fruit and vegetables.`], [`go to bed early`, `Sleep before 10 o’clock.`],
      [`make new friends`, `Meet new people.`], [`learn`, `Study something new.`], [`join a club`, `Be in a group, for example a football team.`],
      [`drink water`, `Not coffee or cola!`], [`read more`, `Read more books.`], [`walk more`, `Go on foot, not by bus.`],
      [`travel`, `Go to new places.`], [`cook`, `Make your own food.`], [`stop`, `Finish doing something.`],
      [`start`, `Begin something new.`], [`try`, `Do something for the first time.`], [`relax`, `Rest and feel calm.`]],
    // 08
    rankLang: `*... is easy. ... is hard. ... is easier than ...*`,
    rank: [`drink more water`, `eat more fruit`, `go to bed early`, `walk to school`, `stop eating chocolate`,
      `save money`, `learn to swim`, `speak English every day`, `get up early`, `call my family every week`],
    rankTypical: [2, 4, 6, 3, 10, 9, 8, 5, 7, 1],
    // 09
    jamSecs: 30, jamChallenge: 45,
    jam: [`My New Year’s resolution`, `A place I want to visit this year`, `My plans for this weekend`, `My dream job`,
      `Something I want to learn`, `What I’m going to do after class`, `My next birthday`, `A sport I want to try`,
      `Food I want to try in England`, `Where I want to live in the future`, `My next holiday`, `What I’m going to do tomorrow`],
    // 10
    cultureLang: `*I agree. / I don’t agree. / I’m not sure. / It depends. / In my country...*`,
    culture: [`Resolutions are a good idea.`, `January is a good month to start something new.`, `I’m going to speak English every day.`,
      `I want to live in Brighton one day.`, `This year will be a good year.`, `It’s easy to get fit.`,
      `Saving money is hard.`, `I want to learn a new sport this year.`],
    extension4c: `Class resolutions poster: each student writes *"This year I want to..."* on a sticky note. Make a wall poster.`,
    // 11
    board: [`START`, `Say 3 things you want to do this year.`, `What are you going to do tomorrow?`, `What’s your New Year’s resolution?`,
      `What do you want to eat tonight?`, `Where do you want to go this weekend?`, `Go forward 2!`, `What sport do you want to try?`,
      `What are you going to do after class?`, `Name 3 places you want to visit in Brighton.`, `What job do you want?`,
      `Ask the player on your left: "What’s your resolution?"`, `What do you want to learn this year?`, `Go back 3!`,
      `Where do you want to go on holiday?`, `What are you going to buy this week?`, `What do you want to cook?`,
      `Who do you want to call this week?`, `Miss a turn!`, `What do you want to do on your birthday?`,
      `What are you going to do this weekend?`, `What do you want to see in London?`, `Go forward 1!`,
      `Where do you want to live in the future?`, `Do you want to swim in the sea? Why?`, `What English word do you want to learn?`,
      `Say 3 things you’re going to do next week.`, `What do you want to be good at?`, `Is your resolution easy or hard?`,
      `FINISH! Say "Happy New Year!"`],
    boardSecs: 10,
  },

  c3: {
    course: 'Course 3', level: 'A2+', dir: 'course-3-A2-plus',
    future: `*going to* (plans), *will* and *might* (predictions), present continuous (arrangements), *would like to*, resolution phrases (*give up, cut down on, take up*)`,
    fswIntro: `Ask *"Are you going to...?"*, *"Have you ever...?"* or *"Do you think you’ll...?"*. Write a name, then ask two follow-up questions: *Why? How? When?*`,
    fswDrill: `*Are you going to...? Have you ever...? Do you think you’ll...? Are you doing anything...?*`,
    fsw: [`has already made a resolution`, `is going to join a club in Brighton`, `has broken a resolution before`,
      `is going to give up something this year`, `thinks they’ll keep their resolution`, `is planning a trip this year`,
      `is meeting a friend this weekend`, `would like to learn to surf`, `might move to a new city next year`, `is going to read a book in English this year`],
    ttlModel: [`I’m going to give up sugar.`, `I’m joining a gym on Monday.`, `I’ll learn to surf in Brighton.`],
    ttlFrames: [`I’m going to ______________________________`, `I’m ______________ (arrangement) ______________________________`, `I’ll / I might ______________________________`],
    ttlQuestions: [`Why did you decide to...?`, `How are you going to...?`, `Do you think you’ll...?`, `When are you starting?`, `Who’s going to help you?`, `What will you do if...?`],
    interviewTitle: 'My year in Brighton',
    interviewIntro: `Interview your partner and make notes. Then report to a new group: *"Maria is going to... She thinks she’ll... She might..."*`,
    interview: [`What are you going to do before you leave Brighton?`, `What’s your resolution? Do you think you’ll keep it?`,
      `What are you doing this weekend?`, `Where will you be this time next year?`, `How are you going to improve your English this year?`,
      `What new thing would you like to try in Brighton?`, `What might be difficult for you this year?`, `How do you think you’ll feel at the end of the course?`],
    reportModel: `*She’s going to... She thinks she’ll... She might...*`,
    dialogue: [
      ['A', `Happy New Year! Did you see the fireworks last night?`],
      ['B', `Yes, they were amazing! So, have you made any resolutions?`],
      ['A', `A few. I’m going to join a running club, and I’m going to cut down on chocolate.`],
      ['B', `Do you think you’ll keep them?`],
      ['A', `Probably not the chocolate one! What about you?`],
      ['B', `I’m taking my English exam in June, so I’m going to study every evening.`],
      ['A', `That’s a good idea. Are you doing anything this weekend?`],
      ['B', `I’m meeting some friends on Saturday. We might go up the i360.`],
      ['A', `I’ve never been up it. Is it expensive?`],
      ['B', `It’s not cheap, but the view is amazing. Why don’t you come?`],
      ['A', `I’d love to! What time are you meeting?`],
      ['B', `At eleven, by the Clock Tower. See you there!`],
    ],
    dialogueNames: ['Sam', 'Mia'],
    dialogueSetting: 'New Year’s Day, walking along Brighton seafront.',
    dialogueFollow: `Write your own version with your real resolutions and weekend plans, then perform it for another pair.`,
    questionStarters: [`Is it near...? Is it opposite...?`, `Can you... there? Have you been there?`, `Was it built before / after...?`, `Is it bigger / more famous than...?`],
    bucketFrame: `I’m definitely going to... / I might... if... / I’d love to... because...`,
    bucketCols: ['Place', 'Why?', 'When / with who?'],
    plannerTitle: 'Plan your first weekend of the year',
    tourists: [`You love **history**, but you hate crowds. You have one afternoon.`,
      `You have **£10** for the whole day and want something unusual to take home.`,
      `It’s going to **rain**, you’re with a **7-year-old**, and you need to be back by 4 pm.`,
      `You want **amazing photos**, but you’re scared of heights.`,
      `You love **nature and walking**. You have a whole day and a bus pass.`,
      `You want **typical English food**, but you’re vegetarian.`],
    touristSay: `*It’s my first weekend of the year here. I’d like to... What would you recommend?*`,
    locals: [`If I were you, I’d...`, `You could... or... ; ... is better because...`, `It’s the best place for..., although...`,
      `If it rains, you might want to...`, `Make sure you... / Don’t forget to...`],
    planFrame: [`On Saturday morning I’m going to...`, `I’m meeting... at... (arrangement)`, `If it rains, I might...`, `I think it’ll be... because...`],
    bingo: [[`cut down on`, `Do or eat less of something.`], [`stick to`, `Continue doing what you planned.`],
      [`take up`, `Start a new hobby.`], [`give up`, `Stop doing something completely.`], [`achieve a goal`, `Do what you wanted to do.`],
      [`get into shape`, `Become fit.`], [`break a resolution`, `Not do what you promised.`], [`keep a resolution`, `Do what you promised.`],
      [`save up for`, `Keep money to buy something later.`], [`work out`, `Do exercise, often in a gym.`],
      [`be more organised`, `Plan your time better.`], [`spend less time on`, `Use less time on something, e.g. your phone.`],
      [`learn a skill`, `Learn to do something new.`], [`volunteer`, `Work for free to help people.`],
      [`make an effort`, `Try hard.`], [`a fresh start`, `A new beginning.`]],
    rankLang: `*I think... will be the hardest because... / ... is much easier than... / It depends on...*`,
    rank: [`give up social media`, `go to the gym three times a week`, `read one book a month`, `cut down on coffee`,
      `learn to cook five new dishes`, `save £50 a month`, `speak only English at home`, `take up a new hobby`,
      `get up at 6 am every day`, `stop eating junk food`],
    rankTypical: [10, 6, 3, 4, 1, 5, 8, 2, 9, 7],
    jamSecs: 60, jamChallenge: 90,
    jam: [`Where I’ll be in five years`, `How I’m going to improve my English this year`, `The perfect year in Brighton`,
      `A resolution I’ll never make`, `Something I’m going to stop doing`, `How the world will be different in 2050`,
      `A skill I’d like to learn, and why`, `My plans for when this course finishes`, `Why resolutions fail`,
      `A goal I’m going to achieve this year`, `The best way to keep a resolution`, `What I’m going to miss about Brighton`],
    cultureLang: `*I completely agree because... / I’m not sure about that... / It depends on... / In my country, people usually...*`,
    culture: [`Resolutions are a waste of time.`, `It’s better to make one big resolution than lots of small ones.`,
      `Most people give up by February.`, `You should tell everyone your resolutions.`, `This year will be better than last year.`,
      `In ten years, we won’t need to learn languages.`, `Studying abroad changes your life.`, `It’s never too late to learn something new.`],
    extension4c: `Class poll: "Who is most likely to keep their resolution?" Students vote and give reasons.`,
    board: [`START`, `Which resolution is going to be hardest for you? Why?`, `What will you be doing this time next year?`,
      `What are you doing this weekend?`, `If you could change one thing this year, what would it be?`, `How are you going to improve your English?`,
      `Go forward 2!`, `What’s a resolution you’ve broken? What happened?`, `Predict 3 things that will happen this year.`,
      `What are you going to do before you leave Brighton?`, `What might you do differently this year?`,
      `Ask the player on your left 3 questions about their plans.`, `Where do you think you’ll live in ten years?`, `Go back 3!`,
      `What new hobby would you like to take up?`, `What’s the best advice for keeping a resolution?`, `What’s something you’ll never do again?`,
      `Describe your perfect weekend in Brighton this year.`, `What will you miss about Brighton?`, `Miss a turn!`,
      `What skill are you going to learn this year?`, `Are you saving up for anything? What?`, `What do you think Brighton will be like in 2050?`,
      `Go forward 1!`, `Persuade the group to make your resolution too.`, `What are you going to cut down on this year?`,
      `Who is going to help you keep your resolution?`, `What’s the most exciting plan you have this year?`,
      `How will you celebrate the end of the course?`, `FINISH! Say "Happy New Year, Brighton!"`],
    boardSecs: 30,
  },
};
// ================= NEW YEAR IN BRIGHTON: ACTIVITY PACKETS =================
const redRun = (text) => new TextRun({ text, color: RED, bold: true });
const redCell = (text, width) => cell([new Paragraph({ children: [redRun(text)] })], width);

function teacherNotes(L, a) {
  const t = [];
  t.push(new Paragraph({ heading: 'Heading1', children: [new TextRun(`Teacher notes: ${a.title(L)}`)] }));
  t.push(Small('Print the whole packet. Keep this page; give students the pages after it.'));
  t.push(table([['Level', `${L.course} (${L.level})`], ['Time', a.time(L)], ['Language', a.focus(L)], ['Prepare', a.prep(L)]], [1800, 8106]));
  t.push(H3('Steps'));
  t.push(...numbered(a.steps(L)));
  t.push(H3('Make it shorter or longer'));
  t.push(...bullets(a.tips(L)));
  if (a.extra) t.push(...a.extra(L));
  return t;
}

function maskWord(tok, mode) {
  const m = tok.match(/^([^A-Za-z0-9’'£]*)([A-Za-z0-9’'£]+)(.*)$/);
  if (!m) return tok;
  const [, pre, core, post] = m;
  if (mode === 'blank') return pre + '_'.repeat(Math.max(3, core.length)) + post;
  return pre + core[0] + '_'.repeat(Math.max(2, core.length - 1)) + post;
}
function vanish(text, li, stage) {
  return text.split(' ').map((tok, i) => {
    const rank = ((li * 31 + i * 17 + 7) * 37) % 100;
    if (stage === 2) return rank < 30 ? maskWord(tok, 'blank') : tok;
    if (stage === 3) return rank < 65 ? maskWord(tok, 'blank') : tok;
    if (stage === 4) return maskWord(tok, 'first');
    return tok;
  }).join(' ');
}
function dialoguePage(L, stage) {
  const labels = { 1: 'Read it together', 2: 'Some words are gone', 3: 'Most words are gone', 4: 'Only the first letters', 5: 'From memory!' };
  const [na, nb] = L.dialogueNames;
  return [
    H1(`Vanishing dialogue ${stage}/5: ${labels[stage]}`),
    Small(`${L.dialogueSetting} Practise with your partner, then swap roles. Turn the page when you’re ready.`),
    table(L.dialogue.map(([w, t], i) => [w === 'A' ? na : nb, stage === 5 ? '____________________________________________________' : vanish(t, i, stage)]),
      [1500, 8406], { rowHeight: stage === 5 ? 700 : 560 }),
  ];
}

const bingoGrid = () => new Table({
  width: { size: 6300, type: WidthType.DXA }, columnWidths: [2100, 2100, 2100], alignment: AlignmentType.CENTER,
  rows: [0, 1, 2].map(() => new TableRow({ height: { value: 1300, rule: HeightRule.EXACT }, children: [0, 1, 2].map(() => cell('', 2100, { border: { style: BorderStyle.SINGLE, size: 12, color: PURPLE } })) })),
});

const ACTIVITIES = [
  {
    num: '01', file: 'Find someone who', title: () => 'Find someone who: New Year, new me',
    time: () => '20–25 minutes', focus: (L) => L.future, prep: () => 'One worksheet per student.',
    steps: (L) => [
      `Drill the question forms: ${L.fswDrill}`,
      'Students turn each prompt into a question. Check two or three together.',
      'Students stand up and mingle. When someone says yes, they write the name and ask a follow-up question.',
      'Stop after about 12 minutes. Feedback: "Who is going to save money? How?" Board useful language.',
    ],
    tips: () => ['Shorter: students only need 5 names.', 'Longer: students report back in pairs: *"Maria wants to get fit."*', 'Warmer: the *New Year’s resolutions vocabulary* Kahoot.'],
    student: (L) => [H1('Find someone who: New Year, new me', true), NameLine(), P(L.fswIntro),
      table([['Find someone who...', 'Name', 'Extra information'], ...L.fsw.map((f) => [f, '', ''])], [4200, 2200, 3506], { header: true, rowHeight: 700 })],
  },
  {
    num: '02', file: 'Two truths and a lie', title: () => 'Two truths and a lie: my resolutions',
    time: () => '20 minutes', focus: (L) => `${L.future}; follow-up questions.`, prep: () => 'One worksheet per student.',
    steps: (L) => [
      `Model: write your own 3 resolutions on the board (two true, one false), e.g. *${L.ttlModel.join(' ')}*`,
      'The class asks you questions, then votes on the lie. Reveal it.',
      'Students write their own 3 resolutions on the worksheet: two true, one false.',
      'Groups of four. Each student reads their sentences; the others ask questions from the box and guess the lie.',
      'One point for each correct guess, and one point for a student who tricks everyone.',
    ],
    tips: () => ['Shorter: pairs instead of groups.', 'Longer: mingle with the whole class and keep score.'],
    student: (L) => [H1('Two truths and a lie: my resolutions', true), NameLine(),
      P('Write 3 New Year’s resolutions. **Two are true. One is a lie!** Your group asks questions and guesses the lie.'),
      box([P('**Example**'), ...L.ttlModel.map((m, i) => P(`${i + 1}. ${m}`, { spacing: { after: 40 } }))]),
      H3('My resolutions'), ...L.ttlFrames.flatMap((f, i) => [P(`${i + 1}. ${f}`, { spacing: { after: 240 } })]),
      P('The lie is number: ____'),
      box([P('**Questions to ask**'), ...bullets(L.ttlQuestions)]),
      H3('Score'), table([['Name', 'My guess', 'Right?'], ['', '', ''], ['', '', ''], ['', '', '']], [4000, 3000, 2906], { header: true, rowHeight: 600 })],
  },
  {
    num: '03', file: 'Interview my year in Brighton', title: (L) => `Interview: ${L.interviewTitle.toLowerCase()}`,
    time: () => '35–40 minutes', focus: (L) => `${L.future}; reporting in the third person.`, prep: () => 'One worksheet per student.',
    steps: (L) => [
      'Lead-in: board *This year in Brighton: one word*. Each student says one word.',
      'Pairs: A asks all 8 questions and makes notes; then swap.',
      `New groups of four: each student tells the group about their partner. Model on the board: ${L.reportModel}`,
      'Listeners ask one question each.',
      'Error clinic: board 5 errors you heard (no names); teams correct them.',
    ],
    tips: () => ['Shorter: choose 4 questions.', 'Longer: class survey; tally the most popular plan or resolution.'],
    student: (L) => [H1(`Interview: ${L.interviewTitle.toLowerCase()}`, true), NameLine(), P(L.interviewIntro), P("My partner's name: ______________________"),
      table([['Question', 'Notes'], ...L.interview.map((q) => [q, ''])], [5000, 4906], { header: true, rowHeight: 900 })],
  },
  {
    num: '04', file: 'Vanishing dialogue', title: () => 'Vanishing dialogue: New Year’s Day',
    time: () => '25–30 minutes', focus: (L) => `${L.future}; pronunciation and fluency.`,
    prep: () => 'One copy of the 5 dialogue pages per pair (or project them one at a time).',
    steps: (L) => [
      'Read page 1 aloud. Students repeat each line (choral drill), then focus on stress and linking.',
      'Pairs read page 1: A and B, then swap roles.',
      'Pairs turn to page 2 and say the dialogue again, filling the gaps from memory. Swap roles.',
      'Continue to pages 3, 4 and 5. Page 5 has only the speakers’ names.',
      'Two or three pairs perform it to the class from memory.',
      `Follow-up: ${L.dialogueFollow}`,
    ],
    tips: () => ['Shorter: stop at page 4.', 'Projecting: show one page at a time; the class says it together before pairs practise.', 'Fun: perform it in a different mood (excited, tired, whispering).'],
    extra: () => [P('**The full dialogue** is on the first student page (1/5).')],
    student: (L) => [1, 2, 3, 4, 5].flatMap((n) => dialoguePage(L, n)),
  },
  {
    num: '05', file: 'My Brighton bucket list', title: () => 'My Brighton bucket list',
    time: () => '30–40 minutes', focus: (L) => `Brighton places; yes/no questions; ${L.future}.`,
    prep: () => 'One copy per student. Optional: Google Maps open on Brighton on the classroom screen.',
    steps: (L) => [
      'Optional warmer: the *Brighton: places* Kahoot.',
      'Guess the place: in pairs, A secretly chooses a place card; B asks up to 10 yes/no questions. Swap.',
      `Bucket list: each student chooses 5 things they are going to do in Brighton this year and writes why. Frame: *${L.bucketFrame}*`,
      'Groups of four compare lists and agree on a group top 5.',
      'Class top 10: groups report; tally the most popular on the board. Optional: show each place on Google Maps.',
    ],
    tips: () => ['Shorter: skip Guess the place.', 'Longer: groups plan when and how they’ll do their top 3 (*"We’re going to... on Saturday."*).'],
    student: (L) => [H1('Guess the place', true), P('Choose a secret place. Your partner asks up to 10 yes/no questions to guess it.'),
      cards(PLACES, 3, 1000, 28), spacer(), box([P('**Question starters**'), ...bullets(L.questionStarters)]),
      H1('My Brighton bucket list', true), NameLine(), P('5 things I’m going to do in Brighton this year.'),
      box([P(`**Ideas:** ${BUCKET_IDEAS}`)]), spacer(),
      table([L.bucketCols, ...[1, 2, 3, 4, 5].map(() => ['', '', ''])], [3000, 4406, 2500], { header: true, rowHeight: 800 }),
      spacer(), P(`**Say it:** *${L.bucketFrame}*`)],
  },
  {
    num: '06', file: 'Plan your first weekend', title: (L) => `Role-play: ${L.plannerTitle.toLowerCase()}`,
    time: () => '40–45 minutes', focus: (L) => `Recommendations; ${L.future}.`,
    prep: () => 'Cut up one set of tourist cards per 6 students. One planner sheet per student.',
    steps: (L) => [
      `Model: you are a newcomer; a strong student is a local. Newcomer: ${L.touristSay}`,
      'Board the locals’ useful language.',
      'Half the class are newcomers (cards), half are locals. Newcomers talk to three locals and note each idea. Swap halfway.',
      'Newcomers complete the weekend planner and tell a partner about their plans.',
      'Class awards: best rainy-day plan, cheapest weekend, best food.',
    ],
    tips: () => ['Shorter: talk to two locals only.', 'Longer: pairs present their weekend like a TV advert.'],
    student: (L) => [H1('Newcomer role cards', true), Small('Cut along the dashed lines. One card per newcomer.'),
      cards(L.tourists.map((t) => `**NEWCOMER** — ${t}`), 2, 1600, 24), spacer(),
      box([P('**Locals: useful language**'), ...bullets(L.locals)]),
      H1(`My plan: ${L.plannerTitle.toLowerCase()}`, true), NameLine(),
      table([['Local’s name', 'Their idea'], ['', ''], ['', ''], ['', '']], [3000, 6906], { header: true, rowHeight: 800 }),
      H3('My weekend'), ...L.planFrame.flatMap((f) => [P(f, { spacing: { after: 0 } }), WriteLine(), spacer()])],
  },
  {
    num: '07', file: 'Resolutions bingo', title: () => 'Resolutions bingo',
    time: () => '15 minutes', focus: () => 'New Year’s resolution phrases.', prep: () => 'One sheet per 2 students (cut in half).',
    steps: () => [
      'Students choose 9 phrases from the box and write them in their grid.',
      'Read the definitions below in any order. Students cross out the phrase.',
      'First full line shouts *"Happy New Year!"* and reads the phrases back to check.',
      'Play again with the second grid. Then each student chooses one phrase as their real resolution.',
    ],
    tips: () => ['Warmer: the *New Year’s resolutions vocabulary* Kahoot.', 'Extension: students make 3 sentences about their plans with the phrases.'],
    extra: (L) => [H3('Definitions to read out'), ...L.bingo.map(([w, d]) => P(`**${w}**: ${d}`, { spacing: { after: 30 } }))],
    student: (L) => [H1('Resolutions bingo', true),
      P('Choose 9 phrases from the box. Write them in your grid. Listen and cross them out. A full line? Shout **"Happy New Year!"**'),
      box([P(L.bingo.map((b) => b[0]).join('  ·  '), { align: AlignmentType.CENTER, run: { size: 23 } })]), spacer(), bingoGrid(),
      new Paragraph({ spacing: { before: 300, after: 300 }, border: { bottom: dashed }, children: [new TextRun({ text: '✂  cut here', size: 16, color: '888888' })] }),
      P('Choose 9 phrases. Write them in your grid.'), bingoGrid()],
  },
  {
    num: '08', file: 'Rank the resolutions', title: () => 'Rank the resolutions', answers: true,
    time: () => '20–25 minutes', focus: (L) => `Comparatives and opinions; ${L.future}.`,
    prep: () => 'One worksheet per student. The ANSWERS file has a typical ranking (in red) to compare with; there is no single right answer.',
    steps: (L) => [
      'Read the 10 resolutions together and check the meaning.',
      'Alone: students rank them 1 (easiest to keep) to 10 (hardest) in the "Me" column.',
      `Groups of three agree on one ranking in the "My group" column. Language: ${L.rankLang}`,
      'Two groups compare and explain their differences.',
      'Show the typical ranking from the ANSWERS file. "Do you agree?"',
    ],
    tips: () => ['Shorter: rank only the top 3 and bottom 3.', 'Longer: each student chooses one resolution from the list and says how they’re going to keep it.'],
    student: (L, ans) => [H1(`Rank the resolutions${ans ? ' (typical ranking)' : ''}`, true), NameLine(),
      P('Which resolutions are easy to keep? Which are hard? Rank them: **1 = easiest**, **10 = hardest**.'),
      new Table({
        width: { size: PORTRAIT_W, type: WidthType.DXA }, columnWidths: [5106, 1600, 1600, 1600],
        rows: [
          new TableRow({ tableHeader: true, children: ['Resolution', 'Me', 'My group', 'Typical'].map((h, i) => cell(h, [5106, 1600, 1600, 1600][i], { fill: PURPLE, run: { bold: true, color: 'FFFFFF' } })) }),
          ...L.rank.map((r, i) => new TableRow({ height: { value: 620, rule: HeightRule.ATLEAST }, children: [cell(r, 5106), cell('', 1600), cell('', 1600), ans ? redCell(String(L.rankTypical[i]), 1600) : cell('', 1600)] })),
        ],
      }),
      spacer(), box([P(`**Useful language:** ${L.rankLang}`)])],
  },
  {
    num: '09', file: 'Just a minute', title: () => 'Just a minute: my future',
    time: () => '20–25 minutes', focus: (L) => `Fluency; ${L.future}.`, prep: () => 'Cut up one set of topic cards per group of four.',
    steps: (L) => [
      `Model: talk about a card for ${L.jamSecs} seconds without stopping.`,
      `Groups of four. One student takes a card and speaks for ${L.jamSecs} seconds (${L.jamChallenge} for strong students).`,
      'The others listen for long stops, then each asks one question.',
      'Rotate until everyone has spoken 2 or 3 times.',
      'Each group’s best speaker talks in front of the class.',
    ],
    tips: () => ['Easier: give 1 minute of thinking time first.', 'Harder: listeners can "buzz" for repetition and take over the topic.'],
    student: (L) => [H1('Just a minute: my future', true),
      Small(`Cut up. Speak for ${L.jamSecs} seconds without stopping (${L.jamChallenge} seconds for a challenge!).`), cards(L.jam, 3, 1600, 24)],
  },
  {
    num: '10', file: 'Four corners', title: () => 'Four corners: resolution opinions',
    time: () => '25–40 minutes', focus: () => 'Opinions; agreeing and disagreeing; the future.',
    prep: () => 'Print the 4 signs and stick one in each corner of the room. The statements are below (read them aloud).',
    steps: (L) => [
      `Board the language: ${L.cultureLang}`,
      'Read a statement. Students walk to the corner that matches their opinion.',
      'Students explain their choice to the others in their corner; then one person from each corner tells the class.',
      'Repeat with 5–8 statements.',
    ],
    tips: (L) => ['Shorter: use 4 statements.', `Extension: ${L.extension4c}`],
    extra: (L) => [H3('Statements to read aloud'), ...numbered(L.culture)],
    student: () => CORNERS.map((c) => new Paragraph({ pageBreakBefore: true, spacing: { before: 3600 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: c, bold: true, size: 150, color: PURPLE })] })),
  },
  {
    num: '11', file: 'Board game', title: () => 'Board game: my year ahead', landscape: true,
    time: () => '30–35 minutes', focus: (L) => `Review of the week; ${L.future}.`,
    prep: () => 'One board per group of 3–4 (A3 if possible), a dice and counters.',
    steps: (L) => [
      'Groups of 3–4 with a board, a dice and counters.',
      `Roll and move. Do the task on the square and speak for at least ${L.boardSecs} seconds. The group asks one question.`,
      'Can’t do it? Go back 2. Monitor and note errors.',
      'Error clinic: board 5 errors; teams correct them.',
      'Finish with the *Brighton & New Year: big review* Kahoot.',
    ],
    tips: () => ['Shorter: the first player to square 20 wins.', 'No dice? Use a dice app on a phone.'],
    student: (L) => {
      const COLS = 6, ROWS = 5; const w = Math.floor(LANDSCAPE_W / COLS); const rows = [];
      for (let r = 0; r < ROWS; r++) {
        const idx = Array.from({ length: COLS }, (_, c) => r * COLS + c); if (r % 2 === 1) idx.reverse();
        rows.push(new TableRow({ height: { value: 1700, rule: HeightRule.EXACT }, cantSplit: true, children: idx.map((i) => {
          const special = i === 0 || i === 29; const fun = /^(Go|Miss)/.test(L.board[i]);
          return cell([
            new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: String(i + 1), bold: true, size: 28, color: special ? 'FFFFFF' : PURPLE })] }),
            new Paragraph({ children: [new TextRun({ text: L.board[i], size: special ? 24 : 17, bold: special || fun, color: special ? 'FFFFFF' : '222222' })] }),
          ], w, { fill: special ? PURPLE : fun ? 'FBE9F1' : (r + idx.indexOf(i)) % 2 ? 'FFFFFF' : LIGHT, border: { style: BorderStyle.SINGLE, size: 8, color: PURPLE } });
        }) }));
      }
      return [new Paragraph({ heading: 'Heading1', children: [new TextRun('Board game: my year ahead')] }),
        P(`Roll the dice and move. Do the task and speak for at least ${L.boardSecs} seconds. The group asks you one question. Can’t do it? Go back 2. Follow the numbers: the path snakes back and forth.`, { run: { size: 18 } }),
        new Table({ width: { size: w * COLS, type: WidthType.DXA }, columnWidths: Array(COLS).fill(w), rows })];
    },
  },
];

function overview(L) {
  const o = [];
  o.push(new Paragraph({ heading: 'Heading1', children: [new TextRun(`New Year in Brighton: ${L.course} (${L.level})`)] }));
  o.push(P('Eleven speaking activities for the start of the year: New Year’s resolutions and what students are going to do in Brighton. Use them in any order and fit them into your week where you like.'));
  o.push(P(`**Language:** ${L.future}.`));
  o.push(H2('How the packets work'));
  o.push(...bullets([
    'Each activity is its own file. Print the whole file.',
    'The **first page is for you** (teacher notes: time, prep, steps, ideas). The pages after it are for students.',
    'Files ending in **ANSWERS** are the same pages with the answers filled in **in red**.',
  ]));
  o.push(H2('The activities'));
  o.push(table([['#', 'Activity', 'Time', 'Answers?'], ...ACTIVITIES.map((a) => [a.num, a.title(L), a.time(L), a.answers ? 'Yes (red)' : '—'])], [700, 5300, 2200, 1706], { header: true }));
  o.push(H2('Kahoots (shared by Course 2 and Course 3)'));
  o.push(P('Each link opens the kahoot’s page: press **Start** to host, or **Duplicate** to copy it into your account. They work once the kahoots are set to public.'));
  o.push(new Table({
    width: { size: PORTRAIT_W, type: WidthType.DXA }, columnWidths: [4000, 3200, 2706],
    rows: [
      new TableRow({ tableHeader: true, children: ['Kahoot', 'Good with', 'Link'].map((h, i) => cell(h, [4000, 3200, 2706][i], { fill: PURPLE, run: { bold: true, color: 'FFFFFF' } })) }),
      ...KAHOOTS.map(([n, g, u]) => new TableRow({ children: [cell(n, 4000), cell(g, 3200), cell([new Paragraph({ children: [new ExternalHyperlink({ link: `https://create.kahoot.it/details/${u}`, children: [new TextRun({ text: 'Open the kahoot', style: 'Hyperlink' })] })] })], 2706)] })),
    ],
  }));
  o.push(H2('Low prep'));
  o.push(...bullets(['Things to cut: place cards, newcomer cards, Just a minute cards, bingo sheets.', 'Bring dice and counters for the board game.', 'Nothing to record and no maps to print.']));
  return o;
}

async function build() {
  const outDir = process.argv[2];
  for (const key of ['c2', 'c3']) {
    const L = LEVELS[key];
    const dir = path.join(outDir, L.dir);
    fs.mkdirSync(dir, { recursive: true });
    const sub = `New Year in Brighton · ${L.course} (${L.level})`;
    const pSec = (children) => ({ properties: portrait, headers: { default: header(sub) }, footers: { default: footer() }, children });
    const lSec = (children) => ({ properties: landscape, headers: { default: landscapeHeader(sub) }, footers: { default: footer() }, children });
    const save = async (name, sections) => {
      const doc = new Document({ creator: 'Kings Brighton', title: name, styles, numbering, sections });
      fs.writeFileSync(path.join(dir, `${name}.docx`), await Packer.toBuffer(doc));
    };
    await save('00 Overview', [pSec(overview(L))]);
    for (const a of ACTIVITIES) {
      const mk = (ans) => a.landscape ? [pSec(teacherNotes(L, a)), lSec(a.student(L, ans))] : [pSec([...teacherNotes(L, a), ...a.student(L, ans)])];
      await save(`${a.num} ${a.file}`, mk(false));
      if (a.answers) await save(`${a.num} ${a.file} - ANSWERS`, mk(true));
    }
  }
  console.log('built');
}
build().catch((e) => { console.error(e); process.exit(1); });

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

// ================= GHOSTS (BBC) SERIES 1: CONTENT =================
const SHOW = 'Ghosts (BBC) · Series 1';

const CHARACTERS = [
  ['Alison', 'Alive. She gets a big old house. After an accident, she can **see and hear** the ghosts.'],
  ['Mike', 'Alive. Alison’s partner. He **can’t** see the ghosts.'],
  ['Robin', 'A **caveman**. He is the oldest ghost. He likes chess and the moon.'],
  ['Mary', 'From the **1600s**. People thought she was a witch.'],
  ['Humphrey', 'A Tudor man. He has **no head**! He carries it.'],
  ['Kitty', 'From the **1700s**. Very friendly and happy. She wants a best friend.'],
  ['Thomas', 'A **poet** from the 1800s. Very romantic. He loves Alison.'],
  ['Lady Fanny Button', 'An **Edwardian** lady (early 1900s). Very posh. Her old house!'],
  ['The Captain', 'An army officer from the **Second World War**. He likes giving orders.'],
  ['Pat', 'A **scout leader** from the 1980s. Very kind. There’s an arrow in his neck.'],
  ['Julian', 'A **politician** from the 1990s. He has **no trousers**! He can push things a little.'],
];

// Each episode: vocab [word, simple meaning], watch-for questions, after T/F, after short answers, speaking, answers.
const EPISODES = [
  {
    n: 1, title: 'Who Do You Think You Are?', kahoot: 'd92761c0-77b2-4059-a838-4af20c460dff',
    story: 'Alison and Mike get a big old house. But the house is full of ghosts, and the ghosts don’t want them there.',
    vocab: [['inherit', 'get money or a house from a person who died'], ['ghost', 'a dead person you can (sometimes) see'],
      ['haunt', 'when a ghost stays in a place and scares people'], ['scare', 'make someone afraid'],
      ['push', 'move something away with your hands'], ['fall', 'go down to the ground fast'],
      ['wake up', 'stop sleeping'], ['falling apart', 'very old and broken']],
    predict: 'Look at the character cards. Which ghost is your favourite? Why?',
    watch: ['What is the house like? Old or new?', 'What does Robin do to the lights?', 'Who pushes Alison?'],
    tf: [['Alison and Mike buy the house.', 'F (they inherit it: a relative leaves it to them)'],
      ['The ghosts want Alison and Mike to stay.', 'F (they want them to leave)'],
      ['Alison falls out of a window.', 'T'],
      ['At the end, Alison can see the ghosts.', 'T']],
    short: [['Where does Alison wake up?', 'In hospital.'], ['Which ghost does Alison see at the end? (Clue: no head!)', 'Humphrey.']],
    speak: 'Would you like to live in a big old house? Why? Why not?',
  },
  {
    n: 2, title: 'Gorilla War', kahoot: '7140d823-29b1-48d1-941a-4c376688ecd5',
    story: 'Alison is home from hospital. Now she can see and hear the ghosts, and they never stop talking to her!',
    vocab: [['believe', 'think something is true'], ['hear', 'get sounds with your ears'], ['doctor', 'a person who helps sick people'],
      ['war', 'a big fight between two groups'], ['annoying', 'something that makes you a little angry'],
      ['boiler', 'a machine that makes hot water'], ['explode', 'break with a big BANG'], ['warn', 'tell someone about danger']],
    predict: 'Alison can see ghosts. Mike can’t. What will Mike think?',
    watch: ['What does Mike think: are the ghosts real?', 'Where does Alison go to stop the ghosts?', 'What is the problem in the cellar?'],
    tf: [['Mike believes Alison at the start.', 'F (he thinks it’s because of the accident)'],
      ['The Captain wants a “war” with Alison.', 'T'],
      ['Alison goes back to the hospital.', 'T'],
      ['At the end, Alison and Mike leave the house.', 'F (they stay)']],
    short: [['What machine is dangerous?', 'The (old) boiler.'], ['Does Mike believe Alison at the end?', 'Yes.']],
    speak: 'If you could see ghosts, would you tell your friends? Would they believe you?',
  },
  {
    n: 3, title: 'Happy Death Day', kahoot: '6cc112cd-0d62-4d2a-895c-e29d3f1af1dd',
    story: 'Every year, on the day Pat died, his family come to the house. This year, Pat asks Alison to help.',
    vocab: [['anniversary', 'the same date every year after something happened'], ['message', 'words you send to someone'],
      ['wife', 'the woman a man is married to'], ['best friend', 'your number one friend'],
      ['grandson', 'the son of your son or daughter'], ['visit', 'go to see a person or a place'],
      ['upset', 'sad and a little angry'], ['worried', 'thinking something bad will happen']],
    predict: 'Pat’s family come to see the place where he died. How do you think Pat feels?',
    watch: ['Who comes to the house?', 'What does Pat ask Alison to do?', 'What is Mike worried about?'],
    tf: [['Pat’s family visit once a year.', 'T'],
      ['Pat asks Alison to give his family a message.', 'T'],
      ['Pat’s wife married his best friend.', 'T'],
      ['Pat is very happy all the time in this episode.', 'F (he is upset)']],
    short: [['What is the name of Pat’s grandson?', 'Pat (he has the same name).'], ['Which room is Mike worried about?', 'The bathroom.']],
    speak: 'Does your family have special days every year? What do you do?',
  },
  {
    n: 4, title: 'Free Pass', kahoot: '977c2dd1-0962-4238-b0b8-df864b29a6ff',
    story: 'Alison and Mike need money, so a film company makes a film in the house. The ghosts are very excited, and some are angry.',
    vocab: [['rent', 'pay money to use something for a short time'], ['film crew', 'the people who make a film'],
      ['actor', 'a person who plays a part in a film'], ['period drama', 'a film or TV story set in the past'],
      ['romantic', 'about love'], ['floor', 'the part of a room you walk on'],
      ['fall through', 'break something and go down through it'], ['bet', 'money you risk on a race or game']],
    predict: 'A film crew comes to the house. Which ghost will be happy? Which ghost will be angry?',
    watch: ['Why do Alison and Mike say yes to the film company?', 'What does the crew move that makes Robin angry?', 'What happens to the actor?'],
    tf: [['Alison and Mike need money.', 'T'],
      ['Kitty loves the romantic story.', 'T'],
      ['The film company pays Alison and Mike.', 'F (they leave without paying)'],
      ['Alison and Mike win some money on a horse race.', 'T (£100)']],
    short: [['What is the film about: the past or the future?', 'The past (it’s a period drama).'], ['How much money do they win?', '£100.']],
    speak: 'Would you like to be in a film? What kind of film?',
  },
  {
    n: 5, title: 'Moonah Ston', kahoot: '625b3d26-457d-4924-921a-b13e18c6dae8',
    story: 'The neighbours say part of the land is theirs. Alison and Mike invite them to dinner. Robin has a special night with the moon.',
    vocab: [['neighbour', 'a person who lives near you'], ['land', 'the ground around a house'],
      ['dinner party', 'a meal at home with guests'], ['guest', 'a person you invite'],
      ['manners', 'polite ways to behave'], ['etiquette', 'the rules of good manners'],
      ['moon', 'the big light in the sky at night'], ['eclipse', 'when the moon or sun goes dark for a short time']],
    predict: 'The ghosts help with the dinner party. Is that a good idea?',
    watch: ['What do the neighbours want?', 'Is the dinner party good or bad?', 'What does Robin want to see in the sky?'],
    tf: [['The neighbours want some of the land.', 'T'],
      ['The ghosts give advice about manners.', 'T'],
      ['The dinner party is perfect.', 'F (it goes badly)'],
      ['Robin is excited about the moon.', 'T']],
    short: [['What does “Moonah Ston” mean?', 'Moon stone (it’s how Robin says it).'], ['What TV show do the ghosts watch on DVD?', 'Friends.']],
    speak: 'What are good manners at dinner in your country?',
  },
  {
    n: 6, title: 'Getting Out', kahoot: 'ab472845-5a29-4e5c-88d0-435027dc6912',
    story: 'Alison and Mike are tired and have no money. A hotel company wants to buy the house. Should they sell?',
    vocab: [['sell', 'give something for money'], ['hotel', 'a place where you pay to sleep'],
      ['offer', 'the money someone says they will pay'], ['price', 'how much something costs'],
      ['cellar', 'a room under the house'], ['pit', 'a big hole in the ground'],
      ['plan', 'what you want to do next'], ['leader', 'the person who tells the group what to do']],
    predict: 'Will Alison and Mike sell the house? What do the ghosts want?',
    watch: ['Who wants to buy the house?', 'Do all the ghosts agree with the Captain?', 'Where does Fiona fall?'],
    tf: [['Alison and Mike have a lot of money.', 'F (they have no money)'],
      ['Fiona offers a good, fair price.', 'F (she offers too little)'],
      ['Fiona falls into a pit in the cellar.', 'T'],
      ['At the end, Fiona still wants to buy the house.', 'F (she leaves)']],
    short: [['Who planned Fiona’s fall?', 'The Captain.'], ['Do Alison and Mike sell the house?', 'No.']],
    speak: 'Which ghost is your favourite after six episodes? Why?',
  },
];

// ================= STUDENT PACK =================
function charactersPage(first = false) {
  const s = [];
  s.push(H1('Meet the characters', !first));
  s.push(P('Button House is a big old house in the countryside. Lots of ghosts live there. Each ghost died in a different time in history.'));
  s.push(cards(CHARACTERS.map(([n, d]) => `**${n}** — ${d}`), 2, 1150, 21));
  s.push(spacer());
  s.push(P('**Quick quiz.** Who...?  1. has no head? ______  2. has no trousers? ______  3. is the oldest? ______  4. can see the ghosts? ______'));
  return s;
}

function episodeStudent(ep) {
  const s = [];
  s.push(H1(`Episode ${ep.n}: ${ep.title}`));
  s.push(NameLine());
  s.push(box([P(`**The story:** ${ep.story}`)]));
  s.push(H2('Before you watch'));
  s.push(H3('1. Words: match the word to the meaning'));
  const letters = 'abcdefgh';
  const shuffled = ep.vocab.map((v, i) => [v[1], i]).sort((a, b) => (a[0] > b[0] ? 1 : -1));
  s.push(table(ep.vocab.map((v, i) => [`${i + 1}. ${v[0]}`, '____', `${letters[i]}) ${shuffled[i][0]}`]), [2400, 900, 6606]));
  s.push(H3('2. Guess'));
  s.push(P(ep.predict, { spacing: { after: 0 } }));
  s.push(...lines(1));
  s.push(H2('While you watch'));
  s.push(P('Answer these while you watch. Your teacher may stop the episode in the middle.'));
  ep.watch.forEach((q, i) => { s.push(P(`${i + 1}. ${q}`, { spacing: { after: 0 } })); s.push(WriteLine()); });
  s.push(new Paragraph({ heading: 'Heading2', pageBreakBefore: true, children: [new TextRun(`Episode ${ep.n}: after you watch`)] }));
  s.push(H3('True or false?'));
  s.push(table([['', 'Sentence', 'T / F'], ...ep.tf.map(([t], i) => [String(i + 1), t, ''])], [500, 8006, 1400], { header: true }));
  s.push(H3('Answer the questions'));
  ep.short.forEach(([q], i) => { s.push(P(`${i + 1}. ${q}`, { spacing: { after: 0 } })); s.push(WriteLine()); });
  s.push(H3('Talk to your partner'));
  s.push(P(ep.speak));
  return s;
}

function studentPack() {
  const s = [];
  s.push(new Paragraph({ heading: 'Heading1', children: [new TextRun('Ghosts: student worksheets')] }));
  s.push(P(`${SHOW}, episodes 1–6. Two pages per episode: page 1 is before and while you watch, page 2 is after you watch.`));
  s.push(table([
    ['Page', 'Worksheet'], ['2', 'Meet the characters (use with every episode)'],
    ...EPISODES.map((e, i) => [`${3 + 2 * i}\u2013${4 + 2 * i}`, `Episode ${e.n}: ${e.title}`]),
  ], [1200, 8706], { header: true }));
  s.push(...charactersPage());
  EPISODES.forEach((e) => s.push(...episodeStudent(e)));
  return s;
}

// The vocab matching key: each word i matches the letter of its meaning in the sorted list.
function vocabKey(ep) {
  const sorted = ep.vocab.map((v) => v[1]).sort();
  return ep.vocab.map((v, i) => `${i + 1}${'abcdefgh'[sorted.indexOf(v[1])]}`).join(' · ');
}

// ================= TEACHER PACK =================
function teacherPack() {
  const t = [];
  t.push(new Paragraph({ heading: 'Heading1', children: [new TextRun('Ghosts: teacher pack')] }));
  t.push(P(`**Show:** ${SHOW}, 6 episodes of about 28 minutes, on BBC iPlayer. **Level:** strong A1 to A2, ages 16–22. **Lesson:** 45 minutes per episode.`));
  t.push(P('Native-speed comedy is above A2, so turn on **English subtitles**, pre-teach the 8 key words, and focus on who / what / where rather than the jokes.'));
  t.push(H2('45-minute lesson plan (same every episode)'));
  t.push(...numbered([
    '**Characters (5 min).** Episode 1: hand out *Meet the characters* and do the quick quiz. Later episodes: "Who can you remember?" Point at the cards.',
    '**Words (5 min).** Students match the 8 words to their meanings in pairs. Check, then drill pronunciation.',
    '**Guess (2 min).** Read *The story* box together and ask the guess question.',
    '**Watch (28 min).** Students answer the *While you watch* questions. Optional: pause at about the halfway point to check questions 1–2.',
    '**After (5 min).** True or false and the two questions, in pairs.',
    '**Kahoot or speaking (5 min).** Run the episode’s 10-question Kahoot (story + key words), or use the *Talk to your partner* question. Give the other as homework.',
  ]));
  t.push(H2('Kahoots'));
  t.push(P('One 10-question Kahoot (story + key words) per episode. These links open each kahoot’s page (press **Start** to host, or **Duplicate** to copy it). They work for colleagues once the kahoots are set to public.'));
  t.push(new Table({
    width: { size: PORTRAIT_W, type: WidthType.DXA }, columnWidths: [1500, 4000, 4406],
    rows: [
      new TableRow({ tableHeader: true, children: ['Episode', 'Title', 'Link'].map((h, i) => cell(h, [1500, 4000, 4406][i], { fill: PURPLE, run: { bold: true, color: 'FFFFFF' } })) }),
      ...EPISODES.map((e) => new TableRow({ children: [cell(String(e.n), 1500), cell(e.title, 4000), cell([new Paragraph({ children: [new ExternalHyperlink({ link: `https://create.kahoot.it/details/${e.kahoot}`, children: [new TextRun({ text: 'Open the kahoot', style: 'Hyperlink' })] })] })], 4406)] })),
    ],
  }));
  t.push(H2('Before you teach'));
  t.push(P('These materials were written from published episode guides, not from watching the episodes. **Watch each episode once before class** and check the questions match what’s on screen; change any that don’t. Rated 12: some mild innuendo.'));

  EPISODES.forEach((ep) => {
    t.push(H1(`Episode ${ep.n}: ${ep.title}: answers`));
    t.push(P(`**Plot:** ${ep.story}`));
    t.push(H3('Words (matching)'));
    t.push(P(vocabKey(ep)));
    t.push(H3('While you watch'));
    t.push(...numbered(ep.watch.map((q, i) => `${q} → ${[ep.tf, ep.short].flat().length ? WATCH_ANS[ep.n - 1][i] : ''}`)));
    t.push(H3('True or false'));
    t.push(...numbered(ep.tf.map(([q, a]) => `${q} **${a}**`)));
    t.push(H3('Questions'));
    t.push(...numbered(ep.short.map(([q, a]) => `${q} **${a}**`)));
  });
  return t;
}

const WATCH_ANS = [
  ['**Old and falling apart.**', '**He makes them flicker.**', '**Julian.**'],
  ['**No, he thinks it’s because of her accident.**', '**The hospital.**', '**The old boiler could explode.**'],
  ['**Pat’s family.**', '**Give his family a message.**', '**The bathroom.**'],
  ['**They need money.**', '**His chess table.**', '**He falls through the bedroom floor.**'],
  ['**Some of the land.**', '**Bad.**', '**The moon (an eclipse).**'],
  ['**A hotel company (Fiona is their scout).**', '**No.**', '**Into a pit in the cellar.**'],
];

// ================= BUILD =================
async function build() {
  const sd = new Document({
    creator: 'Kings Brighton', title: 'Ghosts: student worksheets', styles, numbering,
    sections: [{ properties: portrait, headers: { default: header('Ghosts (BBC) · A1–A2 · Student worksheets') }, footers: { default: footer() }, children: studentPack() }],
  });
  const td = new Document({
    creator: 'Kings Brighton', title: 'Ghosts: teacher pack', styles, numbering,
    sections: [{ properties: portrait, headers: { default: header('Ghosts (BBC) · A1–A2 · Teacher pack') }, footers: { default: footer() }, children: teacherPack() }],
  });
  fs.writeFileSync(path.join(OUT, 'Ghosts - Student Worksheets.docx'), await Packer.toBuffer(sd));
  fs.writeFileSync(path.join(OUT, 'Ghosts - Teacher Pack.docx'), await Packer.toBuffer(td));
  console.log('built');
}
build().catch((e) => { console.error(e); process.exit(1); });

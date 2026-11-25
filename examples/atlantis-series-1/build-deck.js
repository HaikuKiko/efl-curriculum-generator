// Atlantis (BBC) episodes 1-3: before-you-watch board slides.
const pptxgen = require('pptxgenjs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const gi = require('react-icons/gi');
const { applyTheme } = require(process.argv[3]);
const EP = Number(process.argv[4]);

const THEME = {
  name: 'Atlantis Aegean', headFontFace: 'Cambria', bodyFontFace: 'Calibri',
  colors: { dk1: '1A1A1A', lt1: 'FFFFFF', dk2: '0E2A47', lt2: 'EEF3F8', accent1: 'C9A227', accent2: '1F6F8B',
    accent3: '0E2A47', accent4: '8FB8CC', accent5: 'B5483B', accent6: '5A7D3A', hlink: '1F6F8B', folHlink: '6D5A8B' },
};
const NAVY = '0E2A47', GOLD = 'C9A227', TEAL = '1F6F8B', ICE = 'EEF3F8';

async function icon(name, color, px = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(gi[name], { color: '#' + color, size: px }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = `Atlantis episode ${EP}: before you watch`;
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: 'EPISODE_TITLE', background: { color: NAVY },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 2.3, w: 8.2, h: 1.6, fontSize: 48, bold: true, color: 'FFFFFF', align: 'left', valign: 'bottom', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.05, w: 8.2, h: 1.2, fontSize: 24, color: GOLD, align: 'left', valign: 'top', margin: 0 }, text: '' } },
    ],
  });
  pres.defineSlideMaster({
    title: 'BOARD', background: { color: 'FFFFFF' },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontSize: 38, bold: true, color: NAVY, margin: 0 }, text: '' } },
    ],
    slideNumber: { x: 12.4, y: 7.0, w: 0.6, h: 0.3, fontSize: 11, color: '6B7B8C' },
  });

  const titleSlide = async (section, title, sub, iconName) => {
    pres.addSection({ title: section });
    const s = pres.addSlide({ masterName: 'EPISODE_TITLE', sectionTitle: section });
    s.addText(title, { placeholder: 'title' });
    s.addText(sub, { placeholder: 'body' });
    s.addShape(pres.shapes.OVAL, { x: 9.6, y: 1.9, w: 3.0, h: 3.0, fill: { color: '163A5E' }, line: { color: GOLD, width: 2 }, objectName: 'icon circle' });
    s.addImage({ data: await icon(iconName, GOLD), x: 10.25, y: 2.55, w: 1.7, h: 1.7, altText: 'Atlantis icon' });
    s.addText('ATLANTIS (BBC)', { x: 0.8, y: 1.6, w: 6, h: 0.5, fontSize: 18, bold: true, color: GOLD, charSpacing: 4, margin: 0, isTextBox: true });
    return s;
  };

  // Four character cards in a 2x2 grid.
  const characterSlide = async (section, title, people) => {
    const s = pres.addSlide({ masterName: 'BOARD', sectionTitle: section });
    s.addText(title, { placeholder: 'title' });
    const pos = [[0.6, 1.55], [6.8, 1.55], [0.6, 4.3], [6.8, 4.3]];
    for (let i = 0; i < people.length; i++) {
      const [x, y] = pos[i]; const [name, desc, ic] = people[i];
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.9, h: 2.45, rectRadius: 0.15, fill: { color: ICE }, line: { color: ICE }, objectName: `card ${name}` });
      s.addShape(pres.shapes.OVAL, { x: x + 0.35, y: y + 0.45, w: 1.5, h: 1.5, fill: { color: NAVY }, line: { color: NAVY } });
      s.addImage({ data: await icon(ic, GOLD), x: x + 0.65, y: y + 0.75, w: 0.9, h: 0.9, altText: name });
      s.addText(name, { x: x + 2.1, y: y + 0.35, w: 3.6, h: 0.6, fontSize: 26, bold: true, color: NAVY, fontFace: 'Cambria', margin: 0, isTextBox: true });
      s.addText(desc, { x: x + 2.1, y: y + 0.95, w: 3.6, h: 1.3, fontSize: 18, color: '1A1A1A', valign: 'top', margin: 0, isTextBox: true });
    }
    return s;
  };

  // Numbered question list with an icon panel on the right.
  const questionSlide = async (section, title, qs, ic, notes) => {
    const s = pres.addSlide({ masterName: 'BOARD', sectionTitle: section });
    s.addText(title, { placeholder: 'title' });
    qs.forEach((q, i) => {
      const y = 1.6 + i * (5.2 / qs.length);
      s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.7, h: 0.7, fill: { color: GOLD }, line: { color: GOLD } });
      s.addText(String(i + 1), { x: 0.6, y, w: 0.7, h: 0.7, fontSize: 22, bold: true, color: NAVY, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
      s.addText(q, { x: 1.55, y: y - 0.1, w: 7.9, h: 0.95, fontSize: 24, color: '1A1A1A', valign: 'middle', margin: 0, isTextBox: true });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.9, y: 1.7, w: 2.8, h: 4.9, rectRadius: 0.2, fill: { color: NAVY }, line: { color: NAVY } });
    s.addImage({ data: await icon(ic, GOLD), x: 10.5, y: 3.35, w: 1.6, h: 1.6, altText: 'illustration' });
    if (notes) s.addNotes(notes);
    return s;
  };

  // Key words as tiles: word + simple meaning.
  const wordsSlide = (section, title, words) => {
    const s = pres.addSlide({ masterName: 'BOARD', sectionTitle: section });
    s.addText(title, { placeholder: 'title' });
    const cols = 4, w = 2.9, h = 2.45, gx = 0.17, gy = 0.3;
    words.forEach(([word, meaning], i) => {
      const x = 0.6 + (i % cols) * (w + gx), y = 1.6 + Math.floor(i / cols) * (h + gy);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: i % 2 ? ICE : 'E3ECF4' }, line: { color: 'E3ECF4' } });
      s.addText(word, { x: x + 0.2, y: y + 0.25, w: w - 0.4, h: 0.7, fontSize: 26, bold: true, color: TEAL, fontFace: 'Cambria', margin: 0, isTextBox: true });
      s.addText(meaning, { x: x + 0.2, y: y + 1.0, w: w - 0.4, h: 1.3, fontSize: 17, color: '1A1A1A', valign: 'top', margin: 0, isTextBox: true });
    });
    return s;
  };

  // ---------- Episode 1 ----------
  if (EP === 1) {
  const E1 = 'Episode 1';
  await titleSlide(E1, 'Episode 1: The Earth Bull', 'Before you watch: meet the characters', 'GiGreekTemple');
  await characterSlide(E1, 'Meet the characters', [
    ['Jason', 'A young man from our time. He is looking for his father.', 'GiDivingHelmet'],
    ['Pythagoras', 'Young and very clever. He loves maths and triangles.', 'GiTriangleTarget'],
    ['Hercules', 'Older. He says he is very strong. He loves food and wine.', 'GiMuscleUp'],
    ['Ariadne', 'The princess. She is the king’s daughter. Kind and brave.', 'GiWomanElfFace'],
  ]);
  await characterSlide(E1, 'Meet the characters', [
    ['King Minos', 'The king of Atlantis. Ariadne’s father.', 'GiCrown'],
    ['Queen Pasiphae', 'The queen. Ariadne’s stepmother. Be careful of her!', 'GiQueenCrown'],
    ['The Oracle', 'A wise woman. She can see the future.', 'GiCrystalBall'],
    ['The Minotaur', 'A monster: half man, half bull. It lives under the city.', 'GiMinotaur'],
  ]);
  await questionSlide(E1, 'Who is it?', [
    'She can see the future.', 'He loves maths.', 'He is looking for his father.',
    'She is the king’s daughter.', 'Half man, half bull!', 'He loves food and wine.',
  ], 'GiThink', 'Answers: 1 the Oracle, 2 Pythagoras, 3 Jason, 4 Ariadne, 5 the Minotaur, 6 Hercules.');
  wordsSlide(E1, 'Words for episode 1', [
    ['hero', 'a brave person that people admire'], ['monster', 'a big, scary creature'], ['brave', 'not afraid'],
    ['palace', 'the big house of a king and queen'], ['maze', 'a place with many paths; it’s easy to get lost'],
    ['thread', 'a long, thin string'], ['the future', 'the time that comes after now'], ['sacrifice', 'to give someone to a god or monster'],
  ]);
  await questionSlide(E1, 'Before you watch: talk', [
    'Do you know any stories from ancient Greece?', 'What is Atlantis? Where do you think it is?',
    'Jason comes from our time. What will be strange for him?', 'Would you like to be a hero? Why? Why not?',
  ], 'GiTalk', 'Pair work, 3–5 minutes, then share a few ideas with the class.');

  }
  // ---------- Episode 2 ----------
  if (EP === 2) {
  const E2 = 'Episode 2';
  await titleSlide(E2, 'Episode 2: A Girl by Any Other Name', 'Before you watch', 'GiForest');
  await questionSlide(E2, 'Remember episode 1?', [
    'How did Jason get to Atlantis?', 'Who are Jason’s two new friends?',
    'What monster did Jason fight?', 'Who helped him find the way out? How?',
  ], 'GiThink', 'Answers: 1 He was diving in the sea, looking for his father, and went through a white light. 2 Pythagoras and Hercules. 3 The Minotaur. 4 Ariadne, with a thread.');
  wordsSlide(E2, 'Words for episode 2', [
    ['forest', 'a big place with lots of trees'], ['missing', 'lost; nobody knows where they are'], ['daughter', 'your female child'],
    ['god', 'in old stories, a very powerful being'], ['priestess', 'a woman who leads people in praying to a god'],
    ['curse', 'bad magic that hurts someone'], ['escape', 'get away from a dangerous place'], ['rescue', 'save someone from danger'],
  ]);
  await questionSlide(E2, 'Before you watch: think and talk', [
    'An old man’s daughter is missing in the forest. Who do you think took her?',
    'Would you help a stranger find his daughter? Why?',
    'In the old Greek stories, Medusa has snakes for hair. Can you guess her story in this episode?',
  ], 'GiMedusaHead', 'Introduce Medusa: a new character in this episode. Don’t confirm guesses; check after watching.');

  }
  // ---------- Episode 3 ----------
  if (EP === 3) {
  const E3 = 'Episode 3';
  await titleSlide(E3, 'Episode 3: A Boy of No Consequence', 'Before you watch', 'GiBull');
  await questionSlide(E3, 'Remember episode 2?', [
    'Who was missing?', 'Who took her?', 'Which new character did they meet in the forest?', 'Why was Jason worried at the end?',
  ], 'GiThink', 'Answers: 1 Demetria, an old man’s daughter. 2 The Maenads (followers of the god Dionysus). 3 Medusa. 4 He thought the priestess had cursed Medusa.');
  wordsSlide(E3, 'Words for episode 3', [
    ['market', 'a place outside where people buy and sell things'], ['trial', 'when a court decides if someone did something wrong'],
    ['guilty', 'you did the bad thing'], ['slave', 'a person who must work for no money and is not free'],
    ['bull', 'a big, strong male cow'], ['leap', 'jump high or far'], ['spell', 'words that make magic'], ['freedom', 'being free'],
  ]);
  await questionSlide(E3, 'Before you watch: think and talk', [
    'Jason sees a rich man hurting an old man in the market. What should he do?',
    'In ancient Crete, young people jumped over bulls. Would you try it?',
    'What happens to people who break the law in your country?',
  ], 'GiBull', 'Bull leaping was a real sport in ancient Crete (Minoan wall paintings show it).');

  }
  await pres.writeFile({ fileName: process.argv[2] });
  await applyTheme(process.argv[2], THEME);
  console.log('written', process.argv[2]);
})().catch((e) => { console.error(e); process.exit(1); });

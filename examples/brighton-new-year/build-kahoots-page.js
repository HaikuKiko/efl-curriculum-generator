async function buildKahoots() {
  const outDir = process.argv[2];
  for (const key of ['c2', 'c3']) {
    const L = LEVELS[key];
    const dir = path.join(outDir, L.dir);
    fs.mkdirSync(dir, { recursive: true });
    const W = [3800, 3000, 3106];
    const kids = [
      new Paragraph({ heading: 'Heading1', children: [new TextRun('Kahoots: New Year in Brighton')] }),
      P(`${L.course} (${L.level}). These six Kahoots are shared by Course 2 and Course 3.`),
      P('Tap or click a link to open the kahoot’s page, then press **Start** to play it with your class, or **Duplicate** to copy it into your own Kahoot account. The links work once the kahoots are set to public.'),
      new Table({
        width: { size: PORTRAIT_W, type: WidthType.DXA }, columnWidths: W,
        rows: [
          new TableRow({ tableHeader: true, children: ['Kahoot', 'Use it with', 'Link'].map((h, i) => cell(h, W[i], { fill: PURPLE, run: { bold: true, color: 'FFFFFF' } })) }),
          ...KAHOOTS.map(([n, g, u]) => new TableRow({ height: { value: 700, rule: HeightRule.ATLEAST }, children: [
            cell(`**${n}**`, W[0]), cell(g, W[1]),
            cell([new Paragraph({ children: [new ExternalHyperlink({ link: `https://create.kahoot.it/details/${u}`, children: [new TextRun({ text: `create.kahoot.it/details/${u.slice(0, 8)}…`, style: 'Hyperlink' })] })] })], W[2])] })),
        ],
      }),
    ];
    const doc = new Document({ creator: 'Kings Brighton', title: 'Kahoots', styles, numbering,
      sections: [{ properties: portrait, headers: { default: header(`New Year in Brighton · ${L.course} (${L.level})`) }, footers: { default: footer() }, children: kids }] });
    fs.writeFileSync(path.join(dir, 'Kahoots.docx'), await Packer.toBuffer(doc));
  }
  console.log('kahoots built');
}
buildKahoots().catch((e) => { console.error(e); process.exit(1); });

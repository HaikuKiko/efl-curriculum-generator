# Brighton Skills Week (example output)

Hand-made first output for a 2A Elementary class (A1+ to A2). It's the model for
what the generator should produce.

- `Brighton Week - Student Pack.docx`: worksheets, cut-up cards, four corners
  signs and the board game. No answers. Kings Brighton logo in the header.
- `Brighton Week - Teacher Pack.docx`: week overview, Kahoot links, lesson
  plans, listening script and answer keys.
- `build-packs.js`: the script that builds both files with the `docx` npm package:

  ```bash
  node build-packs.js <output-folder> kings-logo.png
  ```

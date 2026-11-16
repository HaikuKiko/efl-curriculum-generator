# Project notes

What has been done so far, why, and what is still open. Newest work at the
bottom of each section.

## Purpose

A tool for EFL teachers that generates a week of **skills lessons**
(speaking, listening, reading, writing) around a theme, at a given level.
The first real use is a Brighton-themed week for a 2A Elementary class
(CEFR A1+ to A2).

## How teachers actually plan (from the scheme of work)

The school's scheme of work (*2A1 SOW, May 2026*, Kings Brighton) is a Word
document with 16 weekly themes per level. Each week has a **grammar** block and
a **skills** block. A skills block holds:

- weekly aims (the same sentence with the topic swapped in);
- a bank of 8–14 tasks, each with skill tags `(S)(L)(R)(W)`, a source
  (coursebook unit and page, Supp File, Linguahouse, ESL Brains, teach-this)
  and cross-curricular tags `[PSHE] [SMSC] [B.V.] [O.S.]`;
- regular slots: weekly speaking questions, A2 extra vocabulary, study
  skills, a DVD lesson, and a writing genre (required every week);
- a Tuesday-to-Monday plan (Monday is often a test or benchmarking day);
- boxes for changes to plan, homework and evaluation, which are usually empty.

In practice teachers don't follow the bank closely. They make their own
lessons, pick what they like from what's on offer, and fill the document in
afterwards. So the tool should **produce ready-to-teach lessons**, not just
fill in the scheme of work.

Things the generator should do:

- produce complete lessons with timings, stages, worksheets and answer keys;
- write original texts and scripts where coursebooks don't cover the theme
  (no coursebook covers Brighton);
- balance the four skills across the week;
- recycle grammar the class already knows (a review when no grammar focus
  is set);
- include a daily quiz (Kahoot) and lots of structured speaking.

## Teacher preferences (Brighton week)

- Mostly speaking games and structured speaking activities.
- One listening lesson (a real online source if one exists at the level).
- One writing lesson, on **Wednesday**, with Thursday as the overflow day.
- A Kahoot every day: Brighton facts and English culture.
- Grammar: a general review (no set focus).

## Brighton week: first hand-made output

Written as a shared doc:
[Brighton Skills Week – 2A Elementary](https://claude.ai/artifact/Bureddmxi4SoapTf9wbzgf).
It's a model of what the generator should eventually produce.

| Day | Lesson 1 | Lesson 2 | Kahoot |
| --- | --- | --- | --- |
| Tue | Find someone who + Brighton true/false (S) | My first week: interview and report (S) | 1 Brighton basics |
| Wed | Writing: a postcard from Brighton (R)(W) | Postcard gallery (R)(S) | 2 English culture |
| Thu | Map info gap and directions (S)(L); overflow for writing | Tourist/local recommendation role-play (S) | 3 Places in Brighton |
| Fri | Listening: *A Saturday in Brighton* (L)(S) | Brighton bingo + Just a minute (S) | 4 Saturday in Brighton |
| Mon | "Is it true in your country?" four corners (S) | Brighton board game: grammar review (S) | 5 Big week review |

Notes:

- **Listening:** the listening sites (ESL Printables, Test-English, elllo,
  YouTube) are blocked from the cloud session, so the lesson uses an original
  A2 script. Two YouTube links from an ESL Printables Brighton worksheet are in
  the doc, not yet checked.
- **Kahoot:** quizzes can't be created in a Kahoot account from here. Each
  quiz is a table matching Kahoot's spreadsheet import template
  (question, answers 1–4, time limit, correct answer number). Import via
  *Add → Import spreadsheet* in the Kahoot editor
  ([help](https://support.kahoot.com/hc/articles/115002812547)).
- **Facts:** Palace Pier's opening date (1899) was checked. The other Brighton
  facts come from general knowledge and should be checked before class.

## App setup

- Expo SDK 58, React Native, TypeScript, Expo Router (`src/app/`), the same
  stack as the `language-learning` repo.
- Template demo screens and images removed; one blank home screen.
- ESLint (`eslint-config-expo`), with `eslint.config.js` added by hand because
  `expo lint` couldn't set itself up through the cloud network.
- Web output is a single-page app (`app.json` → `web.output: "single"`),
  published on **Vercel** (`vercel.json`). Netlify was dropped.
- Checks: `npm run lint`, `npm run typecheck`, and
  `npx expo export --platform web` all pass.

## Working environment

- Development happens in Claude Code cloud sessions (works from a phone).
  Changes are tested through Vercel preview links; Expo Go on a phone needs a
  local machine (WSL or Windows) running `npx expo start`.
- The cloud network blocks YouTube and most ESL resource sites, and Expo's
  own API (`api.expo.dev`).

## Open items

- [ ] Create the GitHub repo `gjball/efl-curriculum-generator` (empty) and
      push this project.
- [ ] Connect the repo to Vercel.
- [ ] Confirm how many skills lessons there are per day (assumed 2 × 45 min).
- [ ] Decide what the generator outputs: lesson pack doc, Word export in the
      school's scheme-of-work layout, Kahoot spreadsheets, or all three.
- [ ] Decide whether the generator uses AI (e.g. Claude) to write texts and
      quizzes, and where the API key lives.

# EFL Curriculum Generator

A tool for building courses for English learners: levels, weeks, topics,
grammar and activities.

Built with Expo (React Native + TypeScript) and Expo Router. It runs as a web
app and can be built for Android and iOS.

## Run it

```bash
npm install
npx expo start
```

Press `w` for the web version, or scan the QR code with Expo Go.

## Checks

```bash
npm run lint
npm run typecheck
```

## Where things are

- `src/app/` — the pages (Expo Router: every file is a screen).
- `src/components/`, `src/hooks/`, `src/constants/` — shared code.
- `netlify.toml` — builds and publishes the web version on Netlify.

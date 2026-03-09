# Present (Electron)

Desktop presentation tool for mixed decks of web pages and slide images.

## Run (Electron)

```bash
npm install
npm run dev
```

## Run renderer only (web preview)

```bash
npm run dev:web
```

## Build renderer

```bash
npm run build
```

## Package desktop app

```bash
npm run package:dir
```

Creates an unpacked Electron app in `release/`.

```bash
npm run package:win
```

Builds a Windows installer (`nsis`) in `release/`.

## Key behavior

- Web presentation steps open as top-level pages in a dedicated fullscreen Electron window (no iframe embedding).
- Left/Right/Space/Esc navigation is intercepted in the main process via `before-input-event`.
- Decks are autosaved to disk under Electron user data (`decks/<deckId>/deck.json` + `slides/`).
- Deck import/export uses single-file `.presentdeck` zip archives.

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

### macOS (recommended — builds both arm64 and x64)

```bash
npm run package:mac
```

Produces a `.dmg` installer and a `.zip` archive in `release/` for both Apple Silicon (`arm64`) and Intel (`x64`).

To quickly test the unpackaged app without creating an installer:

```bash
npm run package:mac-dir
```

### Windows

```bash
npm run package:win
```

Builds an NSIS installer (`release/*.exe`) for 64-bit Windows.

### Unpacked (any platform, quick test)

```bash
npm run package:dir
```

Creates an unpacked Electron app in `release/` using the current host platform.

### Both macOS and Windows

```bash
npm run package:all
```

Builds macOS (dmg + zip) and Windows (NSIS) in one go. Must be run on macOS (cross-compilation from Windows to macOS is not supported by electron-builder).

> **Note:** The built apps are fully self-contained — no `npm run dev` or Node.js installation is required to run them.

## Key behavior

- Web presentation steps open as top-level pages in a dedicated fullscreen Electron window (no iframe embedding).
- Left/Right/Space/Esc navigation is intercepted in the main process via `before-input-event`.
- Decks are autosaved to disk under Electron user data (`decks/<deckId>/deck.json` + `slides/`).
- Deck import/export uses single-file `.presentdeck` zip archives.

# AGENTS Guide

This repository is an Electron desktop presentation tool called Present. It supports mixed presentation decks made of web pages, imported slide images or videos, and parsed PowerPoint slides with animation steps.

The main goal for future agents is to make focused fixes without breaking the Electron runtime, PPTX import pipeline, or deck persistence format.

## What The App Does

- Lets users build decks containing three step types: `web`, `slide`, and `pptx-slide`.
- Stores decks under Electron user data as `decks/<deckId>/deck.json` plus media in `slides/`.
- Opens presentations in a dedicated fullscreen Electron window.
- Imports `.pptx` files by parsing OOXML and expanding slide animations into multiple presentation steps.

## How To Run

- Install + build engine: `npm install && npm run build:engine`
- Electron dev mode: `npm run dev`
- Renderer-only preview: `npm run dev:web`
- Build: `npm run build`
- Full regression tests commonly used here:
  `node --test test/pptxParser.test.cjs test/pptxPresentRenderer.test.cjs test/electronUtils.test.cjs`

## Project Layout

- `electron/`
  Main process and import/render logic.
- `src/`
  React renderer/editor UI.
- `test/`
  Node test runner regression coverage.
- `features/pptx/`
  Real PPTX fixtures used for parser/renderer regressions.
- `release/`, `dist/`, `build/`
  Generated output. Do not hand-edit unless the task explicitly requires it.

## Important Files

- `electron/main.cjs`
  Electron entrypoint, IPC handlers, presentation window lifecycle, navigation behavior.
- `electron/pptxParser.cjs`
  Thin compatibility shim — re-exports `parsePptx` from `@webpresent/pptx-engine`.
  All parsing logic lives in `packages/pptx-engine/src/parser.ts`.
- `electron/pptxPresentRenderer.cjs`
  Builds self-contained HTML for presenting PPTX slides in Electron.
- `packages/pptx-engine/`
  The `@webpresent/pptx-engine` npm workspace package. Contains all PPTX logic.
  Built by tsup to CJS (for Electron) and ESM (for Vite renderer).
  - `src/types.ts` — all PPTX data model types
  - `src/parser.ts` — OOXML parser (migrated from `electron/pptxParser.cjs`)
  - `src/renderCore.ts` — `buildSlideState()` and all rendering primitives
  - `src/animationCss.ts` — CSS @keyframes for entrance effects
  - `src/imageHelpers.ts` — image crop / flip CSS helpers
  - `src/index.ts` — public API re-exports
  - `dist/` — built CJS + ESM outputs (not hand-edited)
- `src/pptxRenderer.ts`
  Renderer-side PPTX preview HTML generator. Keep behavior aligned with the Electron presentation renderer.
- `src/types.ts`
  Shared presentation and PPTX model types.
- `src/App.tsx`
  Top-level editor shell.
- `src/components/`
  Sidebar, preview, and PPTX editing UI.
- `src/hooks/usePresentation.ts`
  Deck loading, autosave, display selection, toast state.
- `src/hooks/useSlideUrls.ts`
  Media URL resolution cache for preview rendering.
- `test/pptxParser.test.cjs`
  Regression tests for PPTX parsing behavior.
- `test/pptxPresentRenderer.test.cjs`
  Regression tests for presentation HTML behavior.

## Architecture Notes

The repo mixes module systems intentionally:

- Renderer code is TypeScript/ESM under `src/`.
- Electron code is CommonJS under `electron/*.cjs`.
- `packages/pptx-engine/` is TypeScript, built by tsup to both CJS and ESM.
  Electron `require()`s the CJS build; Vite resolves the TypeScript source
  directly via a resolve alias in `vite.config.ts` (no pre-build needed for dev).
- `package.json` uses `"type": "module"`, but Electron files remain `.cjs` and
  must stay CommonJS unless the task explicitly migrates them.

Presentation flow:

1. User edits deck in the React UI.
2. Renderer calls IPC exposed by `electron/preload.cjs`.
3. Main process in `electron/main.cjs` loads decks, imports media, parses PPTX, and controls the fullscreen presentation window.
4. PPTX steps are rendered to HTML via `electron/pptxPresentRenderer.cjs`.

PPTX import flow:

1. `electron/pptxParser.cjs` reads `.pptx` zip contents with `adm-zip`.
2. XML is parsed with `fast-xml-parser`.
3. Shapes, images, placeholders, notes, and timing data are converted into the shared slide model.
4. Media assets are extracted into the deck `slides/` folder.
5. Parsed animation groups expand into `pptx-slide` steps in the editor.

## PPTX-Specific Gotchas

These areas have already caused real regressions and should be treated carefully:

- Placeholder inheritance is split across slide, layout, and master. Do not simplify this casually.
- Layouts can suppress inherited master shapes via `showMasterSp`. Honor that or duplicate graphics can appear.
- XML draw order can cross node types such as `p:sp`, `p:pic`, `p:cxnSp`, and `p:grpSp`. If you flatten those in buckets instead of source order, labels can render underneath connectors or cards even when the text parsed correctly.
- Some text content lives in `a:fld`, not just `a:r`.
- Some layouts use repeated `a:br` runs inside a single paragraph to create intentional vertical gaps around nearby content. Compact browser paragraph defaults help tight labels, but they can also compress those spacer-heavy boxes if you do not validate them visually.
- Some shapes live inside `mc:AlternateContent`. Ignoring fallback or choice blocks will drop content.
- Image correctness depends on `a:srcRect` crop data and `flipH` / `flipV` transforms.
- Animation ordering must prefer the main sequence and ignore interactive sequences for normal build order.
- Paragraph builds are separate from whole-shape builds. A shape can appear on one click while later paragraphs reveal on later clicks.
- Editor preview and presentation renderer should stay behaviorally aligned.
- Standalone comparison artifacts and saved decks can both go stale after engine changes. Regenerate `/tmp/ready2026-slide*-render.html` outputs and reimport/save a fresh deck before trusting a visual diff.

## Current Known Issues

See `issues.md`. At the time this guide was written, listed concerns were:

- preview scaling should be fixed to slide dimensions
- some text boxes show unwanted black borders
- moving slides is slow
- moving slides forward has delay

Do not assume these are fully resolved unless you verify them in the live app.

## Recent Regression Areas

Recent fixes added regression coverage for:

- zero-height connector SVG boxes and marker clipping
- compact default paragraph spacing for tight labels
- ordered inline `a:r` / `a:fld` / `a:br` parsing
- gradient-only borders on standard rectangle shapes
- text-box overflow fitting in fullscreen and preview HTML
- mixed node-type draw-order preservation
- slide-number field extraction
- cropped image rendering
- paragraph build ordering
- layout master-shape suppression
- alternate-content fallback shape preservation
- cached PPTX media reuse
- in-place PPTX runtime updates between builds

If you touch any of these areas, re-run the PPTX tests.

## Verification Expectations

Before claiming a fix is done, prefer this verification set:

1. `node --test test/pptxParser.test.cjs test/pptxPresentRenderer.test.cjs test/electronUtils.test.cjs`
2. `npm run build`

If the task affects live presentation behavior, automated tests are not enough on their own. A manual Electron smoke test is the best follow-up when feasible.

For slide-by-slide PDF comparison work:

1. Rebuild the engine after parser or render-core edits.
2. Regenerate any standalone `/tmp/ready2026-slide*-render.html` artifacts before comparing, because those files are only valid for the engine build that produced them.
3. Compare against the correct animation state. For READY2026 PDF pages this is often the final visible build, not step `0`.
4. If a screenshot looks faded or washed out right after reload, recapture after the slide settles before diagnosing opacity or colour bugs.

## Editing Guidance For Future Agents

- Prefer minimal, surgical changes.
- Avoid touching generated files in `dist/`, `release/`, or `build/` unless explicitly requested.
- If you update parser behavior, check whether `src/pptxRenderer.ts` and `electron/pptxPresentRenderer.cjs` must stay in sync.
- If you add a bug fix, add or update a regression test first whenever practical.
- Be careful with XML array handling. `fast-xml-parser` shape/timing structures are sensitive to `isArray` behavior.
- Keep ordered-text parsing and ordered-draw flattening separate in the parser. Both matter, and fixing one does not automatically fix the other.

## Useful Fixtures

These PPTX files are especially useful for regression work:

- `features/pptx/MMC_simpliPyTEM_presentation.pptx`
  Good for animation/build-order and banner-layout behavior.
- `features/pptx/CQC - DV Portfolio_GI.pptx`
  Good for image crop behavior and footer/slide-number content.
- `features/pptx/1-CHEMA1_intro_regression.pptx`
  Good for placeholder inheritance and alternate-content text-box behavior.
- `features/pptx/GabrielIng_slideForGRC.pptx`
  Good for click-build timing checks.

## Success Criteria For Most Changes

A safe change here usually means:

- the targeted tests pass
- the project still builds
- deck import/export format is unchanged unless intentionally migrated
- Electron presentation navigation still works
- PPTX preview and presentation output remain visually consistent

## Git Commit Policy

- **Never add Copilot as a commit author or co-author.** Do not include `Co-authored-by: GitHub Copilot` or any variant containing "copilot" in commit messages or trailers.
- Commits must be authored solely by the human developer.
- When suggesting or generating commit messages, omit all AI attribution lines.

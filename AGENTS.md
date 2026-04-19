# AGENTS Guide

This repository is an Electron desktop presentation tool called Present. It supports mixed presentation decks made of web pages, imported slide images or videos, and parsed PowerPoint slides with animation steps.

The main goal for future agents is to make focused fixes without breaking the Electron runtime, PPTX import pipeline, or deck persistence format.

## What The App Does

- Lets users build decks containing three step types: `web`, `slide`, and `pptx-slide`.
- Stores decks under Electron user data as `decks/<deckId>/deck.json` plus media in `slides/`.
- Opens presentations in a dedicated fullscreen Electron window.
- Imports `.pptx` files by parsing OOXML and expanding slide animations into multiple presentation steps.

## How To Run

- Install: `npm install`
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
  Core PPTX OOXML parser. This is the highest-risk file for rendering regressions.
- `electron/pptxPresentRenderer.cjs`
  Builds self-contained HTML for presenting PPTX slides in Electron.
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
- `package.json` uses `"type": "module"`, but Electron files remain `.cjs` and must stay CommonJS unless the task explicitly migrates them.

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
- Some text content lives in `a:fld`, not just `a:r`.
- Some shapes live inside `mc:AlternateContent`. Ignoring fallback or choice blocks will drop content.
- Image correctness depends on `a:srcRect` crop data and `flipH` / `flipV` transforms.
- Animation ordering must prefer the main sequence and ignore interactive sequences for normal build order.
- Paragraph builds are separate from whole-shape builds. A shape can appear on one click while later paragraphs reveal on later clicks.
- Editor preview and presentation renderer should stay behaviorally aligned.

## Current Known Issues

See `issues.md`. At the time this guide was written, listed concerns were:

- preview scaling should be fixed to slide dimensions
- some text boxes show unwanted black borders
- moving slides is slow
- moving slides forward has delay

Do not assume these are fully resolved unless you verify them in the live app.

## Recent Regression Areas

Recent fixes added regression coverage for:

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

## Editing Guidance For Future Agents

- Prefer minimal, surgical changes.
- Avoid touching generated files in `dist/`, `release/`, or `build/` unless explicitly requested.
- If you update parser behavior, check whether `src/pptxRenderer.ts` and `electron/pptxPresentRenderer.cjs` must stay in sync.
- If you add a bug fix, add or update a regression test first whenever practical.
- Be careful with XML array handling. `fast-xml-parser` shape/timing structures are sensitive to `isArray` behavior.

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
- PPTX preview and presentation output remain visually consistent# AGENTS Guide

This repository is an Electron desktop presentation tool called Present. It supports mixed presentation decks made of web pages, imported slide images or videos, and parsed PowerPoint slides with animation steps.

The main goal for future agents is to make focused fixes without breaking the Electron runtime, PPTX import pipeline, or deck persistence format.

## What The App Does

- Lets users build decks containing three step types: `web`, `slide`, and `pptx-slide`.
- Stores decks under Electron user data as `decks/<deckId>/deck.json` plus media in `slides/`.
- Opens presentations in a dedicated fullscreen Electron window.
- Imports `.pptx` files by parsing OOXML and expanding slide animations into multiple presentation steps.

## How To Run

- Install: `npm install`
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
  Core PPTX OOXML parser. This is the highest-risk file for rendering regressions.
- `electron/pptxPresentRenderer.cjs`
  Builds self-contained HTML for presenting PPTX slides in Electron.
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
- `package.json` uses `"type": "module"`, but Electron files remain `.cjs` and must stay CommonJS unless the task explicitly migrates them.

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
- Some text content lives in `a:fld`, not just `a:r`.
- Some shapes live inside `mc:AlternateContent`. Ignoring fallback or choice blocks will drop content.
- Image correctness depends on `a:srcRect` crop data and `flipH` / `flipV` transforms.
- Animation ordering must prefer the main sequence and ignore interactive sequences for normal build order.
- Paragraph builds are separate from whole-shape builds. A shape can appear on one click while later paragraphs reveal on later clicks.
- Editor preview and presentation renderer should stay behaviorally aligned.

## Current Known Issues

See `issues.md`. At the time this guide was written, listed concerns were:

- preview scaling should be fixed to slide dimensions
- some text boxes show unwanted black borders
- moving slides is slow
- moving slides forward has delay

Do not assume these are fully resolved unless you verify them in the live app.

## Recent Regression Areas

Recent fixes added regression coverage for:

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

## Editing Guidance For Future Agents

- Prefer minimal, surgical changes.
- Avoid touching generated files in `dist/`, `release/`, or `build/` unless explicitly requested.
- If you update parser behavior, check whether `src/pptxRenderer.ts` and `electron/pptxPresentRenderer.cjs` must stay in sync.
- If you add a bug fix, add or update a regression test first whenever practical.
- Be careful with XML array handling. `fast-xml-parser` shape/timing structures are sensitive to `isArray` behavior.

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
# PPTX Engine Package Design

**Date:** 2026-05-01  
**Branch:** ppt_import  
**Status:** Approved

---

## Problem

PPTX rendering logic is fully duplicated between two files:

- `electron/pptxPresentRenderer.cjs` — presentation-time renderer (Electron CJS)
- `src/pptxRenderer.ts` — editor preview renderer (TypeScript ESM)

These share identical implementations of `fillToCss`, `renderTextRun`, `renderParagraph`, `renderShape`, `getCroppedImageStyles`, and CSS animation keyframes. Any fix or animation improvement must be applied twice. The PPTX parser lives in `electron/pptxParser.cjs` (1702 lines of CJS) while types live in `src/types.ts`, further scattering PPTX concerns.

---

## Goal

Create a single `@webpresent/pptx-engine` npm workspace package that owns all PPTX logic: types, rendering core, animation CSS, image helpers, and eventually the parser. Both the Electron main process and the React renderer import from this package.

---

## Architecture

### Package location

```
packages/
  pptx-engine/
    src/
      types.ts          PPTX data model types (moved from src/types.ts)
      animationCss.ts   CSS @keyframes string for entrance effects
      imageHelpers.ts   clampCrop + getCroppedImageStyles
      renderCore.ts     escHtml, fillToCss, renderTextRun, renderParagraph,
                        renderShape, buildSlideState → SlideState
      parser.ts         parsePptx (migrated from electron/pptxParser.cjs)
      index.ts          public re-exports
    tsconfig.json       standalone tsconfig for tsup
    tsup.config.ts      builds CJS + ESM + .d.ts
    package.json        name: "@webpresent/pptx-engine", private: true
```

### Build outputs

`tsup` produces:
- `dist/index.cjs` — CommonJS for `electron/` files
- `dist/index.js` — ESM for Vite renderer
- `dist/index.d.ts` — TypeScript declarations

### Module resolution

| Consumer | Mechanism | Resolves to |
|----------|-----------|-------------|
| `electron/*.cjs` | `require('@webpresent/pptx-engine')` | `dist/index.cjs` (CJS build) |
| `src/*.ts` in dev | Vite alias | `packages/pptx-engine/src/index.ts` (TS source direct) |
| `src/*.ts` in build | `tsc -b` + tsconfig paths | `packages/pptx-engine/src/index.ts` |
| `vite build` | Vite alias | `packages/pptx-engine/src/index.ts` (TS source, Vite transpiles) |

The Vite alias ensures the renderer never needs a pre-built package during dev or `vite build`. The CJS build is required for Electron.

### Public API (Phase 1 scope)

```ts
// Types
export type { PptxShapeType, PptxTextRun, PptxParagraph, PptxImageCrop,
               PptxFill, PptxBorder, PptxShape, PptxSlideData, PptxDeckData } from './types';

// Animation
export { ANIMATION_KEYFRAMES } from './animationCss';

// Image helpers
export { clampCrop, getCroppedImageStyles } from './imageHelpers';

// Rendering
export type { SlideState } from './renderCore';
export { escHtml, fillToCss, renderTextRun, renderParagraph,
         renderShape, buildSlideState } from './renderCore';

// Parser (Phase 2, included in this plan)
export { parsePptx } from './parser';
```

`buildSlideState(slide, animationStep, mediaResolver)` returns `SlideState`:
```ts
type SlideState = { width: number; height: number; backgroundCss: string; shapesHtml: string }
```

Both renderer files become thin wrappers: `pptxPresentRenderer.cjs` adds the `__WEBPRESENT_UPDATE_PPTX` shell and media I/O cache; `src/pptxRenderer.ts` adds the static HTML wrapper and fit() script.

---

## What Changes

| File | Before | After |
|------|--------|-------|
| `src/types.ts` | Owns PPTX types + presentation types | Re-exports PPTX types from package; owns presentation types |
| `src/pptxRenderer.ts` | 318 lines, full rendering impl | ~80 lines, wraps `buildSlideState` |
| `electron/pptxPresentRenderer.cjs` | 256 lines, full rendering impl | ~130 lines, wraps `buildSlideState` |
| `electron/pptxParser.cjs` | 1702 lines CJS | Replaced by `packages/pptx-engine/src/parser.ts` + thin shim |
| `packages/pptx-engine/` | Does not exist | ~500 lines TS across focused modules |

---

## What Does Not Change

- `electron/main.cjs` import names are unchanged (same function names)
- `electron/pptxPresentRenderer.cjs` exported function names are unchanged
- The `__WEBPRESENT_UPDATE_PPTX` in-place update mechanism is unchanged
- `deck.json` format is unchanged
- All test files continue to pass

---

## Electron Packaging

`electron-builder` must include the package dist. Add `packages/pptx-engine/dist/**/*` to the `files` array in `package.json`. With npm workspaces, `node_modules/@webpresent/pptx-engine` is a symlink to `packages/pptx-engine`; including the dist path directly ensures ASAR packaging works regardless of symlink handling.

---

## Future Work (Not In This Plan)

- Full animation CSS effects library (Phase 2 — animation work)
- `parsePptx` typed parameter/return annotations (TypeScript gradual enhancement)
- Animation delay elimination (separate issue)

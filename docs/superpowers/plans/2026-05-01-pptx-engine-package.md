# PPTX Engine Package Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extract all PPTX logic into a `@webpresent/pptx-engine` npm workspace package (TypeScript, dual CJS+ESM), eliminating the rendering duplication between `electron/pptxPresentRenderer.cjs` and `src/pptxRenderer.ts` and consolidating the parser.

**Architecture:** A private `packages/pptx-engine/` workspace package built by tsup (CJS + ESM + .d.ts). The Electron main process `require()`s the CJS build; the Vite renderer resolves the TypeScript source directly via a Vite alias. Both existing renderer files become thin wrappers around the shared `buildSlideState` core.

**Tech Stack:** TypeScript 5, tsup 8, npm workspaces, existing adm-zip + fast-xml-parser deps

**Spec:** `docs/superpowers/specs/2026-05-01-pptx-engine-package-design.md`

---

## Task 1: Scaffold workspace package

**Files:**
- Create: `packages/pptx-engine/package.json`
- Create: `packages/pptx-engine/tsconfig.json`
- Create: `packages/pptx-engine/tsup.config.ts`
- Create: `packages/pptx-engine/src/index.ts` (stub)
- Modify: `package.json` (root)

- [ ] **Step 1: Create `packages/pptx-engine/package.json`**

```json
{
  "name": "@webpresent/pptx-engine",
  "version": "0.1.0",
  "private": true,
  "description": "PPTX parsing and rendering engine for WebPresent",
  "type": "module",
  "main": "./dist/index.cjs",
  "module": "./dist/index.js",
  "types": "./dist/index.d.ts",
  "exports": {
    ".": {
      "require": "./dist/index.cjs",
      "import": "./dist/index.js",
      "types": "./dist/index.d.ts"
    }
  },
  "scripts": {
    "build": "tsup",
    "dev": "tsup --watch"
  },
  "dependencies": {
    "adm-zip": "^0.5.16",
    "fast-xml-parser": "^5.5.9"
  },
  "devDependencies": {
    "tsup": "^8.4.0",
    "typescript": "^5.6.3"
  }
}
```

- [ ] **Step 2: Create `packages/pptx-engine/tsconfig.json`**

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ESNext",
    "moduleResolution": "Bundler",
    "strict": true,
    "declaration": true,
    "composite": false,
    "skipLibCheck": true,
    "outDir": "./dist",
    "rootDir": "./src"
  },
  "include": ["src", "tsup.config.ts"]
}
```

- [ ] **Step 3: Create `packages/pptx-engine/tsup.config.ts`**

```ts
import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['cjs', 'esm'],
  dts: true,
  clean: true,
  sourcemap: true,
});
```

- [ ] **Step 4: Create stub `packages/pptx-engine/src/index.ts`**

```ts
// Public API — populated task by task
export const PPTX_ENGINE_VERSION = '0.1.0';
```

- [ ] **Step 5: Add workspaces to root `package.json`**

Add `"workspaces": ["packages/*"]` after `"type": "module"`. Update the `"build"` and `"dev"` scripts to pre-build the engine package:

```json
"workspaces": ["packages/*"],
"scripts": {
  "dev": "npm run build:engine && concurrently -k \"vite\" \"wait-on tcp:5173 && cross-env ELECTRON_RENDERER_URL=http://localhost:5173 electron ./electron/main.cjs\"",
  "dev:web": "vite",
  "start": "electron ./electron/main.cjs",
  "build:engine": "npm run build --workspace=packages/pptx-engine",
  "build": "npm run build:engine && tsc -b && vite build",
  "preview": "vite preview",
  "package:dir": "npm run build && cross-env CSC_IDENTITY_AUTO_DISCOVERY=false electron-builder --dir",
  "package:mac": "npm run build && cross-env CSC_IDENTITY_AUTO_DISCOVERY=false electron-builder --mac",
  "package:mac-dir": "npm run build && cross-env CSC_IDENTITY_AUTO_DISCOVERY=false electron-builder --mac --dir",
  "package:win": "npm run build && cross-env CSC_IDENTITY_AUTO_DISCOVERY=false electron-builder --win nsis",
  "package:all": "npm run build && cross-env CSC_IDENTITY_AUTO_DISCOVERY=false electron-builder --mac --win"
},
```

Also update the electron-builder `"files"` array to include the package dist:

```json
"files": [
  "dist/**/*",
  "electron/**/*",
  "packages/pptx-engine/dist/**/*",
  "node_modules/**/*",
  "package.json"
],
```

- [ ] **Step 6: Add Vite alias for renderer dev**

Replace the entire content of `vite.config.ts` with:

```ts
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@webpresent/pptx-engine': path.resolve(
        __dirname,
        'packages/pptx-engine/src/index.ts',
      ),
    },
  },
});
```

- [ ] **Step 7: Add TypeScript path mapping for `tsc -b`**

In `tsconfig.app.json`, add `"baseUrl": "."` and a `"paths"` entry so `tsc -b` can resolve the package source during type-checking (without needing a pre-built dist):

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "useDefineForClassFields": true,
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "module": "ESNext",
    "skipLibCheck": true,
    "moduleResolution": "Bundler",
    "allowImportingTsExtensions": false,
    "resolveJsonModule": true,
    "isolatedModules": true,
    "noEmit": true,
    "jsx": "react-jsx",
    "strict": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true,
    "baseUrl": ".",
    "paths": {
      "@webpresent/pptx-engine": ["packages/pptx-engine/src/index.ts"]
    }
  },
  "include": ["src"]
}
```

- [ ] **Step 8: Install and verify scaffold builds**

```bash
cd /Users/gabriel/Developer/WebPresent
npm install
npm run build:engine
```

Expected output: `packages/pptx-engine/dist/index.cjs`, `dist/index.js`, `dist/index.d.ts` created. No errors.

- [ ] **Step 9: Commit scaffold**

```bash
git add packages/ package.json vite.config.ts tsconfig.app.json
git commit -m "feat: scaffold @webpresent/pptx-engine workspace package"
```

---

## Task 2: Extract PPTX types

**Files:**
- Create: `packages/pptx-engine/src/types.ts`
- Modify: `packages/pptx-engine/src/index.ts`
- Modify: `src/types.ts`

- [ ] **Step 1: Create `packages/pptx-engine/src/types.ts`**

Copy only the PPTX type block out of `src/types.ts`. The file should contain exactly these types:

```ts
// ── PPTX Slide Types ─────────────────────────────────────────────────────────

export type PptxShapeType = 'rect' | 'ellipse' | 'roundRect' | 'image' | 'line' | 'freeform';

export type PptxTextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  colour?: string;
};

export type PptxParagraph = {
  runs: PptxTextRun[];
  alignment?: 'left' | 'center' | 'right' | 'justify';
  bulletType?: 'none' | 'bullet' | 'numbered';
  bulletChar?: string;
  level?: number;
  animationGroup?: number;
};

export type PptxImageCrop = {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
};

export type PptxFill = {
  type: 'solid' | 'image' | 'gradient' | 'none';
  colour?: string;
  imageRelativePath?: string;
  gradientStops?: { position: number; colour: string }[];
};

export type PptxBorder = {
  width: number;
  colour: string;
  style?: 'solid' | 'dashed' | 'dotted';
};

export type PptxShape = {
  id: string;
  name?: string;
  type: PptxShapeType;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  fill?: PptxFill;
  border?: PptxBorder;
  paragraphs?: PptxParagraph[];
  imageRelativePath?: string;
  imageCrop?: PptxImageCrop;
  cornerRadius?: number;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  flipH?: boolean;
  flipV?: boolean;
  animationGroup: number;
  animationEffect?: 'appear' | 'fade' | 'fly-left' | 'fly-right' | 'fly-up' | 'fly-down';
};

export type PptxSlideData = {
  slideIndex: number;
  width: number;
  height: number;
  background?: PptxFill;
  shapes: PptxShape[];
  notes?: string;
  animationStepCount: number;
};

export type PptxDeckData = {
  sourceFileName: string;
  slides: PptxSlideData[];
  theme?: {
    colours: Record<string, string>;
    defaultFont?: string;
  };
};
```

- [ ] **Step 2: Export types from `packages/pptx-engine/src/index.ts`**

Replace the stub with:

```ts
export type {
  PptxShapeType,
  PptxTextRun,
  PptxParagraph,
  PptxImageCrop,
  PptxFill,
  PptxBorder,
  PptxShape,
  PptxSlideData,
  PptxDeckData,
} from './types';
```

- [ ] **Step 3: Update `src/types.ts` to re-export PPTX types from package**

Replace the entire PPTX types block (lines 12–100) with a re-export. The file should read:

```ts
export type {
  PptxShapeType,
  PptxTextRun,
  PptxParagraph,
  PptxImageCrop,
  PptxFill,
  PptxBorder,
  PptxShape,
  PptxSlideData,
  PptxDeckData,
} from '@webpresent/pptx-engine';

export type StepType = 'web' | 'slide' | 'pptx-slide';

export type SlideMediaKind = 'image' | 'video';

export type SlideRef = {
  id: string;
  relativePath: string;
  sourceFileName?: string;
  mediaKind?: SlideMediaKind;
};

// ── Presentation Types ───────────────────────────────────────────────────────

export type PresentationStep = {
  id: string;
  type: StepType;
  title?: string;
  notes?: string;
  groupId?: string;
  url?: string;
  webZoom?: number;
  slideRef?: SlideRef;
  pptxSlideData?: PptxSlideData;
  pptxAnimationStep?: number;
};

export type Presentation = {
  id: string;
  title: string;
  createdAt: string;
  updatedAt: string;
  items: PresentationStep[];
};

export type SlideImportMode = 'separate' | 'grouped';

export type DisplayInfo = {
  id: number;
  label: string;
  width: number;
  height: number;
};
```

Note: `PptxSlideData` is still usable in `PresentationStep` because it is re-exported from this file.

- [ ] **Step 4: Build engine and verify type check passes**

```bash
cd /Users/gabriel/Developer/WebPresent
npm run build:engine
npx tsc -b --noEmit
```

Expected: no errors.

- [ ] **Step 5: Commit**

```bash
git add packages/pptx-engine/src/types.ts packages/pptx-engine/src/index.ts src/types.ts tsconfig.app.json
git commit -m "feat(pptx-engine): extract PPTX types into package"
```

---

## Task 3: Extract animation CSS and image helpers

**Files:**
- Create: `packages/pptx-engine/src/animationCss.ts`
- Create: `packages/pptx-engine/src/imageHelpers.ts`
- Modify: `packages/pptx-engine/src/index.ts`

- [ ] **Step 1: Create `packages/pptx-engine/src/animationCss.ts`**

```ts
/**
 * CSS @keyframe declarations for PPTX entrance animation effects.
 *
 * These are embedded in every rendered slide HTML document.
 * Extend this list as additional effect types are supported.
 */
export const ANIMATION_KEYFRAMES = `
@keyframes pptx-appear {
  from { opacity: 0; }
  to   { opacity: 1; }
}
@keyframes pptx-fade {
  from { opacity: 0; }
  to   { opacity: 1; }
}
@keyframes pptx-fly-left {
  from { opacity: 0; transform: translateX(-60px); }
  to   { opacity: 1; transform: translateX(0); }
}
@keyframes pptx-fly-right {
  from { opacity: 0; transform: translateX(60px); }
  to   { opacity: 1; transform: translateX(0); }
}
@keyframes pptx-fly-up {
  from { opacity: 0; transform: translateY(-60px); }
  to   { opacity: 1; transform: translateY(0); }
}
@keyframes pptx-fly-down {
  from { opacity: 0; transform: translateY(60px); }
  to   { opacity: 1; transform: translateY(0); }
}
`;
```

- [ ] **Step 2: Create `packages/pptx-engine/src/imageHelpers.ts`**

```ts
import type { PptxShape } from './types';

/**
 * Clamps a fractional crop value to the valid range [0, 0.9999].
 * Values outside this range would produce invisible or inverted images.
 */
export function clampCrop(value: number | undefined): number {
  const n = Number(value) || 0;
  return Math.min(0.9999, Math.max(0, n));
}

/**
 * Produces inline CSS for an image that needs crop and/or flip transforms.
 *
 * The parent container uses `overflow:hidden` and the image is sized/
 * positioned to show only the cropped region. Flip transforms are applied
 * via CSS `scaleX(-1)` / `scaleY(-1)`.
 */
export function getCroppedImageStyles(shape: PptxShape): string {
  const leftCrop = clampCrop(shape.imageCrop?.left);
  const topCrop = clampCrop(shape.imageCrop?.top);
  const rightCrop = clampCrop(shape.imageCrop?.right);
  const bottomCrop = clampCrop(shape.imageCrop?.bottom);
  const visibleWidth = Math.max(0.0001, 1 - leftCrop - rightCrop);
  const visibleHeight = Math.max(0.0001, 1 - topCrop - bottomCrop);

  const transforms: string[] = [];
  if (shape.flipH) transforms.push('scaleX(-1)');
  if (shape.flipV) transforms.push('scaleY(-1)');

  const styles: string[] = [
    'position:absolute',
    `left:${(-leftCrop / visibleWidth) * 100}%`,
    `top:${(-topCrop / visibleHeight) * 100}%`,
    `width:${(1 / visibleWidth) * 100}%`,
    `height:${(1 / visibleHeight) * 100}%`,
    'display:block',
  ];

  if (transforms.length) {
    styles.push(`transform:${transforms.join(' ')}`);
    styles.push('transform-origin:center center');
  }

  return styles.join(';');
}
```

- [ ] **Step 3: Export from `packages/pptx-engine/src/index.ts`**

Append to the existing exports:

```ts
export type {
  PptxShapeType,
  PptxTextRun,
  PptxParagraph,
  PptxImageCrop,
  PptxFill,
  PptxBorder,
  PptxShape,
  PptxSlideData,
  PptxDeckData,
} from './types';

export { ANIMATION_KEYFRAMES } from './animationCss';
export { clampCrop, getCroppedImageStyles } from './imageHelpers';
```

- [ ] **Step 4: Build and verify**

```bash
cd /Users/gabriel/Developer/WebPresent
npm run build:engine
```

Expected: no errors.

- [ ] **Step 5: Commit**

```bash
git add packages/pptx-engine/src/
git commit -m "feat(pptx-engine): add animationCss and imageHelpers modules"
```

---

## Task 4: Extract shared rendering core

**Files:**
- Create: `packages/pptx-engine/src/renderCore.ts`
- Modify: `packages/pptx-engine/src/index.ts`

- [ ] **Step 1: Create `packages/pptx-engine/src/renderCore.ts`**

```ts
import type {
  PptxSlideData,
  PptxShape,
  PptxFill,
  PptxParagraph,
  PptxTextRun,
} from './types';
import { getCroppedImageStyles } from './imageHelpers';

/**
 * The minimal serialisable state needed to render or update a slide.
 * Passed as JSON to `window.__WEBPRESENT_UPDATE_PPTX` for in-place updates.
 */
export type SlideState = {
  width: number;
  height: number;
  backgroundCss: string;
  shapesHtml: string;
};

// ── HTML escaping ─────────────────────────────────────────────────────────────

/** Escapes a string for safe inline use in HTML text content and attribute values. */
export function escHtml(s: string): string {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Fill → CSS ────────────────────────────────────────────────────────────────

/**
 * Converts a PptxFill to a CSS background value.
 *
 * @param fill - The fill descriptor from the parsed shape or slide background.
 * @param mediaResolver - Called with a relative media path; returns a data URL
 *   or absolute URL. Defaults to returning an empty string (no media).
 */
export function fillToCss(
  fill: PptxFill | undefined,
  mediaResolver: (path: string) => string = () => '',
): string {
  if (!fill || fill.type === 'none') return 'transparent';
  if (fill.type === 'solid' && fill.colour) return fill.colour;
  if (fill.type === 'image' && fill.imageRelativePath) {
    const url = mediaResolver(fill.imageRelativePath);
    return url ? `url("${escHtml(url)}") center/cover no-repeat` : '#cccccc';
  }
  if (fill.type === 'gradient' && fill.gradientStops?.length) {
    const stops = fill.gradientStops
      .map((s) => `${s.colour} ${Math.round(s.position)}%`)
      .join(', ');
    return `linear-gradient(180deg, ${stops})`;
  }
  return 'transparent';
}

// ── Text run → HTML ───────────────────────────────────────────────────────────

/** Renders a single styled text run to an HTML `<span>`. */
export function renderTextRun(run: PptxTextRun): string {
  if (run.text === '\n') return '<br/>';
  const styles: string[] = [];
  if (run.bold) styles.push('font-weight:bold');
  if (run.italic) styles.push('font-style:italic');
  if (run.underline) styles.push('text-decoration:underline');
  if (run.fontSize) styles.push(`font-size:${run.fontSize}pt`);
  if (run.fontFamily) styles.push(`font-family:${escHtml(run.fontFamily)},sans-serif`);
  if (run.colour) styles.push(`color:${run.colour}`);
  const text = escHtml(run.text);
  return styles.length
    ? `<span style="${styles.join(';')}">${text}</span>`
    : `<span>${text}</span>`;
}

// ── Paragraph → HTML ──────────────────────────────────────────────────────────

/** Renders a single paragraph (already filtered for visibility) to a `<p>`. */
export function renderParagraph(para: PptxParagraph): string {
  const styles = ['margin:0', 'padding:0 0 0.15em 0'];
  if (para.alignment) styles.push(`text-align:${para.alignment}`);
  if (para.level && para.level > 0) styles.push(`padding-left:${para.level * 1.5}em`);
  const runs = para.runs.map(renderTextRun).join('');
  const bullet =
    para.bulletType === 'bullet'
      ? `<span style="margin-right:0.4em">${escHtml(para.bulletChar || '•')}</span>`
      : '';
  return `<p style="${styles.join(';')}">${bullet}${runs}</p>`;
}

// ── Shape → HTML element ──────────────────────────────────────────────────────

/**
 * Renders a single shape to an HTML string.
 *
 * Returns an empty string for shapes whose `animationGroup` is greater than
 * `animationStep` (not yet visible at the current build step).
 *
 * @param shape - Parsed shape from PptxSlideData.
 * @param animationStep - Current build step index (0 = all non-animated shapes visible).
 * @param mediaResolver - Resolves a relative media path to a URL.
 */
export function renderShape(
  shape: PptxShape,
  animationStep: number,
  mediaResolver: (path: string) => string = () => '',
): string {
  if (shape.animationGroup > animationStep) return '';

  const isNewlyRevealed = shape.animationGroup === animationStep && animationStep > 0;
  const visibleParagraphs = (shape.paragraphs || []).filter(
    (p) => (p.animationGroup || 0) <= animationStep,
  );

  const styles: string[] = [
    'position:absolute',
    'box-sizing:border-box',
    `left:${shape.x}px`,
    `top:${shape.y}px`,
    `width:${shape.width}px`,
    `height:${shape.height}px`,
    'overflow:hidden',
    'word-wrap:break-word',
  ];

  if (shape.rotation) styles.push(`transform:rotate(${shape.rotation}deg)`);

  if (shape.type !== 'image' && shape.type !== 'line') {
    const bg = fillToCss(shape.fill, mediaResolver);
    if (bg !== 'transparent') styles.push(`background:${bg}`);
  }

  if (shape.border) {
    styles.push(
      `border:${shape.border.width}pt ${shape.border.style || 'solid'} ${shape.border.colour}`,
    );
  }

  if (shape.type === 'roundRect' && shape.cornerRadius) {
    styles.push(`border-radius:${shape.cornerRadius}px`);
  }
  if (shape.type === 'ellipse') styles.push('border-radius:50%');

  if (isNewlyRevealed) {
    const effect = shape.animationEffect || 'appear';
    styles.push(`animation:pptx-${effect} 0.5s ease both`);
  }

  if (visibleParagraphs.length) {
    styles.push('display:flex', 'flex-direction:column');
    const vAlign = shape.verticalAlign || 'top';
    if (vAlign === 'middle') styles.push('justify-content:center');
    else if (vAlign === 'bottom') styles.push('justify-content:flex-end');
    else styles.push('justify-content:flex-start');
    styles.push('padding:4px 8px');
  }

  if (shape.type === 'image' && shape.imageRelativePath) {
    const url = mediaResolver(shape.imageRelativePath);
    const imageStyles =
      shape.imageCrop || shape.flipH || shape.flipV
        ? getCroppedImageStyles(shape)
        : 'width:100%;height:100%;object-fit:contain;display:block';
    return `<div style="${styles.join(';')}"><img src="${escHtml(url)}" alt="" style="${imageStyles}"/></div>`;
  }

  if (shape.type === 'line') {
    const colour = shape.border?.colour || shape.fill?.colour || '#000000';
    const weight = shape.border?.width || 1;
    styles.push(`border-top:${weight}pt solid ${colour}`, 'height:0');
    return `<div style="${styles.join(';')}"></div>`;
  }

  const inner = visibleParagraphs.map(renderParagraph).join('');
  return `<div style="${styles.join(';')}">${inner}</div>`;
}

// ── Slide → SlideState ────────────────────────────────────────────────────────

/**
 * Converts a PptxSlideData snapshot at a given animation step into a
 * serialisable `SlideState` object containing all pre-rendered HTML/CSS.
 *
 * This is the shared core used by both the Electron presentation renderer
 * (which serialises state as JSON for `__WEBPRESENT_UPDATE_PPTX`) and the
 * React editor preview renderer (which wraps it in a static HTML document).
 *
 * @param slide - Parsed slide data from PptxDeckData.
 * @param animationStep - Which animation build step to render (0-based).
 * @param mediaResolver - Resolves a relative media path to a data URL or URL.
 */
export function buildSlideState(
  slide: PptxSlideData,
  animationStep: number,
  mediaResolver: (path: string) => string = () => '',
): SlideState {
  return {
    width: slide.width || 960,
    height: slide.height || 540,
    backgroundCss: slide.background
      ? fillToCss(slide.background, mediaResolver)
      : '#ffffff',
    shapesHtml: (slide.shapes || [])
      .map((s) => renderShape(s, animationStep, mediaResolver))
      .join('\n'),
  };
}
```

- [ ] **Step 2: Export renderCore from `packages/pptx-engine/src/index.ts`**

Replace the full index.ts with:

```ts
export type {
  PptxShapeType,
  PptxTextRun,
  PptxParagraph,
  PptxImageCrop,
  PptxFill,
  PptxBorder,
  PptxShape,
  PptxSlideData,
  PptxDeckData,
} from './types';

export { ANIMATION_KEYFRAMES } from './animationCss';
export { clampCrop, getCroppedImageStyles } from './imageHelpers';
export type { SlideState } from './renderCore';
export {
  escHtml,
  fillToCss,
  renderTextRun,
  renderParagraph,
  renderShape,
  buildSlideState,
} from './renderCore';
```

- [ ] **Step 3: Build engine and verify**

```bash
cd /Users/gabriel/Developer/WebPresent
npm run build:engine
```

Expected: no errors. Check `packages/pptx-engine/dist/index.d.ts` contains `buildSlideState`, `SlideState`, `ANIMATION_KEYFRAMES`.

- [ ] **Step 4: Commit**

```bash
git add packages/pptx-engine/src/
git commit -m "feat(pptx-engine): add renderCore with buildSlideState"
```

---

## Task 5: Update `src/pptxRenderer.ts`

**Files:**
- Modify: `src/pptxRenderer.ts`

- [ ] **Step 1: Replace `src/pptxRenderer.ts` with package-backed implementation**

The entire file becomes:

```ts
import { buildSlideState, ANIMATION_KEYFRAMES } from '@webpresent/pptx-engine';
import type { PptxSlideData } from '@webpresent/pptx-engine';

// ── Slide fit() script ────────────────────────────────────────────────────────

function makeFitScript(width: number, height: number): string {
  return `(function(){
  function fit(){
    var root=document.querySelector('.slide-root');
    if(!root)return;
    var sw=${width},sh=${height};
    var vw=window.innerWidth||document.documentElement.clientWidth||${width};
    var vh=window.innerHeight||document.documentElement.clientHeight||${height};
    if(vw<1||vh<1)return;
    var scale=Math.min(vw/sw,vh/sh);
    root.style.transform='scale('+scale+')';
    root.style.left=Math.round((vw-sw*scale)/2)+'px';
    root.style.top=Math.round((vh-sh*scale)/2)+'px';
  }
  window.addEventListener('resize',fit);
  window.addEventListener('load',fit);
  setTimeout(fit,0);
  setTimeout(fit,50);
  setTimeout(fit,200);
  if(window.ResizeObserver){new ResizeObserver(fit).observe(document.documentElement);}
})();`;
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Renders a PptxSlideData snapshot to a self-contained HTML document string.
 *
 * Used for the editor preview pane (via iframe srcdoc). The output includes
 * inline CSS, animation keyframes, and a fit() script that scales the slide
 * to fill the viewport.
 *
 * @param slide - Parsed slide data.
 * @param animationStep - Build step index to render (0 = base state).
 * @param mediaResolver - Maps a relative media path to a data URL or URL.
 */
export function renderPptxSlideToHtml(
  slide: PptxSlideData,
  animationStep: number,
  mediaResolver: (relativePath: string) => string,
): string {
  const { width, height, backgroundCss, shapesHtml } = buildSlideState(
    slide,
    animationStep,
    mediaResolver,
  );

  return `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<style>
*{margin:0;padding:0;box-sizing:border-box}
html,body{width:100%;height:100%;overflow:hidden;background:#000}
${ANIMATION_KEYFRAMES}
.slide-root{
  position:absolute;
  left:0;top:0;
  width:${width}px;
  height:${height}px;
  background:${backgroundCss};
  transform-origin:0 0;
  overflow:hidden;
  font-family:Calibri,Arial,Helvetica,sans-serif;
  font-size:18pt;
  color:#000;
}
</style>
<script>${makeFitScript(width, height)}</script>
</head>
<body>
<div class="slide-root">
${shapesHtml}
</div>
</body>
</html>`;
}

/**
 * Alias for `renderPptxSlideToHtml`. Kept for backwards compatibility with
 * call sites that use the preview-specific name.
 */
export function renderPptxSlidePreviewHtml(
  slide: PptxSlideData,
  animationStep: number,
  mediaResolver: (relativePath: string) => string,
): string {
  return renderPptxSlideToHtml(slide, animationStep, mediaResolver);
}
```

- [ ] **Step 2: Verify type-check passes**

```bash
cd /Users/gabriel/Developer/WebPresent
npm run build:engine && npx tsc -b --noEmit
```

Expected: no errors.

- [ ] **Step 3: Run full regression suite**

```bash
node --test test/pptxParser.test.cjs test/pptxPresentRenderer.test.cjs test/electronUtils.test.cjs
```

Expected: all tests pass.

- [ ] **Step 4: Commit**

```bash
git add src/pptxRenderer.ts
git commit -m "refactor(renderer): replace duplicated rendering logic with @webpresent/pptx-engine"
```

---

## Task 6: Update `electron/pptxPresentRenderer.cjs`

**Files:**
- Modify: `electron/pptxPresentRenderer.cjs`

- [ ] **Step 1: Replace `electron/pptxPresentRenderer.cjs` with package-backed implementation**

The entire file becomes:

```js
const path = require('node:path');
const fs = require('node:fs/promises');
const { existsSync } = require('node:fs');
const { readFileAsDataUrl } = require('./utils.cjs');
const { buildSlideState, ANIMATION_KEYFRAMES } = require('@webpresent/pptx-engine');

// ── Media collection + data-URL resolution ────────────────────────────────────

/**
 * Collects all relative media paths referenced by a slide (background,
 * shape images, fill images). Used to prefetch all needed data URLs.
 */
function collectSlideMediaPaths(slide) {
  const paths = new Set();
  if (slide.background?.imageRelativePath) {
    paths.add(slide.background.imageRelativePath);
  }
  for (const shape of slide.shapes || []) {
    if (shape.imageRelativePath) paths.add(shape.imageRelativePath);
    if (shape.fill?.imageRelativePath) paths.add(shape.fill.imageRelativePath);
  }
  return [...paths];
}

/**
 * Reads media files from disk and returns a map of relative path → data URL.
 * An optional `cache` Map<string, Promise<string>> avoids redundant disk reads
 * across multiple slides in the same presentation session.
 */
async function buildMediaDataUrlMap(deckDir, relativePaths, options = {}) {
  const readFile = options.readFile || fs.readFile;
  const fileExists = options.existsSync || existsSync;
  const inlineMedia = options.readFileAsDataUrl || readFileAsDataUrl;
  const cache = options.cache;
  const mediaUrls = {};

  for (const relativePath of relativePaths) {
    if (!relativePath) continue;
    if (cache?.has(relativePath)) {
      mediaUrls[relativePath] = await cache.get(relativePath);
      continue;
    }
    const absolutePath = path.join(deckDir, relativePath);
    if (!fileExists(absolutePath)) continue;
    const mediaUrlPromise = inlineMedia(absolutePath, { readFile });
    if (cache) cache.set(relativePath, mediaUrlPromise);
    mediaUrls[relativePath] = await mediaUrlPromise;
  }

  return mediaUrls;
}

// ── Slide HTML builder ────────────────────────────────────────────────────────

/**
 * Builds the full presentation HTML document for a single PPTX slide step.
 *
 * The generated document:
 * - Embeds all media as data URLs (no external requests)
 * - Scales the slide to fill the window via a fit() function
 * - Defines `window.__WEBPRESENT_UPDATE_PPTX(state)` for in-place updates
 *   (avoids a full page reload when navigating between PPTX steps)
 */
function buildPptxSlideHtml(slideData, animationStep, mediaUrls) {
  const mediaResolver = (relativePath) =>
    relativePath ? mediaUrls[relativePath] || '' : '';
  const state = buildSlideState(slideData, animationStep, mediaResolver);
  const initialState = JSON.stringify(state);

  return `<!doctype html><html><head><meta charset="utf-8"/><style>
*{margin:0;padding:0;box-sizing:border-box}
html,body{width:100%;height:100%;overflow:hidden;background:#000}
${ANIMATION_KEYFRAMES}
.slide-root{position:absolute;left:0;top:0;transform-origin:0 0;overflow:hidden;font-family:Calibri,Arial,Helvetica,sans-serif;font-size:18pt;color:#000}
</style><script>
(function(){
  function fit(){
    var r=document.querySelector('.slide-root');
    if(!r)return;
    var sw=Number(r.dataset.slideWidth)||1,sh=Number(r.dataset.slideHeight)||1;
    var vw=window.innerWidth||document.documentElement.clientWidth||sw;
    var vh=window.innerHeight||document.documentElement.clientHeight||sh;
    if(vw<1||vh<1)return;
    var sc=Math.min(vw/sw,vh/sh);
    r.style.transform='scale('+sc+')';
    r.style.left=Math.round((vw-sw*sc)/2)+'px';
    r.style.top=Math.round((vh-sh*sc)/2)+'px';
  }
  window.__WEBPRESENT_UPDATE_PPTX=function(state){
    var r=document.querySelector('.slide-root');
    if(!r||!state)return;
    r.dataset.slideWidth=state.width;
    r.dataset.slideHeight=state.height;
    r.style.width=state.width+'px';
    r.style.height=state.height+'px';
    r.style.background=state.backgroundCss;
    r.innerHTML=state.shapesHtml||'';
    fit();
  };
  window.addEventListener('resize',fit);
  window.addEventListener('load',fit);
  setTimeout(fit,0);
  setTimeout(fit,50);
  setTimeout(fit,200);
  if(window.ResizeObserver){new ResizeObserver(fit).observe(document.documentElement);}
  window.__WEBPRESENT_UPDATE_PPTX(${initialState});
})()
</script></head><body><div class="slide-root"></div></body></html>`;
}

// ── Public API ────────────────────────────────────────────────────────────────

async function buildPptxPresentationDocument(slideData, animationStep, deckDir) {
  const relativePaths = collectSlideMediaPaths(slideData);
  const mediaUrls = await buildMediaDataUrlMap(deckDir, relativePaths);
  return buildPptxSlideHtml(slideData, animationStep, mediaUrls);
}

async function buildPptxRuntimeUpdateScript(slideData, animationStep, deckDir, options = {}) {
  const relativePaths = collectSlideMediaPaths(slideData);
  const mediaUrls = await buildMediaDataUrlMap(deckDir, relativePaths, options);
  const mediaResolver = (p) => (p ? mediaUrls[p] || '' : '');
  return `window.__WEBPRESENT_UPDATE_PPTX(${JSON.stringify(buildSlideState(slideData, animationStep, mediaResolver))});`;
}

/**
 * Creates a cached builder for a single presentation session.
 *
 * Caches:
 * - `mediaCache` — data URLs per relative path (shared across all slides)
 * - `htmlCache` — full HTML per (slideIndex, animationStep)
 * - `updateScriptCache` — update script per (slideIndex, animationStep)
 */
function createPptxPresentationDocumentBuilder(deckDir, options = {}) {
  const mediaCache = new Map();
  const htmlCache = new Map();
  const updateScriptCache = new Map();

  const builder = async function buildCachedPptxPresentationDocument(slideData, animationStep) {
    const cacheKey = `${slideData.slideIndex}:${animationStep}`;
    if (htmlCache.has(cacheKey)) return htmlCache.get(cacheKey);

    const relativePaths = collectSlideMediaPaths(slideData);
    const mediaUrls = await buildMediaDataUrlMap(deckDir, relativePaths, {
      ...options,
      cache: mediaCache,
    });
    const html = buildPptxSlideHtml(slideData, animationStep, mediaUrls);
    htmlCache.set(cacheKey, html);
    return html;
  };

  builder.buildUpdateScript = async function buildCachedPptxRuntimeUpdateScript(
    slideData,
    animationStep,
  ) {
    const cacheKey = `${slideData.slideIndex}:${animationStep}`;
    if (updateScriptCache.has(cacheKey)) return updateScriptCache.get(cacheKey);

    const relativePaths = collectSlideMediaPaths(slideData);
    const mediaUrls = await buildMediaDataUrlMap(deckDir, relativePaths, {
      ...options,
      cache: mediaCache,
    });
    const mediaResolver = (p) => (p ? mediaUrls[p] || '' : '');
    const script = `window.__WEBPRESENT_UPDATE_PPTX(${JSON.stringify(buildSlideState(slideData, animationStep, mediaResolver))});`;
    updateScriptCache.set(cacheKey, script);
    return script;
  };

  return builder;
}

module.exports = {
  buildMediaDataUrlMap,
  buildPptxPresentationDocument,
  buildPptxRuntimeUpdateScript,
  createPptxPresentationDocumentBuilder,
  collectSlideMediaPaths,
};
```

- [ ] **Step 2: Run full regression suite**

```bash
node --test test/pptxParser.test.cjs test/pptxPresentRenderer.test.cjs test/electronUtils.test.cjs
```

Expected: all tests pass. If any `pptxPresentRenderer` test fails, verify the `buildSlideState` output shape matches what the tests expect (`width`, `height`, `backgroundCss`, `shapesHtml`).

- [ ] **Step 3: Run full build**

```bash
npm run build
```

Expected: no errors.

- [ ] **Step 4: Commit**

```bash
git add electron/pptxPresentRenderer.cjs
git commit -m "refactor(pptxPresentRenderer): replace duplicated rendering logic with @webpresent/pptx-engine"
```

---

## Task 7: Migrate parser to TypeScript

> **Risk note:** This is the highest-risk task. `pptxParser.cjs` is 1702 lines of complex OOXML parsing with carefully tuned `isArray` lists. The strategy is: copy the logic faithfully, add types at function boundaries only (not inside every helper), and run the full regression suite at every checkpoint. Do not simplify or restructure the logic as part of this migration.

**Files:**
- Create: `packages/pptx-engine/src/parser.ts`
- Modify: `packages/pptx-engine/src/index.ts`
- Modify: `electron/pptxParser.cjs` (replace with thin shim)

- [ ] **Step 1: Add `@types/node` to pptx-engine dev deps**

In `packages/pptx-engine/package.json`, add to `devDependencies`:

```json
"@types/node": "^22.10.1"
```

Run `npm install` from the workspace root.

- [ ] **Step 2: Create `packages/pptx-engine/src/parser.ts`**

This is a mechanical TypeScript migration of `electron/pptxParser.cjs`. Rules:

1. Replace `const X = require('X')` with `import X from 'X'` (or named imports).
2. Replace `module.exports = { parsePptx }` with `export { parsePptx }`.
3. Add the `PptxDeckData` return type annotation to `parsePptx`.
4. Add `string` / `number` / `boolean` annotations only to the top-level public function parameter `filePath: string` and `deckId: string` and `getDeckDir: (id: string) => string`.
5. Every other function stays loosely typed (`any` is acceptable for internal helpers during this migration pass).
6. Do NOT change logic, rename variables, or restructure code.

Start the file:

```ts
import AdmZip from 'adm-zip';
import { XMLParser } from 'fast-xml-parser';
import path from 'node:path';
import fs from 'node:fs/promises';
import { writeFileSync } from 'node:fs';
import type { PptxDeckData } from './types';
```

Copy all remaining content from `electron/pptxParser.cjs` verbatim after the imports, with these minimal changes:
- Remove the `require()` calls at the top (replaced by imports above)
- Change the final `module.exports = { parsePptx }` to `export { parsePptx }`
- Add `: Promise<PptxDeckData>` return type annotation to `async function parsePptx`
- In `parsePptx`, change `const fs = require('node:fs/promises')` and `const { writeFileSync } = require('node:fs')` to use the already-imported `fs` and `writeFileSync` (remove any local re-requires)

- [ ] **Step 3: Build engine and verify type compilation**

```bash
cd /Users/gabriel/Developer/WebPresent
npm run build:engine 2>&1 | head -60
```

TypeScript may emit `any` warnings but should not fail with errors. If it does, check that `skipLibCheck: true` is in `packages/pptx-engine/tsconfig.json` and that all imports resolve. Fix any import-related errors (do not fix logic).

- [ ] **Step 4: Export parsePptx from index**

In `packages/pptx-engine/src/index.ts`, add at the end:

```ts
export { parsePptx } from './parser';
```

- [ ] **Step 5: Replace `electron/pptxParser.cjs` with a thin shim**

```js
/**
 * Thin compatibility shim — delegates to @webpresent/pptx-engine.
 *
 * All PPTX parsing logic now lives in packages/pptx-engine/src/parser.ts.
 * This file exists only so that any existing require('./pptxParser.cjs')
 * calls continue to work without changes.
 */
const { parsePptx } = require('@webpresent/pptx-engine');

module.exports = { parsePptx };
```

- [ ] **Step 6: Run full regression suite**

```bash
node --test test/pptxParser.test.cjs test/pptxPresentRenderer.test.cjs test/electronUtils.test.cjs
```

Expected: all tests pass with no regressions. If a test fails, compare the failing output against the original `pptxParser.cjs` implementation — the cause will be an import or module-resolution difference, not a logic change.

- [ ] **Step 7: Run full build**

```bash
npm run build
```

Expected: no errors.

- [ ] **Step 8: Commit**

```bash
git add packages/pptx-engine/src/parser.ts packages/pptx-engine/src/index.ts electron/pptxParser.cjs packages/pptx-engine/package.json
git commit -m "feat(pptx-engine): migrate pptxParser to TypeScript; pptxParser.cjs becomes a shim"
```

---

## Task 8: Update AGENTS.md

**Files:**
- Modify: `AGENTS.md`

- [ ] **Step 1: Update the Important Files section**

Find the "## Important Files" section in `AGENTS.md`. After the existing entries, update or add:

```markdown
- `packages/pptx-engine/`
  The `@webpresent/pptx-engine` npm workspace package. Contains all PPTX types,
  shared rendering logic, animation CSS, image helpers, and the parser.
  Built by tsup to CJS + ESM. Both `electron/` and `src/` consumers import from here.
  - `src/types.ts` — all PPTX data model types
  - `src/parser.ts` — OOXML parser (migrated from `electron/pptxParser.cjs`)
  - `src/renderCore.ts` — `buildSlideState()` and rendering primitives
  - `src/animationCss.ts` — CSS @keyframes for entrance effects
  - `src/imageHelpers.ts` — image crop / flip CSS helpers
```

Replace the entry for `electron/pptxParser.cjs` with:

```markdown
- `electron/pptxParser.cjs`
  Thin shim — re-exports `parsePptx` from `@webpresent/pptx-engine`.
  All parsing logic lives in `packages/pptx-engine/src/parser.ts`.
```

- [ ] **Step 2: Update the Architecture Notes section**

Update the "The repo mixes module systems intentionally" paragraph:

```markdown
The repo mixes module systems intentionally:

- Renderer code is TypeScript/ESM under `src/`.
- Electron code is CommonJS under `electron/*.cjs`.
- `packages/pptx-engine/` is TypeScript built to both CJS and ESM by tsup.
  Electron uses the CJS build; Vite uses the TypeScript source directly via a
  resolve alias in `vite.config.ts`.
- `package.json` uses `"type": "module"`, but Electron files remain `.cjs` and
  must stay CommonJS unless the task explicitly migrates them.
```

- [ ] **Step 3: Update the How To Run section**

Add a note under "Install":

```markdown
- Install + build engine: `npm install && npm run build:engine`
```

- [ ] **Step 4: Commit**

```bash
git add AGENTS.md
git commit -m "docs: update AGENTS.md to reflect @webpresent/pptx-engine package structure"
```

---

## Final Verification

- [ ] **Run full test suite**

```bash
node --test test/pptxParser.test.cjs test/pptxPresentRenderer.test.cjs test/electronUtils.test.cjs
```

All tests must pass.

- [ ] **Run full build**

```bash
npm run build
```

No errors.

- [ ] **Smoke test in Electron dev mode**

```bash
npm run dev
```

Open a deck containing PPTX slides. Verify:
1. Editor preview renders the slide correctly
2. Starting a presentation shows the slide correctly
3. Advancing an animation step shows the animated shape with a CSS transition
4. Advancing past the last PPTX step works

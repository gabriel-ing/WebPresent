/**
 * Browser-safe entry point for @webpresent/pptx-engine.
 *
 * Excludes the PPTX parser (which depends on Node.js `fs`, `path`, and
 * `adm-zip`) so this file can be imported in Vite's renderer without
 * bundling Node-only dependencies into the browser bundle.
 *
 * Use this entry for the Vite alias in vite.config.ts.
 * The full entry (src/index.ts) is used by the Electron main process via
 * the CJS dist build.
 */
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

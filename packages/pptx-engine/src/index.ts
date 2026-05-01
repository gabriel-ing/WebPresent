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
export { parsePptx } from './parser';

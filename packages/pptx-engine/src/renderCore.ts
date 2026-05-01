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

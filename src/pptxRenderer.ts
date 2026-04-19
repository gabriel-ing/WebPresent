/**
 * Renders a PptxSlideData object to a self-contained HTML string.
 *
 * Used both for the editor preview (via iframe srcdoc) and the
 * presentation window (loaded as a data: URL or temp file).
 */

import type { PptxSlideData, PptxShape, PptxFill, PptxParagraph, PptxTextRun } from './types';

// ── CSS animation keyframes ──────────────────────────────────────────────────

const ANIMATION_KEYFRAMES = `
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

// ── Helpers ──────────────────────────────────────────────────────────────────

function escHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escAttr(str: string): string {
  return escHtml(str);
}

// ── Fill → CSS ───────────────────────────────────────────────────────────────

function fillToCss(fill: PptxFill | undefined, mediaResolver?: (path: string) => string): string {
  if (!fill || fill.type === 'none') return 'transparent';
  if (fill.type === 'solid' && fill.colour) return fill.colour;
  if (fill.type === 'image' && fill.imageRelativePath) {
    const url = mediaResolver ? mediaResolver(fill.imageRelativePath) : '';
    return url ? `url("${escAttr(url)}") center/cover no-repeat` : '#cccccc';
  }
  if (fill.type === 'gradient' && fill.gradientStops?.length) {
    const stops = fill.gradientStops
      .map((s) => `${s.colour} ${Math.round(s.position)}%`)
      .join(', ');
    return `linear-gradient(180deg, ${stops})`;
  }
  return 'transparent';
}

// ── Text runs → HTML ─────────────────────────────────────────────────────────

function renderTextRun(run: PptxTextRun): string {
  if (run.text === '\n') return '<br/>';
  const styles: string[] = [];
  if (run.bold) styles.push('font-weight:bold');
  if (run.italic) styles.push('font-style:italic');
  if (run.underline) styles.push('text-decoration:underline');
  if (run.fontSize) styles.push(`font-size:${run.fontSize}pt`);
  if (run.fontFamily) styles.push(`font-family:${escAttr(run.fontFamily)},sans-serif`);
  if (run.colour) styles.push(`color:${run.colour}`);

  const text = escHtml(run.text);
  if (!styles.length) return `<span>${text}</span>`;
  return `<span style="${styles.join(';')}">${text}</span>`;
}

function renderParagraph(para: PptxParagraph): string {
  const styles: string[] = [];
  styles.push('margin:0');
  styles.push('padding:0 0 0.15em 0');
  if (para.alignment) styles.push(`text-align:${para.alignment}`);
  if (para.level && para.level > 0) {
    styles.push(`padding-left:${para.level * 1.5}em`);
  }

  const runsHtml = para.runs.map(renderTextRun).join('');

  let prefix = '';
  if (para.bulletType === 'bullet') {
    const char = para.bulletChar || '•';
    prefix = `<span style="margin-right:0.4em">${escHtml(char)}</span>`;
  }

  return `<p style="${styles.join(';')}">${prefix}${runsHtml}</p>`;
}

function clampCrop(value: number | undefined): number {
  const number = Number(value) || 0;
  return Math.min(0.9999, Math.max(0, number));
}

function getCroppedImageStyles(shape: PptxShape): string {
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

// ── Shape → HTML element ─────────────────────────────────────────────────────

function renderShape(
  shape: PptxShape,
  animationStep: number,
  mediaResolver?: (path: string) => string,
): string {
  // Determine visibility
  if (shape.animationGroup > animationStep) {
    return ''; // not yet visible at this animation step
  }

  const isNewlyRevealed = shape.animationGroup === animationStep && animationStep > 0;
  const visibleParagraphs = (shape.paragraphs || []).filter((paragraph) => (paragraph.animationGroup || 0) <= animationStep);

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

  if (shape.rotation) {
    styles.push(`transform:rotate(${shape.rotation}deg)`);
  }

  // Fill
  if (shape.type !== 'image' && shape.type !== 'line') {
    const bg = fillToCss(shape.fill, mediaResolver);
    if (bg !== 'transparent') {
      styles.push(`background:${bg}`);
    }
  }

  // Border
  if (shape.border) {
    styles.push(`border:${shape.border.width}pt ${shape.border.style || 'solid'} ${shape.border.colour}`);
  }

  // Corner radius
  if (shape.type === 'roundRect' && shape.cornerRadius) {
    styles.push(`border-radius:${shape.cornerRadius}px`);
  }
  if (shape.type === 'ellipse') {
    styles.push('border-radius:50%');
  }

  // Animation
  if (isNewlyRevealed) {
    const effect = shape.animationEffect || 'appear';
    styles.push(`animation:pptx-${effect} 0.5s ease both`);
  }

  // Flexbox for text layout
  if (visibleParagraphs.length) {
    styles.push('display:flex');
    styles.push('flex-direction:column');
    // Use vertical alignment from shape data, default to top
    const vAlign = shape.verticalAlign || 'top';
    if (vAlign === 'middle') {
      styles.push('justify-content:center');
    } else if (vAlign === 'bottom') {
      styles.push('justify-content:flex-end');
    } else {
      styles.push('justify-content:flex-start');
    }
    styles.push('padding:4px 8px');
  }

  // Image shapes
  if (shape.type === 'image' && shape.imageRelativePath) {
    const url = mediaResolver ? mediaResolver(shape.imageRelativePath) : '';
    const imageStyles = shape.imageCrop || shape.flipH || shape.flipV
      ? getCroppedImageStyles(shape)
      : 'width:100%;height:100%;object-fit:contain;display:block';
    return `<div style="${styles.join(';')}"><img src="${escAttr(url)}" alt="" style="${imageStyles}"/></div>`;
  }

  // Line shapes
  if (shape.type === 'line') {
    const colour = shape.border?.colour || shape.fill?.colour || '#000000';
    const weight = shape.border?.width || 1;
    styles.push(`border-top:${weight}pt solid ${colour}`);
    styles.push('height:0');
    return `<div style="${styles.join(';')}"></div>`;
  }

  // Text content
  const inner = visibleParagraphs.map(renderParagraph).join('');
  return `<div style="${styles.join(';')}">${inner}</div>`;
}

// ── Full slide → HTML document ───────────────────────────────────────────────

export function renderPptxSlideToHtml(
  slide: PptxSlideData,
  animationStep: number,
  mediaResolver: (relativePath: string) => string,
): string {
  const bgCss = slide.background
    ? fillToCss(slide.background, mediaResolver)
    : '#ffffff';

  const shapesHtml = slide.shapes
    .map((shape) => renderShape(shape, animationStep, mediaResolver))
    .join('\n');

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
  width:${slide.width}px;
  height:${slide.height}px;
  background:${bgCss};
  transform-origin:0 0;
  overflow:hidden;
  font-family:Calibri,Arial,Helvetica,sans-serif;
  font-size:18pt;
  color:#000;
}
</style>
<script>
(function(){
  function fit(){
    var root=document.querySelector('.slide-root');
    if(!root)return;
    var sw=${slide.width},sh=${slide.height};
    var vw=window.innerWidth||document.documentElement.clientWidth||${slide.width};
    var vh=window.innerHeight||document.documentElement.clientHeight||${slide.height};
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
  if(window.ResizeObserver){
    new ResizeObserver(fit).observe(document.documentElement);
  }
})();
</script>
</head>
<body>
<div class="slide-root">
${shapesHtml}
</div>
</body>
</html>`;
}

/**
 * Renders a preview-friendly HTML snippet (not a full document) for embedding
 * in the editor preview pane via an iframe srcdoc.
 */
export function renderPptxSlidePreviewHtml(
  slide: PptxSlideData,
  animationStep: number,
  mediaResolver: (relativePath: string) => string,
): string {
  return renderPptxSlideToHtml(slide, animationStep, mediaResolver);
}

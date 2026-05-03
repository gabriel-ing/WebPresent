import type {
  PptxBorder,
  PptxFill,
  PptxGradientStop,
  PptxParagraph,
  PptxShape,
  PptxSlideData,
  PptxTextRun,
} from './types';
import { getCroppedImageStyles } from './imageHelpers';

export type SlideState = {
  width: number;
  height: number;
  backgroundCss: string;
  shapesHtml: string;
};

export function escHtml(s: string): string {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatNumeric(value: number): number {
  return Number(value.toFixed(3));
}

function hexToRgb(colour: string): [number, number, number] | undefined {
  const normalized = colour.replace(/^#/, '');
  if (!/^[0-9a-fA-F]{6}$/.test(normalized)) return undefined;
  return [
    Number.parseInt(normalized.slice(0, 2), 16),
    Number.parseInt(normalized.slice(2, 4), 16),
    Number.parseInt(normalized.slice(4, 6), 16),
  ];
}

function applyOpacity(colour: string, opacity?: number): string {
  if (opacity === undefined || opacity >= 1) return colour;
  const rgb = hexToRgb(colour);
  if (!rgb) return colour;
  return `rgba(${rgb[0]}, ${rgb[1]}, ${rgb[2]}, ${formatNumeric(opacity)})`;
}

function sortGradientStops(stops: PptxGradientStop[]): PptxGradientStop[] {
  return [...stops].sort((left, right) => left.position - right.position);
}

function gradientStopsToCss(stops: PptxGradientStop[]): string {
  return sortGradientStops(stops)
    .map((stop) => `${applyOpacity(stop.colour, stop.opacity)} ${formatNumeric(stop.position)}%`)
    .join(', ');
}

function gradientAngle(value: { gradientAngle?: number } | undefined, fallback = 180): number {
  return typeof value?.gradientAngle === 'number' ? formatNumeric(value.gradientAngle) : fallback;
}

function gradientVector(angle: number, width: number, height: number) {
  const radians = (angle * Math.PI) / 180;
  const dx = Math.sin(radians);
  const dy = -Math.cos(radians);
  const halfWidth = width / 2;
  const halfHeight = height / 2;
  const scaleX = Math.abs(dx) > 0.000001 ? halfWidth / Math.abs(dx) : Number.POSITIVE_INFINITY;
  const scaleY = Math.abs(dy) > 0.000001 ? halfHeight / Math.abs(dy) : Number.POSITIVE_INFINITY;
  const extent = Math.min(scaleX, scaleY);
  const centerX = halfWidth;
  const centerY = halfHeight;

  return {
    x1: formatNumeric(centerX - dx * extent),
    y1: formatNumeric(centerY - dy * extent),
    x2: formatNumeric(centerX + dx * extent),
    y2: formatNumeric(centerY + dy * extent),
  };
}

function buildSvgGradient(id: string, stops: PptxGradientStop[], angle: number, width: number, height: number): string {
  const { x1, y1, x2, y2 } = gradientVector(angle, width, height);
  const stopMarkup = sortGradientStops(stops)
    .map((stop) => {
      const opacity = stop.opacity !== undefined && stop.opacity < 1 ? ` stop-opacity="${formatNumeric(stop.opacity)}"` : '';
      return `<stop offset="${formatNumeric(stop.position)}%" stop-color="${escHtml(stop.colour)}"${opacity}/>`;
    })
    .join('');
  return `<linearGradient id="${escHtml(id)}" gradientUnits="userSpaceOnUse" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stopMarkup}</linearGradient>`;
}

function buildSvgPaint(
  id: string,
  fill: Pick<PptxFill, 'type' | 'colour' | 'gradientStops' | 'gradientAngle'> | undefined,
  width: number,
  height: number,
): { value: string; defs?: string } | undefined {
  if (!fill || fill.type === 'none') return undefined;
  if (fill.type === 'solid' && fill.colour) return { value: fill.colour };
  if (fill.type === 'gradient' && fill.gradientStops?.length) {
    return {
      value: `url(#${id})`,
      defs: buildSvgGradient(id, fill.gradientStops, gradientAngle(fill), width, height),
    };
  }
  return undefined;
}

function buildSvgBorderPaint(
  id: string,
  border: PptxBorder | undefined,
  width: number,
  height: number,
): { value: string; defs?: string } | undefined {
  if (!border) return undefined;
  if (border.gradientStops?.length) {
    return {
      value: `url(#${id})`,
      defs: buildSvgGradient(id, border.gradientStops, gradientAngle(border), width, height),
    };
  }
  if (border.colour) return { value: border.colour };
  return undefined;
}

function paragraphHasVisibleText(paragraph: PptxParagraph): boolean {
  return paragraph.runs.some((run) => run.text !== '' && run.text !== '\n');
}

function trimTrailingEmptyParagraphs(paragraphs: PptxParagraph[]): PptxParagraph[] {
  let lastVisibleIndex = paragraphs.length - 1;
  while (lastVisibleIndex >= 0 && !paragraphHasVisibleText(paragraphs[lastVisibleIndex])) {
    lastVisibleIndex -= 1;
  }
  return lastVisibleIndex >= 0 ? paragraphs.slice(0, lastVisibleIndex + 1) : [];
}

function buildTextPadding(shape: PptxShape): string {
  const top = shape.textInsets?.top ?? 4;
  const right = shape.textInsets?.right ?? 8;
  const bottom = shape.textInsets?.bottom ?? 4;
  const left = shape.textInsets?.left ?? 8;
  return `padding:${formatNumeric(top)}px ${formatNumeric(right)}px ${formatNumeric(bottom)}px ${formatNumeric(left)}px`;
}

function buildTextOverlayStyles(shape: PptxShape): string {
  const styles = ['position:absolute', 'inset:0', 'display:flex', 'flex-direction:column'];
  const verticalAlign = shape.verticalAlign || 'top';
  if (verticalAlign === 'middle') styles.push('justify-content:center');
  else if (verticalAlign === 'bottom') styles.push('justify-content:flex-end');
  else styles.push('justify-content:flex-start');
  styles.push(buildTextPadding(shape));
  return styles.join(';');
}

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
    return `linear-gradient(${gradientAngle(fill)}deg, ${gradientStopsToCss(fill.gradientStops)})`;
  }
  return 'transparent';
}

export function renderTextRun(run: PptxTextRun, fontScale = 1): string {
  if (run.text === '\n') return '<br/>';
  const styles = buildTextRunStyles(run, fontScale);
  const text = escHtml(run.text);
  return styles.length ? `<span style="${styles.join(';')}">${text}</span>` : `<span>${text}</span>`;
}

function fontFamilyToCssStack(fontFamily: string): string {
  const normalized = fontFamily.trim().toLowerCase();

  if (normalized === 'gotham book' || normalized === 'gotham' || normalized.startsWith('gotham ')) {
    return `${fontFamily},Avenir Next,Avenir,Helvetica Neue,Arial,sans-serif`;
  }

  if (normalized === 'segoe script') {
    return `${fontFamily},Bradley Hand,Marker Felt,Snell Roundhand,Apple Chancery,cursive`;
  }

  if (normalized === 'segoe print') {
    return `${fontFamily},Bradley Hand,Marker Felt,Comic Sans MS,cursive`;
  }

  return `${fontFamily},sans-serif`;
}

function buildTextRunStyles(run: PptxTextRun, fontScale = 1): string[] {
  const styles: string[] = [];
  if (run.bold) styles.push('font-weight:bold');
  if (run.italic) styles.push('font-style:italic');
  if (run.underline) styles.push('text-decoration:underline');
  if (run.fontSize) styles.push(`font-size:${formatNumeric(run.fontSize * fontScale)}pt`);
  if (run.fontFamily) styles.push(`font-family:${escHtml(fontFamilyToCssStack(run.fontFamily))}`);
  if (run.colour) styles.push(`color:${run.colour}`);
  return styles;
}

export function renderParagraph(para: PptxParagraph, fontScale = 1): string {
  const styles = ['margin:0', 'padding:0'];
  if (para.alignment) styles.push(`text-align:${para.alignment}`);
  if (para.lineSpacing && para.lineSpacing > 0) styles.push(`line-height:${formatNumeric(para.lineSpacing)}`);
  else styles.push('line-height:1');
  if (para.level && para.level > 0) styles.push(`padding-left:${para.level * 1.5}em`);
  const hasVisibleText = paragraphHasVisibleText(para);
  const runs = para.runs.map((run) => renderTextRun(run, fontScale)).join('');
  const bulletRun = para.runs.find((run) => run.text && run.text !== '\n');
  const bulletStyles = bulletRun ? buildTextRunStyles(bulletRun, fontScale) : [];
  bulletStyles.push('margin-right:0.4em');
  const bullet =
    para.bulletType === 'bullet' && hasVisibleText
      ? `<span style="${bulletStyles.join(';')}">${escHtml(para.bulletChar || '•')}</span>`
      : '';
  return `<p style="${styles.join(';')}">${bullet}${hasVisibleText ? runs : '&nbsp;'}</p>`;
}

function renderFreeformShape(shape: PptxShape, styles: string[], innerHtml: string): string {
  if (!shape.svgPath) return `<div style="${styles.join(';')}">${innerHtml}</div>`;

  const viewBoxWidth = shape.svgViewBoxWidth || shape.width || 1;
  const viewBoxHeight = shape.svgViewBoxHeight || shape.height || 1;
  const defs: string[] = [];
  const fillPaint = buildSvgPaint(`${shape.id}-fill`, shape.fill, viewBoxWidth, viewBoxHeight);
  const strokePaint = buildSvgBorderPaint(`${shape.id}-stroke`, shape.border, viewBoxWidth, viewBoxHeight);

  if (fillPaint?.defs) defs.push(fillPaint.defs);
  if (strokePaint?.defs) defs.push(strokePaint.defs);

  const pathAttributes = [`d="${escHtml(shape.svgPath)}"`, `fill="${escHtml(fillPaint?.value || 'none')}"`];

  if (strokePaint?.value && shape.border?.width) {
    pathAttributes.push(
      `stroke="${escHtml(strokePaint.value)}"`,
      `stroke-width="${formatNumeric((shape.border.width * 96) / 72)}"`,
      'vector-effect="non-scaling-stroke"',
    );
    if (shape.border.style === 'dashed') pathAttributes.push('stroke-dasharray="6 4"');
    if (shape.border.style === 'dotted') pathAttributes.push('stroke-dasharray="1 4"', 'stroke-linecap="round"');
  }

  const svg = [
    `<svg width="100%" height="100%" viewBox="0 0 ${formatNumeric(viewBoxWidth)} ${formatNumeric(viewBoxHeight)}" preserveAspectRatio="none" aria-hidden="true">`,
    defs.length ? `<defs>${defs.join('')}</defs>` : '',
    `<path ${pathAttributes.join(' ')}/>`,
    '</svg>',
  ].join('');

  const overlay = innerHtml ? `<div style="${buildTextOverlayStyles(shape)}">${innerHtml}</div>` : '';
  return `<div style="${styles.join(';')}">${svg}${overlay}</div>`;
}

function buildLineMarker(id: string, colour: string, strokeWidthPx: number, position: 'start' | 'end'): string {
  const arrowSize = Math.max(strokeWidthPx * 2.5, 10);
  const halfSize = arrowSize / 2;
  const path =
    position === 'start'
      ? `M ${formatNumeric(arrowSize)} 0 L 0 ${formatNumeric(halfSize)} L ${formatNumeric(arrowSize)} ${formatNumeric(arrowSize)} Z`
      : `M 0 0 L ${formatNumeric(arrowSize)} ${formatNumeric(halfSize)} L 0 ${formatNumeric(arrowSize)} Z`;
  const refX = position === 'start' ? 0 : arrowSize;
  return `<marker id="${escHtml(id)}" markerWidth="${formatNumeric(arrowSize)}" markerHeight="${formatNumeric(arrowSize)}" refX="${formatNumeric(refX)}" refY="${formatNumeric(halfSize)}" orient="auto" markerUnits="userSpaceOnUse"><path d="${path}" fill="${escHtml(colour)}"/></marker>`;
}

function renderLineShape(shape: PptxShape, styles: string[]): string {
  const strokeWidthPt = shape.border?.width || 1;
  const strokeWidthPx = (strokeWidthPt * 96) / 72;
  const markerSize =
    shape.lineHead === 'triangle' || shape.lineTail === 'triangle'
      ? Math.max(strokeWidthPx * 2.5, 10)
      : 0;
  const collapsedAxisSize = Math.max(strokeWidthPx, markerSize, 1);
  const lineWidth = Math.max(shape.width, 1);
  const lineHeight = Math.max(shape.height, 1);
  const renderWidth = shape.width === 0 ? collapsedAxisSize : lineWidth;
  const renderHeight = shape.height === 0 ? collapsedAxisSize : lineHeight;
  const offsetX = shape.width === 0 ? renderWidth / 2 : 0;
  const offsetY = shape.height === 0 ? renderHeight / 2 : 0;
  const strokePaint = buildSvgBorderPaint(`${shape.id}-stroke`, shape.border, renderWidth, renderHeight);
  const stroke = strokePaint?.value || shape.fill?.colour || '#000000';
  const defs: string[] = [];

  if (strokePaint?.defs) defs.push(strokePaint.defs);
  if (shape.lineHead === 'triangle') defs.push(buildLineMarker(`${shape.id}-head`, stroke, strokeWidthPx, 'start'));
  if (shape.lineTail === 'triangle') defs.push(buildLineMarker(`${shape.id}-tail`, stroke, strokeWidthPx, 'end'));

  const x1 = shape.width > 0 ? (shape.flipH ? lineWidth : 0) : offsetX;
  const y1 = shape.height > 0 ? (shape.flipV ? lineHeight : 0) : offsetY;
  const x2 = shape.width > 0 ? (shape.flipH ? 0 : lineWidth) : offsetX;
  const y2 = shape.height > 0 ? (shape.flipV ? 0 : lineHeight) : offsetY;

  const lineAttributes = [
    `x1="${formatNumeric(x1)}"`,
    `y1="${formatNumeric(y1)}"`,
    `x2="${formatNumeric(x2)}"`,
    `y2="${formatNumeric(y2)}"`,
    `stroke="${escHtml(stroke)}"`,
    `stroke-width="${formatNumeric(strokeWidthPx)}"`,
    'fill="none"',
  ];

  if (shape.border?.style === 'dashed') lineAttributes.push('stroke-dasharray="6 4"');
  if (shape.border?.style === 'dotted') lineAttributes.push('stroke-dasharray="1 4"', 'stroke-linecap="round"');
  if (shape.lineHead === 'triangle') lineAttributes.push(`marker-start="url(#${shape.id}-head)"`);
  if (shape.lineTail === 'triangle') lineAttributes.push(`marker-end="url(#${shape.id}-tail)"`);

  styles.push(
    `left:${formatNumeric(shape.x - offsetX)}px`,
    `top:${formatNumeric(shape.y - offsetY)}px`,
    `width:${formatNumeric(renderWidth)}px`,
    `height:${formatNumeric(renderHeight)}px`,
    'overflow:visible',
  );

  return [
    `<div style="${styles.join(';')}">`,
    `<svg width="100%" height="100%" viewBox="0 0 ${formatNumeric(renderWidth)} ${formatNumeric(renderHeight)}" preserveAspectRatio="none" aria-hidden="true">`,
    defs.length ? `<defs>${defs.join('')}</defs>` : '',
    `<line ${lineAttributes.join(' ')}/>`,
    '</svg>',
    '</div>',
  ].join('');
}

function renderGradientBorderOverlay(shape: PptxShape): string {
  if (!shape.border?.gradientStops?.length || !shape.border.width) return '';

  const width = Math.max(shape.width || 0, 1);
  const height = Math.max(shape.height || 0, 1);
  const strokeWidthPx = formatNumeric((shape.border.width * 96) / 72);
  const halfStroke = strokeWidthPx / 2;
  const innerWidth = Math.max(width - strokeWidthPx, 0);
  const innerHeight = Math.max(height - strokeWidthPx, 0);
  const paint = buildSvgBorderPaint(`${shape.id}-stroke`, shape.border, width, height);

  if (!paint) return '';

  const defs = paint.defs ? `<defs>${paint.defs}</defs>` : '';
  const strokeAttributes = [
    'fill="none"',
    `stroke="${escHtml(paint.value)}"`,
    `stroke-width="${formatNumeric(strokeWidthPx)}"`,
    'vector-effect="non-scaling-stroke"',
  ];

  if (shape.border.style === 'dashed') strokeAttributes.push('stroke-dasharray="6 4"');
  if (shape.border.style === 'dotted') strokeAttributes.push('stroke-dasharray="1 4"', 'stroke-linecap="round"');

  let borderMarkup = '';
  if (shape.type === 'ellipse') {
    borderMarkup = `<ellipse cx="${formatNumeric(width / 2)}" cy="${formatNumeric(height / 2)}" rx="${formatNumeric(innerWidth / 2)}" ry="${formatNumeric(innerHeight / 2)}" ${strokeAttributes.join(' ')}/>`;
  } else {
    const rectAttributes = [
      `x="${formatNumeric(halfStroke)}"`,
      `y="${formatNumeric(halfStroke)}"`,
      `width="${formatNumeric(innerWidth)}"`,
      `height="${formatNumeric(innerHeight)}"`,
    ];

    if (shape.type === 'roundRect' && shape.cornerRadius) {
      const radius = Math.max(shape.cornerRadius - halfStroke, 0);
      rectAttributes.push(`rx="${formatNumeric(radius)}"`, `ry="${formatNumeric(radius)}"`);
    }

    borderMarkup = `<rect ${rectAttributes.join(' ')} ${strokeAttributes.join(' ')}/>`;
  }

  return `<svg style="position:absolute;inset:0;width:100%;height:100%;pointer-events:none" viewBox="0 0 ${formatNumeric(width)} ${formatNumeric(height)}" preserveAspectRatio="none" aria-hidden="true">${defs}${borderMarkup}</svg>`;
}

export function renderShape(
  shape: PptxShape,
  animationStep: number,
  mediaResolver: (path: string) => string = () => '',
): string {
  if (shape.animationGroup > animationStep) return '';

  const isNewlyRevealed = shape.animationGroup === animationStep && animationStep > 0;
  const visibleParagraphs = trimTrailingEmptyParagraphs((shape.paragraphs || []).filter(
    (paragraph) => (paragraph.animationGroup || 0) <= animationStep,
  ));
  const hasVisibleText = visibleParagraphs.some(paragraphHasVisibleText);
  const textFitScale = shape.textFitScale && shape.textFitScale > 0 ? shape.textFitScale : 1;
  const innerHtml = hasVisibleText ? visibleParagraphs.map((paragraph) => renderParagraph(paragraph, textFitScale)).join('') : '';
  const wrappedTextHtml = hasVisibleText
    ? `<div class="pptx-text-content" style="width:100%;display:block;flex-shrink:0;transform-origin:top left">${innerHtml}</div>`
    : '';

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

  if (isNewlyRevealed) {
    const effect = shape.animationEffect || 'appear';
    styles.push(`animation:pptx-${effect} 0.5s ease both`);
  }

  if (shape.type === 'freeform') {
    styles.push('overflow:visible');
    return renderFreeformShape(shape, styles, innerHtml);
  }

  if (shape.type !== 'image' && shape.type !== 'line') {
    const bg = fillToCss(shape.fill, mediaResolver);
    if (bg !== 'transparent') styles.push(`background:${bg}`);
  }

  const gradientBorderOverlay =
    shape.type !== 'image' && shape.type !== 'line'
      ? renderGradientBorderOverlay(shape)
      : '';

  if (shape.border?.colour && shape.type !== 'line') {
    styles.push(`border:${shape.border.width}pt ${shape.border.style || 'solid'} ${shape.border.colour}`);
  }

  if (shape.type === 'roundRect' && shape.cornerRadius) styles.push(`border-radius:${shape.cornerRadius}px`);
  if (shape.type === 'ellipse') styles.push('border-radius:50%');

  if (hasVisibleText) {
    styles.push('display:flex', 'flex-direction:column');
    const verticalAlign = shape.verticalAlign || 'top';
    if (verticalAlign === 'middle') styles.push('justify-content:center');
    else if (verticalAlign === 'bottom') styles.push('justify-content:flex-end');
    else styles.push('justify-content:flex-start');
    styles.push(buildTextPadding(shape));
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
    return renderLineShape(shape, styles);
  }

  const containerAttributes = [`style="${styles.join(';')}"`];
  if (hasVisibleText) {
    containerAttributes.push('class="pptx-text-shape"');
  }

  return `<div ${containerAttributes.join(' ')}>${gradientBorderOverlay}${wrappedTextHtml}</div>`;
}

export function buildSlideState(
  slide: PptxSlideData,
  animationStep: number,
  mediaResolver: (path: string) => string = () => '',
): SlideState {
  return {
    width: slide.width || 960,
    height: slide.height || 540,
    backgroundCss: slide.background ? fillToCss(slide.background, mediaResolver) : '#ffffff',
    shapesHtml: (slide.shapes || []).map((shape) => renderShape(shape, animationStep, mediaResolver)).join('\n'),
  };
}

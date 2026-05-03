// src/animationCss.ts
var ANIMATION_KEYFRAMES = `
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

// src/imageHelpers.ts
function clampCrop(value) {
  const n = Number(value) || 0;
  return Math.min(0.9999, Math.max(0, n));
}
function getCroppedImageStyles(shape) {
  const leftCrop = clampCrop(shape.imageCrop?.left);
  const topCrop = clampCrop(shape.imageCrop?.top);
  const rightCrop = clampCrop(shape.imageCrop?.right);
  const bottomCrop = clampCrop(shape.imageCrop?.bottom);
  const visibleWidth = Math.max(1e-4, 1 - leftCrop - rightCrop);
  const visibleHeight = Math.max(1e-4, 1 - topCrop - bottomCrop);
  const transforms = [];
  if (shape.flipH) transforms.push("scaleX(-1)");
  if (shape.flipV) transforms.push("scaleY(-1)");
  const styles = [
    "position:absolute",
    `left:${-leftCrop / visibleWidth * 100}%`,
    `top:${-topCrop / visibleHeight * 100}%`,
    `width:${1 / visibleWidth * 100}%`,
    `height:${1 / visibleHeight * 100}%`,
    "display:block"
  ];
  if (transforms.length) {
    styles.push(`transform:${transforms.join(" ")}`);
    styles.push("transform-origin:center center");
  }
  return styles.join(";");
}

// src/renderCore.ts
function escHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function formatNumeric(value) {
  return Number(value.toFixed(3));
}
function hexToRgb(colour) {
  const normalized = colour.replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(normalized)) return void 0;
  return [
    Number.parseInt(normalized.slice(0, 2), 16),
    Number.parseInt(normalized.slice(2, 4), 16),
    Number.parseInt(normalized.slice(4, 6), 16)
  ];
}
function applyOpacity(colour, opacity) {
  if (opacity === void 0 || opacity >= 1) return colour;
  const rgb = hexToRgb(colour);
  if (!rgb) return colour;
  return `rgba(${rgb[0]}, ${rgb[1]}, ${rgb[2]}, ${formatNumeric(opacity)})`;
}
function sortGradientStops(stops) {
  return [...stops].sort((left, right) => left.position - right.position);
}
function gradientStopsToCss(stops) {
  return sortGradientStops(stops).map((stop) => `${applyOpacity(stop.colour, stop.opacity)} ${formatNumeric(stop.position)}%`).join(", ");
}
function gradientAngle(value, fallback = 180) {
  return typeof value?.gradientAngle === "number" ? formatNumeric(value.gradientAngle) : fallback;
}
function gradientVector(angle, width, height) {
  const radians = angle * Math.PI / 180;
  const dx = Math.sin(radians);
  const dy = -Math.cos(radians);
  const halfWidth = width / 2;
  const halfHeight = height / 2;
  const scaleX = Math.abs(dx) > 1e-6 ? halfWidth / Math.abs(dx) : Number.POSITIVE_INFINITY;
  const scaleY = Math.abs(dy) > 1e-6 ? halfHeight / Math.abs(dy) : Number.POSITIVE_INFINITY;
  const extent = Math.min(scaleX, scaleY);
  const centerX = halfWidth;
  const centerY = halfHeight;
  return {
    x1: formatNumeric(centerX - dx * extent),
    y1: formatNumeric(centerY - dy * extent),
    x2: formatNumeric(centerX + dx * extent),
    y2: formatNumeric(centerY + dy * extent)
  };
}
function buildSvgGradient(id, stops, angle, width, height) {
  const { x1, y1, x2, y2 } = gradientVector(angle, width, height);
  const stopMarkup = sortGradientStops(stops).map((stop) => {
    const opacity = stop.opacity !== void 0 && stop.opacity < 1 ? ` stop-opacity="${formatNumeric(stop.opacity)}"` : "";
    return `<stop offset="${formatNumeric(stop.position)}%" stop-color="${escHtml(stop.colour)}"${opacity}/>`;
  }).join("");
  return `<linearGradient id="${escHtml(id)}" gradientUnits="userSpaceOnUse" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stopMarkup}</linearGradient>`;
}
function buildSvgPaint(id, fill, width, height) {
  if (!fill || fill.type === "none") return void 0;
  if (fill.type === "solid" && fill.colour) return { value: fill.colour };
  if (fill.type === "gradient" && fill.gradientStops?.length) {
    return {
      value: `url(#${id})`,
      defs: buildSvgGradient(id, fill.gradientStops, gradientAngle(fill), width, height)
    };
  }
  return void 0;
}
function buildSvgBorderPaint(id, border, width, height) {
  if (!border) return void 0;
  if (border.gradientStops?.length) {
    return {
      value: `url(#${id})`,
      defs: buildSvgGradient(id, border.gradientStops, gradientAngle(border), width, height)
    };
  }
  if (border.colour) return { value: border.colour };
  return void 0;
}
function paragraphHasVisibleText(paragraph) {
  return paragraph.runs.some((run) => run.text !== "" && run.text !== "\n");
}
function trimTrailingEmptyParagraphs(paragraphs) {
  let lastVisibleIndex = paragraphs.length - 1;
  while (lastVisibleIndex >= 0 && !paragraphHasVisibleText(paragraphs[lastVisibleIndex])) {
    lastVisibleIndex -= 1;
  }
  return lastVisibleIndex >= 0 ? paragraphs.slice(0, lastVisibleIndex + 1) : [];
}
function buildTextPadding(shape) {
  const top = shape.textInsets?.top ?? 4;
  const right = shape.textInsets?.right ?? 8;
  const bottom = shape.textInsets?.bottom ?? 4;
  const left = shape.textInsets?.left ?? 8;
  return `padding:${formatNumeric(top)}px ${formatNumeric(right)}px ${formatNumeric(bottom)}px ${formatNumeric(left)}px`;
}
function buildTextOverlayStyles(shape) {
  const styles = ["position:absolute", "inset:0", "display:flex", "flex-direction:column"];
  const verticalAlign = shape.verticalAlign || "top";
  if (verticalAlign === "middle") styles.push("justify-content:center");
  else if (verticalAlign === "bottom") styles.push("justify-content:flex-end");
  else styles.push("justify-content:flex-start");
  styles.push(buildTextPadding(shape));
  return styles.join(";");
}
function fillToCss(fill, mediaResolver = () => "") {
  if (!fill || fill.type === "none") return "transparent";
  if (fill.type === "solid" && fill.colour) return fill.colour;
  if (fill.type === "image" && fill.imageRelativePath) {
    const url = mediaResolver(fill.imageRelativePath);
    return url ? `url("${escHtml(url)}") center/cover no-repeat` : "#cccccc";
  }
  if (fill.type === "gradient" && fill.gradientStops?.length) {
    return `linear-gradient(${gradientAngle(fill)}deg, ${gradientStopsToCss(fill.gradientStops)})`;
  }
  return "transparent";
}
function renderTextRun(run, fontScale = 1) {
  if (run.text === "\n") return "<br/>";
  const styles = buildTextRunStyles(run, fontScale);
  const text = escHtml(run.text);
  return styles.length ? `<span style="${styles.join(";")}">${text}</span>` : `<span>${text}</span>`;
}
function fontFamilyToCssStack(fontFamily) {
  const normalized = fontFamily.trim().toLowerCase();
  if (normalized === "gotham book" || normalized === "gotham" || normalized.startsWith("gotham ")) {
    return `${fontFamily},Avenir Next,Avenir,Helvetica Neue,Arial,sans-serif`;
  }
  if (normalized === "segoe script") {
    return `${fontFamily},Bradley Hand,Marker Felt,Snell Roundhand,Apple Chancery,cursive`;
  }
  if (normalized === "segoe print") {
    return `${fontFamily},Bradley Hand,Marker Felt,Comic Sans MS,cursive`;
  }
  return `${fontFamily},sans-serif`;
}
function buildTextRunStyles(run, fontScale = 1) {
  const styles = [];
  if (run.bold) styles.push("font-weight:bold");
  if (run.italic) styles.push("font-style:italic");
  if (run.underline) styles.push("text-decoration:underline");
  if (run.fontSize) styles.push(`font-size:${formatNumeric(run.fontSize * fontScale)}pt`);
  if (run.fontFamily) styles.push(`font-family:${escHtml(fontFamilyToCssStack(run.fontFamily))}`);
  if (run.colour) styles.push(`color:${run.colour}`);
  return styles;
}
function renderParagraph(para, fontScale = 1) {
  const styles = ["margin:0", "padding:0"];
  if (para.alignment) styles.push(`text-align:${para.alignment}`);
  if (para.lineSpacing && para.lineSpacing > 0) styles.push(`line-height:${formatNumeric(para.lineSpacing)}`);
  else styles.push("line-height:1");
  if (para.level && para.level > 0) styles.push(`padding-left:${para.level * 1.5}em`);
  const hasVisibleText = paragraphHasVisibleText(para);
  const runs = para.runs.map((run) => renderTextRun(run, fontScale)).join("");
  const bulletRun = para.runs.find((run) => run.text && run.text !== "\n");
  const bulletStyles = bulletRun ? buildTextRunStyles(bulletRun, fontScale) : [];
  bulletStyles.push("margin-right:0.4em");
  const bullet = para.bulletType === "bullet" && hasVisibleText ? `<span style="${bulletStyles.join(";")}">${escHtml(para.bulletChar || "\u2022")}</span>` : "";
  return `<p style="${styles.join(";")}">${bullet}${hasVisibleText ? runs : "&nbsp;"}</p>`;
}
function renderFreeformShape(shape, styles, innerHtml) {
  if (!shape.svgPath) return `<div style="${styles.join(";")}">${innerHtml}</div>`;
  const viewBoxWidth = shape.svgViewBoxWidth || shape.width || 1;
  const viewBoxHeight = shape.svgViewBoxHeight || shape.height || 1;
  const defs = [];
  const fillPaint = buildSvgPaint(`${shape.id}-fill`, shape.fill, viewBoxWidth, viewBoxHeight);
  const strokePaint = buildSvgBorderPaint(`${shape.id}-stroke`, shape.border, viewBoxWidth, viewBoxHeight);
  if (fillPaint?.defs) defs.push(fillPaint.defs);
  if (strokePaint?.defs) defs.push(strokePaint.defs);
  const pathAttributes = [`d="${escHtml(shape.svgPath)}"`, `fill="${escHtml(fillPaint?.value || "none")}"`];
  if (strokePaint?.value && shape.border?.width) {
    pathAttributes.push(
      `stroke="${escHtml(strokePaint.value)}"`,
      `stroke-width="${formatNumeric(shape.border.width * 96 / 72)}"`,
      'vector-effect="non-scaling-stroke"'
    );
    if (shape.border.style === "dashed") pathAttributes.push('stroke-dasharray="6 4"');
    if (shape.border.style === "dotted") pathAttributes.push('stroke-dasharray="1 4"', 'stroke-linecap="round"');
  }
  const svg = [
    `<svg width="100%" height="100%" viewBox="0 0 ${formatNumeric(viewBoxWidth)} ${formatNumeric(viewBoxHeight)}" preserveAspectRatio="none" aria-hidden="true">`,
    defs.length ? `<defs>${defs.join("")}</defs>` : "",
    `<path ${pathAttributes.join(" ")}/>`,
    "</svg>"
  ].join("");
  const overlay = innerHtml ? `<div style="${buildTextOverlayStyles(shape)}">${innerHtml}</div>` : "";
  return `<div style="${styles.join(";")}">${svg}${overlay}</div>`;
}
function buildLineMarker(id, colour, strokeWidthPx, position) {
  const arrowSize = Math.max(strokeWidthPx * 2.5, 10);
  const halfSize = arrowSize / 2;
  const path2 = position === "start" ? `M ${formatNumeric(arrowSize)} 0 L 0 ${formatNumeric(halfSize)} L ${formatNumeric(arrowSize)} ${formatNumeric(arrowSize)} Z` : `M 0 0 L ${formatNumeric(arrowSize)} ${formatNumeric(halfSize)} L 0 ${formatNumeric(arrowSize)} Z`;
  const refX = position === "start" ? 0 : arrowSize;
  return `<marker id="${escHtml(id)}" markerWidth="${formatNumeric(arrowSize)}" markerHeight="${formatNumeric(arrowSize)}" refX="${formatNumeric(refX)}" refY="${formatNumeric(halfSize)}" orient="auto" markerUnits="userSpaceOnUse"><path d="${path2}" fill="${escHtml(colour)}"/></marker>`;
}
function renderLineShape(shape, styles) {
  const strokeWidthPt = shape.border?.width || 1;
  const strokeWidthPx = strokeWidthPt * 96 / 72;
  const markerSize = shape.lineHead === "triangle" || shape.lineTail === "triangle" ? Math.max(strokeWidthPx * 2.5, 10) : 0;
  const collapsedAxisSize = Math.max(strokeWidthPx, markerSize, 1);
  const lineWidth = Math.max(shape.width, 1);
  const lineHeight = Math.max(shape.height, 1);
  const renderWidth = shape.width === 0 ? collapsedAxisSize : lineWidth;
  const renderHeight = shape.height === 0 ? collapsedAxisSize : lineHeight;
  const offsetX = shape.width === 0 ? renderWidth / 2 : 0;
  const offsetY = shape.height === 0 ? renderHeight / 2 : 0;
  const strokePaint = buildSvgBorderPaint(`${shape.id}-stroke`, shape.border, renderWidth, renderHeight);
  const stroke = strokePaint?.value || shape.fill?.colour || "#000000";
  const defs = [];
  if (strokePaint?.defs) defs.push(strokePaint.defs);
  if (shape.lineHead === "triangle") defs.push(buildLineMarker(`${shape.id}-head`, stroke, strokeWidthPx, "start"));
  if (shape.lineTail === "triangle") defs.push(buildLineMarker(`${shape.id}-tail`, stroke, strokeWidthPx, "end"));
  const x1 = shape.width > 0 ? shape.flipH ? lineWidth : 0 : offsetX;
  const y1 = shape.height > 0 ? shape.flipV ? lineHeight : 0 : offsetY;
  const x2 = shape.width > 0 ? shape.flipH ? 0 : lineWidth : offsetX;
  const y2 = shape.height > 0 ? shape.flipV ? 0 : lineHeight : offsetY;
  const lineAttributes = [
    `x1="${formatNumeric(x1)}"`,
    `y1="${formatNumeric(y1)}"`,
    `x2="${formatNumeric(x2)}"`,
    `y2="${formatNumeric(y2)}"`,
    `stroke="${escHtml(stroke)}"`,
    `stroke-width="${formatNumeric(strokeWidthPx)}"`,
    'fill="none"'
  ];
  if (shape.border?.style === "dashed") lineAttributes.push('stroke-dasharray="6 4"');
  if (shape.border?.style === "dotted") lineAttributes.push('stroke-dasharray="1 4"', 'stroke-linecap="round"');
  if (shape.lineHead === "triangle") lineAttributes.push(`marker-start="url(#${shape.id}-head)"`);
  if (shape.lineTail === "triangle") lineAttributes.push(`marker-end="url(#${shape.id}-tail)"`);
  styles.push(
    `left:${formatNumeric(shape.x - offsetX)}px`,
    `top:${formatNumeric(shape.y - offsetY)}px`,
    `width:${formatNumeric(renderWidth)}px`,
    `height:${formatNumeric(renderHeight)}px`,
    "overflow:visible"
  );
  return [
    `<div style="${styles.join(";")}">`,
    `<svg width="100%" height="100%" viewBox="0 0 ${formatNumeric(renderWidth)} ${formatNumeric(renderHeight)}" preserveAspectRatio="none" aria-hidden="true">`,
    defs.length ? `<defs>${defs.join("")}</defs>` : "",
    `<line ${lineAttributes.join(" ")}/>`,
    "</svg>",
    "</div>"
  ].join("");
}
function renderGradientBorderOverlay(shape) {
  if (!shape.border?.gradientStops?.length || !shape.border.width) return "";
  const width = Math.max(shape.width || 0, 1);
  const height = Math.max(shape.height || 0, 1);
  const strokeWidthPx = formatNumeric(shape.border.width * 96 / 72);
  const halfStroke = strokeWidthPx / 2;
  const innerWidth = Math.max(width - strokeWidthPx, 0);
  const innerHeight = Math.max(height - strokeWidthPx, 0);
  const paint = buildSvgBorderPaint(`${shape.id}-stroke`, shape.border, width, height);
  if (!paint) return "";
  const defs = paint.defs ? `<defs>${paint.defs}</defs>` : "";
  const strokeAttributes = [
    'fill="none"',
    `stroke="${escHtml(paint.value)}"`,
    `stroke-width="${formatNumeric(strokeWidthPx)}"`,
    'vector-effect="non-scaling-stroke"'
  ];
  if (shape.border.style === "dashed") strokeAttributes.push('stroke-dasharray="6 4"');
  if (shape.border.style === "dotted") strokeAttributes.push('stroke-dasharray="1 4"', 'stroke-linecap="round"');
  let borderMarkup = "";
  if (shape.type === "ellipse") {
    borderMarkup = `<ellipse cx="${formatNumeric(width / 2)}" cy="${formatNumeric(height / 2)}" rx="${formatNumeric(innerWidth / 2)}" ry="${formatNumeric(innerHeight / 2)}" ${strokeAttributes.join(" ")}/>`;
  } else {
    const rectAttributes = [
      `x="${formatNumeric(halfStroke)}"`,
      `y="${formatNumeric(halfStroke)}"`,
      `width="${formatNumeric(innerWidth)}"`,
      `height="${formatNumeric(innerHeight)}"`
    ];
    if (shape.type === "roundRect" && shape.cornerRadius) {
      const radius = Math.max(shape.cornerRadius - halfStroke, 0);
      rectAttributes.push(`rx="${formatNumeric(radius)}"`, `ry="${formatNumeric(radius)}"`);
    }
    borderMarkup = `<rect ${rectAttributes.join(" ")} ${strokeAttributes.join(" ")}/>`;
  }
  return `<svg style="position:absolute;inset:0;width:100%;height:100%;pointer-events:none" viewBox="0 0 ${formatNumeric(width)} ${formatNumeric(height)}" preserveAspectRatio="none" aria-hidden="true">${defs}${borderMarkup}</svg>`;
}
function renderShape(shape, animationStep, mediaResolver = () => "") {
  if (shape.animationGroup > animationStep) return "";
  const isNewlyRevealed = shape.animationGroup === animationStep && animationStep > 0;
  const visibleParagraphs = trimTrailingEmptyParagraphs((shape.paragraphs || []).filter(
    (paragraph) => (paragraph.animationGroup || 0) <= animationStep
  ));
  const hasVisibleText = visibleParagraphs.some(paragraphHasVisibleText);
  const textFitScale = shape.textFitScale && shape.textFitScale > 0 ? shape.textFitScale : 1;
  const innerHtml = hasVisibleText ? visibleParagraphs.map((paragraph) => renderParagraph(paragraph, textFitScale)).join("") : "";
  const wrappedTextHtml = hasVisibleText ? `<div class="pptx-text-content" style="width:100%;display:block;flex-shrink:0;transform-origin:top left">${innerHtml}</div>` : "";
  const styles = [
    "position:absolute",
    "box-sizing:border-box",
    `left:${shape.x}px`,
    `top:${shape.y}px`,
    `width:${shape.width}px`,
    `height:${shape.height}px`,
    "overflow:hidden",
    "word-wrap:break-word"
  ];
  if (shape.rotation) styles.push(`transform:rotate(${shape.rotation}deg)`);
  if (isNewlyRevealed) {
    const effect = shape.animationEffect || "appear";
    styles.push(`animation:pptx-${effect} 0.5s ease both`);
  }
  if (shape.type === "freeform") {
    styles.push("overflow:visible");
    return renderFreeformShape(shape, styles, innerHtml);
  }
  if (shape.type !== "image" && shape.type !== "line") {
    const bg = fillToCss(shape.fill, mediaResolver);
    if (bg !== "transparent") styles.push(`background:${bg}`);
  }
  const gradientBorderOverlay = shape.type !== "image" && shape.type !== "line" ? renderGradientBorderOverlay(shape) : "";
  if (shape.border?.colour && shape.type !== "line") {
    styles.push(`border:${shape.border.width}pt ${shape.border.style || "solid"} ${shape.border.colour}`);
  }
  if (shape.type === "roundRect" && shape.cornerRadius) styles.push(`border-radius:${shape.cornerRadius}px`);
  if (shape.type === "ellipse") styles.push("border-radius:50%");
  if (hasVisibleText) {
    styles.push("display:flex", "flex-direction:column");
    const verticalAlign = shape.verticalAlign || "top";
    if (verticalAlign === "middle") styles.push("justify-content:center");
    else if (verticalAlign === "bottom") styles.push("justify-content:flex-end");
    else styles.push("justify-content:flex-start");
    styles.push(buildTextPadding(shape));
  }
  if (shape.type === "image" && shape.imageRelativePath) {
    const url = mediaResolver(shape.imageRelativePath);
    const imageStyles = shape.imageCrop || shape.flipH || shape.flipV ? getCroppedImageStyles(shape) : "width:100%;height:100%;object-fit:contain;display:block";
    return `<div style="${styles.join(";")}"><img src="${escHtml(url)}" alt="" style="${imageStyles}"/></div>`;
  }
  if (shape.type === "line") {
    return renderLineShape(shape, styles);
  }
  const containerAttributes = [`style="${styles.join(";")}"`];
  if (hasVisibleText) {
    containerAttributes.push('class="pptx-text-shape"');
  }
  return `<div ${containerAttributes.join(" ")}>${gradientBorderOverlay}${wrappedTextHtml}</div>`;
}
function buildSlideState(slide, animationStep, mediaResolver = () => "") {
  return {
    width: slide.width || 960,
    height: slide.height || 540,
    backgroundCss: slide.background ? fillToCss(slide.background, mediaResolver) : "#ffffff",
    shapesHtml: (slide.shapes || []).map((shape) => renderShape(shape, animationStep, mediaResolver)).join("\n")
  };
}

// src/parser.ts
import AdmZip from "adm-zip";
import { XMLParser } from "fast-xml-parser";
import path from "path";
import fs from "fs/promises";
import { writeFileSync } from "fs";
var EMU_PER_PT = 12700;
var EMU_PER_PX = 9525;
function emuToPx(emu) {
  return Math.round((Number(emu) || 0) / EMU_PER_PX);
}
function emuToPt(emu) {
  return Math.round((Number(emu) || 0) / EMU_PER_PT * 10) / 10;
}
function hundredthPtToPt(val) {
  return Math.round((Number(val) || 0) / 100 * 10) / 10;
}
function pickDefined(...values) {
  for (const value of values) {
    if (value !== void 0 && value !== null) return value;
  }
  return void 0;
}
function cloneValue(value) {
  if (value === void 0) return void 0;
  return JSON.parse(JSON.stringify(value));
}
function asArray(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}
function hasAnyValue(obj) {
  return Boolean(obj) && Object.values(obj).some((value) => value !== void 0);
}
function createXmlParser(options = {}) {
  if (options.preserveOrder) {
    return new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      allowBooleanAttributes: true,
      parseAttributeValue: false,
      parseTagValue: false,
      trimValues: false,
      preserveOrder: true
    });
  }
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    allowBooleanAttributes: true,
    parseAttributeValue: false,
    // keep as strings to avoid precision loss
    parseTagValue: false,
    trimValues: false,
    isArray: (name) => {
      const alwaysArray = /* @__PURE__ */ new Set([
        "p:sp",
        "p:pic",
        "p:grpSp",
        "p:graphicFrame",
        "p:cxnSp",
        "a:p",
        "a:r",
        "a:br",
        "a:gs",
        "Relationship",
        "a:buNone",
        "a:buChar",
        "a:buAutoNum",
        "p:childTnLst",
        "p:par",
        "p:seq",
        "p:anim",
        "p:animEffect",
        "p:set",
        "p:cTn",
        "p:stCondLst",
        "p:cond",
        "p:sldId",
        "mc:AlternateContent",
        "mc:Choice",
        "mc:Fallback"
      ]);
      return alwaysArray.has(name);
    }
  });
}
function findOrderedChildNode(children, name) {
  if (!Array.isArray(children)) return void 0;
  return children.find((child) => child && typeof child === "object" && Object.prototype.hasOwnProperty.call(child, name));
}
function findOrderedChild(children, name) {
  return findOrderedChildNode(children, name)?.[name];
}
function getOrderedDrawableId(children, nvName) {
  const nvChildren = findOrderedChild(children, nvName);
  const cNvPrNode = nvChildren ? findOrderedChildNode(nvChildren, "p:cNvPr") : void 0;
  const id = cNvPrNode?.[":@"]?.["@_id"];
  return id !== void 0 ? String(id) : void 0;
}
function collectOrderedShapeMetadata(nodes, state = { textBodyMap: /* @__PURE__ */ new Map(), drawableIds: [] }) {
  if (!Array.isArray(nodes)) return map;
  for (const node of nodes) {
    if (!node || typeof node !== "object") continue;
    const shapeChildren = node["p:sp"] || node["p:cxnSp"];
    if (shapeChildren) {
      const id = node["p:sp"] ? getOrderedDrawableId(shapeChildren, "p:nvSpPr") : getOrderedDrawableId(shapeChildren, "p:nvCxnSpPr");
      const txBody = findOrderedChild(shapeChildren, "p:txBody");
      if (id) state.drawableIds.push(id);
      if (id && txBody) state.textBodyMap.set(id, txBody);
      collectOrderedShapeMetadata(shapeChildren, state);
      continue;
    }
    const picChildren = node["p:pic"];
    if (picChildren) {
      const id = getOrderedDrawableId(picChildren, "p:nvPicPr");
      if (id) state.drawableIds.push(id);
      collectOrderedShapeMetadata(picChildren, state);
      continue;
    }
    const groupChildren = node["p:grpSp"];
    if (groupChildren) {
      collectOrderedShapeMetadata(groupChildren, state);
      continue;
    }
    const alternateContentChildren = node["mc:AlternateContent"] || node.AlternateContent || node["mc:Choice"] || node.Choice || node["mc:Fallback"] || node.Fallback;
    if (alternateContentChildren) {
      collectOrderedShapeMetadata(alternateContentChildren, state);
      continue;
    }
    for (const value of Object.values(node)) {
      if (Array.isArray(value)) collectOrderedShapeMetadata(value, state);
    }
  }
  return state;
}
function parseOrderedShapeMetadata(xml, orderedParser) {
  if (!xml) return { orderedTextBodyMap: void 0, orderedDrawableIds: void 0 };
  try {
    const { textBodyMap, drawableIds } = collectOrderedShapeMetadata(orderedParser.parse(xml));
    return {
      orderedTextBodyMap: textBodyMap.size ? textBodyMap : void 0,
      orderedDrawableIds: drawableIds.length ? drawableIds : void 0
    };
  } catch {
    return { orderedTextBodyMap: void 0, orderedDrawableIds: void 0 };
  }
}
function sortShapesByDrawOrder(shapes, orderedDrawableIds) {
  if (!Array.isArray(shapes) || !Array.isArray(orderedDrawableIds) || !orderedDrawableIds.length) {
    return shapes;
  }
  const drawOrder = /* @__PURE__ */ new Map();
  orderedDrawableIds.forEach((id, index) => {
    if (!drawOrder.has(id)) drawOrder.set(id, index);
  });
  return shapes.map((shape, index) => ({
    shape,
    index,
    order: drawOrder.get(String(shape?.sourceShapeId || ""))
  })).sort((left, right) => {
    const leftKnown = left.order !== void 0;
    const rightKnown = right.order !== void 0;
    if (leftKnown && rightKnown) return left.order - right.order || left.index - right.index;
    if (leftKnown) return -1;
    if (rightKnown) return 1;
    return left.index - right.index;
  }).map(({ shape }) => shape);
}
function getOrderedParagraphNodes(orderedTxBody) {
  if (!Array.isArray(orderedTxBody)) return [];
  return orderedTxBody.filter((child) => child && typeof child === "object" && child["a:p"]).map((child) => child["a:p"]);
}
function normColour(val) {
  if (!val) return void 0;
  const s = String(val).replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(s)) return `#${s}`;
  return void 0;
}
var SCHEME_COLOUR_MAP = {
  dk1: "tx1",
  lt1: "bg1",
  dk2: "tx2",
  lt2: "bg2",
  tx1: "dk1",
  bg1: "lt1",
  tx2: "dk2",
  bg2: "lt2"
};
function resolveSchemeColour(schemeKey, themeColours) {
  if (!themeColours) return void 0;
  const mapped = SCHEME_COLOUR_MAP[schemeKey] || schemeKey;
  return themeColours[mapped] || themeColours[schemeKey] || void 0;
}
function clampColourChannel(value) {
  return Math.max(0, Math.min(255, Math.round(value)));
}
function applyColourTransforms(colour, colourNode) {
  const normalized = normColour(colour);
  if (!normalized || !colourNode) return normalized || colour;
  let channels = [
    Number.parseInt(normalized.slice(1, 3), 16),
    Number.parseInt(normalized.slice(3, 5), 16),
    Number.parseInt(normalized.slice(5, 7), 16)
  ];
  const tint = Number(colourNode["a:tint"]?.["@_val"]);
  if (Number.isFinite(tint)) {
    const amount = tint / 1e5;
    channels = channels.map((channel) => channel + (255 - channel) * amount);
  }
  const shade = Number(colourNode["a:shade"]?.["@_val"]);
  if (Number.isFinite(shade)) {
    const amount = shade / 1e5;
    channels = channels.map((channel) => channel * amount);
  }
  const lumMod = Number(colourNode["a:lumMod"]?.["@_val"]);
  if (Number.isFinite(lumMod)) {
    const amount = lumMod / 1e5;
    channels = channels.map((channel) => channel * amount);
  }
  const lumOff = Number(colourNode["a:lumOff"]?.["@_val"]);
  if (Number.isFinite(lumOff)) {
    const amount = lumOff / 1e5;
    channels = channels.map((channel) => channel + 255 * amount);
  }
  return `#${channels.map((channel) => clampColourChannel(channel).toString(16).padStart(2, "0").toUpperCase()).join("")}`;
}
function parseLineDashStyle(ln) {
  const prstDash = ln?.["a:prstDash"]?.["@_val"];
  if (prstDash === "dash" || prstDash === "lgDash") return "dashed";
  if (prstDash === "dot" || prstDash === "sysDot") return "dotted";
  return "solid";
}
function parseLineEndType(ln, key) {
  const type = ln?.[key]?.["@_type"];
  return type && type !== "none" ? String(type) : void 0;
}
function parseTheme(xmlStr, parser) {
  if (!xmlStr) return { colours: {}, defaultFont: void 0, lineStyles: [] };
  const doc = parser.parse(xmlStr);
  const theme = doc?.["a:theme"];
  const elements = theme?.["a:themeElements"];
  const colours = {};
  let defaultFont;
  const clrScheme = elements?.["a:clrScheme"];
  if (clrScheme) {
    for (const [key, val] of Object.entries(clrScheme)) {
      if (key.startsWith("@_")) continue;
      const srgb = val?.["a:srgbClr"]?.["@_val"];
      const sysClr = val?.["a:sysClr"]?.["@_lastClr"] || val?.["a:sysClr"]?.["@_val"];
      const hex = srgb || sysClr;
      if (hex) {
        colours[key.replace("a:", "")] = `#${hex}`;
      }
    }
  }
  const fontScheme = elements?.["a:fontScheme"];
  const majorFont = fontScheme?.["a:majorFont"]?.["a:latin"]?.["@_typeface"];
  const minorFont = fontScheme?.["a:minorFont"]?.["a:latin"]?.["@_typeface"];
  defaultFont = minorFont || majorFont || void 0;
  const lineStyles = asArray(elements?.["a:fmtScheme"]?.["a:lnStyleLst"]?.["a:ln"]).map((ln) => ({
    width: ln?.["@_w"] ? emuToPt(ln["@_w"]) : void 0,
    style: parseLineDashStyle(ln)
  }));
  return { colours, defaultFont, lineStyles };
}
function parseRels(xmlStr, parser) {
  if (!xmlStr) return {};
  const doc = parser.parse(xmlStr);
  const rels = {};
  const relationships = doc?.Relationships?.Relationship;
  if (!relationships) return rels;
  const list = Array.isArray(relationships) ? relationships : [relationships];
  for (const r of list) {
    if (r?.["@_Id"] && r?.["@_Target"]) {
      rels[r["@_Id"]] = {
        target: r["@_Target"],
        type: r["@_Type"] || ""
      };
    }
  }
  return rels;
}
function parsePresentationXml(xmlStr, parser) {
  const doc = parser.parse(xmlStr);
  const pres = doc?.["p:presentation"];
  const sldSz = pres?.["p:sldSz"];
  const width = emuToPx(sldSz?.["@_cx"]);
  const height = emuToPx(sldSz?.["@_cy"]);
  const sldIdLst = pres?.["p:sldIdLst"]?.["p:sldId"];
  const slideRIds = [];
  if (sldIdLst) {
    const list = Array.isArray(sldIdLst) ? sldIdLst : [sldIdLst];
    for (const s of list) {
      const rId = s?.["@_r:id"];
      if (rId) slideRIds.push(rId);
    }
  }
  return { width: width || 960, height: height || 540, slideRIds };
}
function parseFill(spPr, slideRels, themeColours) {
  if (!spPr) return void 0;
  const solidFill = spPr["a:solidFill"];
  if (solidFill) {
    const colour = extractColour(solidFill, themeColours);
    if (colour) return { type: "solid", colour };
  }
  const blipFill = spPr["a:blipFill"];
  if (blipFill) {
    const embed = extractBlipEmbed(blipFill?.["a:blip"]);
    if (embed && slideRels[embed]) {
      return { type: "image", imageTarget: slideRels[embed].target };
    }
  }
  if (spPr["a:noFill"] !== void 0) return { type: "none" };
  const gradFill = spPr["a:gradFill"];
  if (gradFill) {
    const stops = parseGradientStops(gradFill, themeColours);
    if (stops.length) {
      return {
        type: "gradient",
        gradientStops: stops,
        gradientAngle: parseGradientAngle(gradFill)
      };
    }
  }
  return void 0;
}
function parseGradientAngle(gradFill) {
  const rawAngle = gradFill?.["a:lin"]?.["@_ang"];
  return rawAngle !== void 0 ? ((Number(rawAngle) / 6e4 + 90) % 360 + 360) % 360 : void 0;
}
function extractColourStop(node, themeColours) {
  if (!node) return void 0;
  const srgb = node["a:srgbClr"];
  if (srgb) {
    return {
      colour: applyColourTransforms(normColour(srgb["@_val"]), srgb),
      opacity: srgb["a:alpha"]?.["@_val"] !== void 0 ? Number(srgb["a:alpha"]["@_val"]) / 1e5 : void 0
    };
  }
  const schm = node["a:schemeClr"];
  if (schm) {
    return {
      colour: applyColourTransforms(resolveSchemeColour(schm["@_val"], themeColours), schm),
      opacity: schm["a:alpha"]?.["@_val"] !== void 0 ? Number(schm["a:alpha"]["@_val"]) / 1e5 : void 0
    };
  }
  return void 0;
}
function parseGradientStops(gradFill, themeColours) {
  const gsLst = gradFill?.["a:gsLst"]?.["a:gs"];
  if (!gsLst) return [];
  const list = Array.isArray(gsLst) ? gsLst : [gsLst];
  return list.map((gs) => {
    const colourStop = extractColourStop(gs, themeColours);
    if (!colourStop?.colour) return void 0;
    return {
      position: Number(gs["@_pos"] || 0) / 1e3,
      colour: colourStop.colour,
      opacity: colourStop.opacity
    };
  }).filter(Boolean);
}
function extractColour(node, themeColours) {
  return extractColourStop(node, themeColours)?.colour;
}
function parseBorder(spPr, themeColours, shapeStyle, themeLineStyles) {
  const ln = spPr?.["a:ln"];
  if (!ln) return void 0;
  if (ln["a:noFill"] !== void 0) return void 0;
  const lineStyleIndex = Number(shapeStyle?.["a:lnRef"]?.["@_idx"] || 0);
  const inheritedLineStyle = Number.isFinite(lineStyleIndex) && lineStyleIndex > 0 ? themeLineStyles?.[lineStyleIndex - 1] : void 0;
  const widthEmu = Number(ln["@_w"] || 0);
  const widthPt = widthEmu ? emuToPt(widthEmu) : inheritedLineStyle?.width;
  if (!widthPt) return void 0;
  const style = ln["a:prstDash"]?.["@_val"] ? parseLineDashStyle(ln) : inheritedLineStyle?.style || "solid";
  const solidFill = ln["a:solidFill"];
  if (solidFill) {
    const colour = extractColour(solidFill, themeColours);
    if (colour) return { width: widthPt, colour, style };
  }
  const gradientStops = parseGradientStops(ln["a:gradFill"], themeColours);
  if (gradientStops.length) {
    return {
      width: widthPt,
      style,
      gradientStops,
      gradientAngle: parseGradientAngle(ln["a:gradFill"])
    };
  }
  return void 0;
}
function parseRunStyleDefaults(rPr, themeColours, defaultFont, options = {}) {
  const runDefaults = {};
  if (!rPr) return runDefaults;
  const normaliseMissingEmphasis = options.normaliseMissingEmphasis === true;
  if (rPr["@_b"] !== void 0) {
    runDefaults.bold = rPr["@_b"] === "1" || rPr["@_b"] === "true";
  } else if (normaliseMissingEmphasis) {
    runDefaults.bold = false;
  }
  if (rPr["@_i"] !== void 0) {
    runDefaults.italic = rPr["@_i"] === "1" || rPr["@_i"] === "true";
  } else if (normaliseMissingEmphasis) {
    runDefaults.italic = false;
  }
  if (rPr["@_u"] !== void 0) {
    runDefaults.underline = rPr["@_u"] !== "none";
  } else if (normaliseMissingEmphasis) {
    runDefaults.underline = false;
  }
  const sz = rPr["@_sz"];
  if (sz) runDefaults.fontSize = hundredthPtToPt(sz);
  const typeface = rPr["a:latin"]?.["@_typeface"];
  if (typeface) runDefaults.fontFamily = typeface;
  const solidFill = rPr["a:solidFill"];
  if (solidFill) {
    runDefaults.colour = extractColour(solidFill, themeColours);
  }
  return runDefaults;
}
function mergeRunDefaults(base, override) {
  if (!base && !override) return void 0;
  const merged = {
    bold: pickDefined(override?.bold, base?.bold),
    italic: pickDefined(override?.italic, base?.italic),
    underline: pickDefined(override?.underline, base?.underline),
    fontSize: pickDefined(override?.fontSize, base?.fontSize),
    fontFamily: pickDefined(override?.fontFamily, base?.fontFamily),
    colour: pickDefined(override?.colour, base?.colour)
  };
  return hasAnyValue(merged) ? merged : void 0;
}
function applyRunDefaults(run, defaults, defaultFont) {
  const merged = { ...run };
  if (defaults) {
    for (const key of ["bold", "italic", "underline", "fontSize", "fontFamily", "colour"]) {
      if (merged[key] === void 0 && defaults[key] !== void 0) {
        merged[key] = defaults[key];
      }
    }
  }
  if (merged.fontFamily === void 0 && defaultFont) {
    merged.fontFamily = defaultFont;
  }
  return merged;
}
function applyDefaultFontToParagraphs(paragraphs, defaultFont) {
  if (!paragraphs?.length || !defaultFont) return paragraphs;
  return paragraphs.map((paragraph) => ({
    ...paragraph,
    runs: (paragraph.runs || []).map(
      (run) => run.fontFamily === void 0 ? { ...run, fontFamily: defaultFont } : run
    )
  }));
}
function parseParagraphDefaults(pPrNode, themeColours, defaultFont, level) {
  if (!pPrNode) return void 0;
  const defaults = {};
  if (level !== void 0) defaults.level = level;
  const algn = pPrNode["@_algn"];
  if (algn === "ctr") defaults.alignment = "center";
  else if (algn === "r") defaults.alignment = "right";
  else if (algn === "just") defaults.alignment = "justify";
  else if (algn === "l") defaults.alignment = "left";
  if (pPrNode["a:buChar"]) {
    defaults.bulletType = "bullet";
    defaults.bulletChar = pPrNode["a:buChar"]["@_char"] || void 0;
  } else if (pPrNode["a:buAutoNum"]) {
    defaults.bulletType = "numbered";
  } else if (pPrNode["a:buNone"] !== void 0) {
    defaults.bulletType = "none";
  }
  const lineSpacingPercent = Number(pPrNode["a:lnSpc"]?.["a:spcPct"]?.["@_val"]);
  if (Number.isFinite(lineSpacingPercent) && lineSpacingPercent > 0) {
    defaults.lineSpacing = lineSpacingPercent / 1e5;
  }
  const runDefaults = mergeRunDefaults(
    parseRunStyleDefaults(pPrNode["a:endParaRPr"], themeColours, defaultFont),
    parseRunStyleDefaults(pPrNode["a:defRPr"], themeColours, defaultFont, {
      normaliseMissingEmphasis: true
    })
  );
  if (runDefaults) defaults.runDefaults = runDefaults;
  return hasAnyValue(defaults) ? defaults : void 0;
}
function mergeParagraphDefaults(base, override) {
  if (!base && !override) return void 0;
  const merged = {
    alignment: pickDefined(override?.alignment, base?.alignment),
    bulletType: pickDefined(override?.bulletType, base?.bulletType),
    bulletChar: pickDefined(override?.bulletChar, base?.bulletChar),
    lineSpacing: pickDefined(override?.lineSpacing, base?.lineSpacing),
    level: pickDefined(override?.level, base?.level),
    runDefaults: mergeRunDefaults(base?.runDefaults, override?.runDefaults)
  };
  return hasAnyValue(merged) ? merged : void 0;
}
function parseListStyleDefaults(styleNode, themeColours, defaultFont) {
  const levels = {};
  if (!styleNode) return levels;
  const defaultLevel = parseParagraphDefaults(styleNode["a:defPPr"], themeColours, defaultFont, 0);
  if (defaultLevel) levels[0] = defaultLevel;
  for (let level = 0; level < 9; level++) {
    const node = styleNode[`a:lvl${level + 1}pPr`];
    const parsed = parseParagraphDefaults(node, themeColours, defaultFont, level);
    if (parsed) {
      levels[level] = mergeParagraphDefaults(levels[level], parsed);
    }
  }
  return levels;
}
function mergeTextStyleMaps(base, override) {
  const merged = {};
  for (let level = 0; level < 9; level++) {
    const mergedLevel = mergeParagraphDefaults(base?.[level], override?.[level]);
    if (mergedLevel) merged[level] = mergedLevel;
  }
  return merged;
}
function resolveParagraphDefaults(textStyleMap, level) {
  return textStyleMap?.[level] || textStyleMap?.[0];
}
function parseMasterTextStyles(txStyles, themeColours, defaultFont) {
  return {
    title: parseListStyleDefaults(txStyles?.["p:titleStyle"], themeColours, defaultFont),
    body: parseListStyleDefaults(txStyles?.["p:bodyStyle"], themeColours, defaultFont),
    other: parseListStyleDefaults(txStyles?.["p:otherStyle"], themeColours, defaultFont)
  };
}
function getPlaceholderInfo(node) {
  const ph = node?.["p:nvSpPr"]?.["p:nvPr"]?.["p:ph"] || node?.["p:nvPicPr"]?.["p:nvPr"]?.["p:ph"] || node?.["p:nvCxnSpPr"]?.["p:nvPr"]?.["p:ph"];
  if (!ph) return void 0;
  return {
    type: ph["@_type"] || "body",
    idx: ph["@_idx"] != null ? String(ph["@_idx"]) : void 0,
    orient: ph["@_orient"] || void 0,
    size: ph["@_sz"] || void 0
  };
}
function areCompatiblePlaceholderTypes(sourceType, candidateType) {
  if (sourceType === candidateType) return true;
  const titleTypes = /* @__PURE__ */ new Set(["title", "ctrTitle"]);
  if (titleTypes.has(sourceType) && titleTypes.has(candidateType)) return true;
  const compatibleBodyTypes = /* @__PURE__ */ new Set(["body", "obj", "subTitle"]);
  return compatibleBodyTypes.has(sourceType) && compatibleBodyTypes.has(candidateType);
}
function scorePlaceholderMatch(source, candidate) {
  if (!source || !candidate) return -1;
  const looseIndexTypes = /* @__PURE__ */ new Set(["dt", "ftr", "hdr", "sldNum"]);
  const allowsLooseIndexMatch = looseIndexTypes.has(source.type) && looseIndexTypes.has(candidate.type);
  let score = 0;
  if (source.idx !== void 0 || candidate.idx !== void 0) {
    if (source.idx === candidate.idx) score += 12;
    else if (source.idx !== void 0 && candidate.idx !== void 0 && !allowsLooseIndexMatch) return -1;
  }
  if (source.type !== void 0 || candidate.type !== void 0) {
    if (source.type === candidate.type) {
      score += 8;
    } else if (source.type && candidate.type) {
      if (areCompatiblePlaceholderTypes(source.type, candidate.type)) {
        score += 3;
      } else {
        return -1;
      }
    }
  }
  if (source.orient !== void 0 || candidate.orient !== void 0) {
    if (source.orient === candidate.orient) score += 2;
    else if (source.orient !== void 0 && candidate.orient !== void 0) return -1;
  }
  if (source.size !== void 0 || candidate.size !== void 0) {
    if (source.size === candidate.size) score += 1;
    else if (source.size !== void 0 && candidate.size !== void 0) return -1;
  }
  return score;
}
function findPlaceholderTemplate(shape, templates) {
  if (!shape?.placeholder || !templates?.length) return void 0;
  let bestMatch;
  let bestScore = -1;
  for (const template of templates) {
    const score = scorePlaceholderMatch(shape.placeholder, template?.placeholder);
    if (score > bestScore) {
      bestScore = score;
      bestMatch = template;
    }
  }
  return bestScore >= 0 ? bestMatch : void 0;
}
function getTextStyleMapForPlaceholder(placeholder, masterTextStyles) {
  if (!masterTextStyles || !placeholder) return void 0;
  switch (placeholder.type) {
    case "title":
    case "ctrTitle":
      return masterTextStyles.title;
    case "body":
    case "obj":
    case "subTitle":
      return masterTextStyles.body;
    default:
      return masterTextStyles.other;
  }
}
function getTemplateRunDefaults(paragraphs, index) {
  if (!paragraphs?.length) return void 0;
  const templateParagraph = paragraphs[index] || paragraphs[paragraphs.length - 1];
  if (!templateParagraph?.runs?.length) return void 0;
  const candidate = templateParagraph.runs.find((run) => run.text !== "\n") || templateParagraph.runs[0];
  if (!candidate) return void 0;
  const { text: _text, ...defaults } = candidate;
  return hasAnyValue(defaults) ? defaults : void 0;
}
function mergeParagraphsFromTemplate(paragraphs, templateParagraphs, masterTextStyles, placeholder, defaultFont) {
  if (!paragraphs?.length) return paragraphs;
  const masterStyleMap = getTextStyleMapForPlaceholder(placeholder, masterTextStyles);
  return paragraphs.map((paragraph, index) => {
    const templateParagraph = templateParagraphs?.[index] || templateParagraphs?.[templateParagraphs.length - 1];
    const paragraphDefaults = mergeParagraphDefaults(
      resolveParagraphDefaults(masterStyleMap, paragraph.level || 0),
      templateParagraph ? {
        alignment: templateParagraph.alignment,
        bulletType: templateParagraph.bulletType,
        bulletChar: templateParagraph.bulletChar,
        lineSpacing: templateParagraph.lineSpacing,
        level: templateParagraph.level,
        runDefaults: getTemplateRunDefaults(templateParagraphs, index)
      } : void 0
    );
    return {
      ...paragraph,
      alignment: pickDefined(paragraph.alignment, paragraphDefaults?.alignment),
      bulletType: pickDefined(paragraph.bulletType, paragraphDefaults?.bulletType),
      bulletChar: pickDefined(paragraph.bulletChar, paragraphDefaults?.bulletChar),
      lineSpacing: pickDefined(paragraph.lineSpacing, paragraphDefaults?.lineSpacing),
      level: pickDefined(paragraph.level, paragraphDefaults?.level, 0),
      runs: (paragraph.runs || []).map((run) => applyRunDefaults(run, paragraphDefaults?.runDefaults))
    };
  });
}
function appendTextRun(runs, node, paragraphRunDefaults, themeColours, defaultFont) {
  if (!node) return;
  const rPr = node["a:rPr"];
  const text = node["a:t"] != null ? String(node["a:t"]) : "";
  const runDefaults = mergeRunDefaults(paragraphRunDefaults, parseRunStyleDefaults(rPr, themeColours, defaultFont));
  runs.push(applyRunDefaults({ text }, runDefaults));
}
function parseTextBody(txBody, themeColours, defaultFont, inheritedTextStyles, orderedTxBody) {
  if (!txBody) return [];
  const paragraphs = [];
  const pList = txBody["a:p"];
  if (!pList) return [];
  const pArray = Array.isArray(pList) ? pList : [pList];
  const localTextStyles = parseListStyleDefaults(txBody["a:lstStyle"], themeColours, defaultFont);
  const textStyleMap = mergeTextStyleMaps(inheritedTextStyles, localTextStyles);
  const orderedParagraphs = getOrderedParagraphNodes(orderedTxBody);
  for (let paragraphIndex = 0; paragraphIndex < pArray.length; paragraphIndex++) {
    const p = pArray[paragraphIndex];
    const orderedParagraph = orderedParagraphs[paragraphIndex];
    const pPr = p["a:pPr"];
    const level = Number(pPr?.["@_lvl"] || 0);
    const paragraphDefaults = mergeParagraphDefaults(
      resolveParagraphDefaults(textStyleMap, level),
      parseParagraphDefaults(pPr, themeColours, defaultFont, level)
    );
    let alignment = paragraphDefaults?.alignment;
    let bulletType = paragraphDefaults?.bulletType;
    let bulletChar = paragraphDefaults?.bulletChar;
    let lineSpacing = paragraphDefaults?.lineSpacing;
    const runs = [];
    const rArray = asArray(p["a:r"]);
    const fldArray = asArray(p["a:fld"]);
    const brArray = asArray(p["a:br"]);
    let runIndex = 0;
    let fieldIndex = 0;
    let breakIndex = 0;
    const paragraphRunDefaults = mergeRunDefaults(
      paragraphDefaults?.runDefaults,
      parseRunStyleDefaults(pPr?.["a:defRPr"], themeColours, defaultFont)
    );
    const emptyParagraphRunDefaults = mergeRunDefaults(
      paragraphRunDefaults,
      parseRunStyleDefaults(p["a:endParaRPr"], themeColours, defaultFont)
    );
    if (orderedParagraph?.length) {
      for (const child of orderedParagraph) {
        if (child["a:r"]) {
          appendTextRun(runs, rArray[runIndex], paragraphRunDefaults, themeColours, defaultFont);
          runIndex += 1;
          continue;
        }
        if (child["a:fld"]) {
          appendTextRun(runs, fldArray[fieldIndex], paragraphRunDefaults, themeColours, defaultFont);
          fieldIndex += 1;
          continue;
        }
        if (child["a:br"]) {
          runs.push({ text: "\n" });
          breakIndex += 1;
        }
      }
    }
    while (runIndex < rArray.length) {
      appendTextRun(runs, rArray[runIndex], paragraphRunDefaults, themeColours, defaultFont);
      runIndex += 1;
    }
    while (fieldIndex < fldArray.length) {
      appendTextRun(runs, fldArray[fieldIndex], paragraphRunDefaults, themeColours, defaultFont);
      fieldIndex += 1;
    }
    while (breakIndex < brArray.length) {
      runs.push({ text: "\n" });
      breakIndex += 1;
    }
    if (runs.length === 0) {
      runs.push(applyRunDefaults({ text: "" }, emptyParagraphRunDefaults));
    }
    paragraphs.push({ runs, alignment, bulletType, bulletChar, lineSpacing, level, animationGroup: 0 });
  }
  return paragraphs;
}
function parseShapeTransformRaw(spNode) {
  const xfrm = spNode?.["p:spPr"]?.["a:xfrm"] || spNode?.["p:grpSpPr"]?.["a:xfrm"] || spNode?.["p:spPr"]?.["a:off"]?.[".."] || // shouldn't happen but defensive
  null;
  if (!xfrm) {
    const off2 = spNode?.["p:spPr"]?.["a:off"];
    const ext2 = spNode?.["p:spPr"]?.["a:ext"];
    if (off2 || ext2) {
      return {
        x: Number(off2?.["@_x"] || 0),
        y: Number(off2?.["@_y"] || 0),
        width: Number(ext2?.["@_cx"] || 0),
        height: Number(ext2?.["@_cy"] || 0),
        rotation: 0
      };
    }
    return { x: 0, y: 0, width: 0, height: 0, rotation: 0 };
  }
  const off = xfrm["a:off"];
  const ext = xfrm["a:ext"];
  const rot = Number(xfrm["@_rot"] || 0) / 6e4;
  return {
    x: Number(off?.["@_x"] || 0),
    y: Number(off?.["@_y"] || 0),
    width: Number(ext?.["@_cx"] || 0),
    height: Number(ext?.["@_cy"] || 0),
    rotation: rot
  };
}
function transformToPx(transform) {
  return {
    x: emuToPx(transform.x),
    y: emuToPx(transform.y),
    width: emuToPx(transform.width),
    height: emuToPx(transform.height),
    rotation: transform.rotation || 0
  };
}
function applyGroupTransform(transform, groupContext) {
  if (!groupContext) return transform;
  const scaleX = groupContext.childWidth ? groupContext.width / groupContext.childWidth : 1;
  const scaleY = groupContext.childHeight ? groupContext.height / groupContext.childHeight : 1;
  return {
    x: groupContext.x + (transform.x - groupContext.childX) * scaleX,
    y: groupContext.y + (transform.y - groupContext.childY) * scaleY,
    width: transform.width * scaleX,
    height: transform.height * scaleY,
    rotation: (transform.rotation || 0) + (groupContext.rotation || 0)
  };
}
function parseGroupContext(grpNode, parentGroupContext) {
  const xfrm = grpNode?.["p:grpSpPr"]?.["a:xfrm"];
  if (!xfrm) return parentGroupContext;
  const ownContext = {
    x: Number(xfrm?.["a:off"]?.["@_x"] || 0),
    y: Number(xfrm?.["a:off"]?.["@_y"] || 0),
    width: Number(xfrm?.["a:ext"]?.["@_cx"] || 0),
    height: Number(xfrm?.["a:ext"]?.["@_cy"] || 0),
    childX: Number(xfrm?.["a:chOff"]?.["@_x"] || 0),
    childY: Number(xfrm?.["a:chOff"]?.["@_y"] || 0),
    childWidth: Number(xfrm?.["a:chExt"]?.["@_cx"] || xfrm?.["a:ext"]?.["@_cx"] || 0),
    childHeight: Number(xfrm?.["a:chExt"]?.["@_cy"] || xfrm?.["a:ext"]?.["@_cy"] || 0),
    rotation: Number(xfrm?.["@_rot"] || 0) / 6e4
  };
  if (!parentGroupContext) {
    return ownContext;
  }
  const transformedBounds = applyGroupTransform(ownContext, parentGroupContext);
  return {
    ...ownContext,
    x: transformedBounds.x,
    y: transformedBounds.y,
    width: transformedBounds.width,
    height: transformedBounds.height,
    rotation: transformedBounds.rotation
  };
}
function getGroupNodeId(grpNode) {
  return String(grpNode?.["p:nvGrpSpPr"]?.["p:cNvPr"]?.["@_id"] || "");
}
function getPresetGeometry(spPr) {
  if (spPr?.["a:custGeom"]) return "freeform";
  const prst = spPr?.["a:prstGeom"]?.["@_prst"];
  if (!prst) return "rect";
  if (prst === "ellipse") return "ellipse";
  if (prst === "roundRect") return "roundRect";
  if (prst === "line" || prst.startsWith("straightConnector")) return "line";
  return "rect";
}
function parseFlipFlags(xfrm) {
  if (!xfrm) return { flipH: false, flipV: false };
  return {
    flipH: xfrm["@_flipH"] === "1" || xfrm["@_flipH"] === "true",
    flipV: xfrm["@_flipV"] === "1" || xfrm["@_flipV"] === "true"
  };
}
function parseImageCrop(blipFill) {
  const srcRect = blipFill?.["a:srcRect"];
  if (!srcRect) return void 0;
  const crop = {
    left: Number(srcRect["@_l"] || 0) / 1e5,
    top: Number(srcRect["@_t"] || 0) / 1e5,
    right: Number(srcRect["@_r"] || 0) / 1e5,
    bottom: Number(srcRect["@_b"] || 0) / 1e5
  };
  return Object.values(crop).some((value) => value > 0) ? crop : void 0;
}
function parseTextFitScale(bodyPr) {
  const fontScale = Number(bodyPr?.["a:normAutofit"]?.["@_fontScale"]);
  if (!Number.isFinite(fontScale) || fontScale <= 0) return void 0;
  const scale = fontScale / 1e5;
  return scale > 0 && scale < 1 ? scale : void 0;
}
function parseTextInsets(bodyPr) {
  if (!bodyPr) return void 0;
  const mapping = [
    ["@_lIns", "left"],
    ["@_tIns", "top"],
    ["@_rIns", "right"],
    ["@_bIns", "bottom"]
  ];
  let hasExplicitInsets = false;
  const insets = {};
  for (const [attr, key] of mapping) {
    if (bodyPr[attr] === void 0) continue;
    insets[key] = emuToPx(bodyPr[attr]);
    hasExplicitInsets = true;
  }
  return hasExplicitInsets ? insets : void 0;
}
function extractBlipEmbed(blip) {
  if (!blip) return void 0;
  if (blip["@_r:embed"]) return blip["@_r:embed"];
  for (const ext of asArray(blip["a:extLst"]?.["a:ext"])) {
    const svgBlip = ext?.["asvg:svgBlip"];
    if (svgBlip?.["@_r:embed"]) return svgBlip["@_r:embed"];
  }
  return void 0;
}
function parseCustomGeometry(spPr) {
  const custGeom = spPr?.["a:custGeom"];
  if (!custGeom) return void 0;
  const pathNodes = asArray(custGeom?.["a:pathLst"]?.["a:path"]);
  if (!pathNodes.length) return void 0;
  const segments = [];
  let svgViewBoxWidth = 0;
  let svgViewBoxHeight = 0;
  for (const pathNode of pathNodes) {
    if (!svgViewBoxWidth) svgViewBoxWidth = Number(pathNode?.["@_w"] || 0);
    if (!svgViewBoxHeight) svgViewBoxHeight = Number(pathNode?.["@_h"] || 0);
    for (const moveTo of asArray(pathNode?.["a:moveTo"])) {
      const point = moveTo?.["a:pt"];
      if (point) segments.push(`M ${Number(point["@_x"] || 0)} ${Number(point["@_y"] || 0)}`);
    }
    for (const lineTo of asArray(pathNode?.["a:lnTo"])) {
      const point = lineTo?.["a:pt"];
      if (point) segments.push(`L ${Number(point["@_x"] || 0)} ${Number(point["@_y"] || 0)}`);
    }
    for (const cubicTo of asArray(pathNode?.["a:cubicBezTo"])) {
      const points = asArray(cubicTo?.["a:pt"]);
      if (points.length === 3) {
        segments.push(
          `C ${Number(points[0]["@_x"] || 0)} ${Number(points[0]["@_y"] || 0)} ${Number(points[1]["@_x"] || 0)} ${Number(points[1]["@_y"] || 0)} ${Number(points[2]["@_x"] || 0)} ${Number(points[2]["@_y"] || 0)}`
        );
      }
    }
    for (const quadTo of asArray(pathNode?.["a:quadBezTo"])) {
      const points = asArray(quadTo?.["a:pt"]);
      if (points.length === 2) {
        segments.push(
          `Q ${Number(points[0]["@_x"] || 0)} ${Number(points[0]["@_y"] || 0)} ${Number(points[1]["@_x"] || 0)} ${Number(points[1]["@_y"] || 0)}`
        );
      }
    }
    if (pathNode?.["a:close"] !== void 0) segments.push("Z");
  }
  if (!segments.length) return void 0;
  return {
    svgPath: segments.join(" "),
    svgViewBoxWidth: svgViewBoxWidth || void 0,
    svgViewBoxHeight: svgViewBoxHeight || void 0
  };
}
function parseShape(spNode, slideRels, themeColours, defaultFont, themeLineStyles, groupContext, ancestorGroupIds = [], orderedTextBodyMap) {
  const nvSpPr = spNode["p:nvSpPr"] || spNode["p:nvCxnSpPr"];
  const cNvPr = nvSpPr?.["p:cNvPr"];
  const id = String(cNvPr?.["@_id"] || "");
  const name = cNvPr?.["@_name"] || "";
  const placeholder = getPlaceholderInfo(spNode);
  const spPr = spNode["p:spPr"];
  const transform = transformToPx(applyGroupTransform(parseShapeTransformRaw(spNode), groupContext));
  const customGeometry = parseCustomGeometry(spPr);
  const shapeType = getPresetGeometry(spPr);
  const fill = parseFill(spPr, slideRels, themeColours);
  const border = parseBorder(spPr, themeColours, spNode?.["p:style"], themeLineStyles);
  const lineHead = parseLineEndType(spPr?.["a:ln"], "a:headEnd");
  const lineTail = parseLineEndType(spPr?.["a:ln"], "a:tailEnd");
  const flip = parseFlipFlags(spPr?.["a:xfrm"]);
  const txBody = spNode["p:txBody"];
  const paragraphs = parseTextBody(txBody, themeColours, defaultFont, void 0, orderedTextBodyMap?.get(id));
  let verticalAlign;
  const bodyPr = txBody?.["a:bodyPr"];
  if (bodyPr) {
    const anchor = bodyPr["@_anchor"];
    if (anchor === "ctr" || anchor === "mid") verticalAlign = "middle";
    else if (anchor === "b") verticalAlign = "bottom";
    else if (anchor === "t") verticalAlign = "top";
  }
  const textFitScale = parseTextFitScale(bodyPr);
  const textInsets = parseTextInsets(bodyPr);
  let cornerRadius;
  if (shapeType === "roundRect") {
    cornerRadius = Math.round(Math.min(transform.width, transform.height) * 0.167);
  }
  return {
    id: `shape-${id}`,
    name,
    type: shapeType,
    ...transform,
    fill,
    border,
    paragraphs: paragraphs.length ? paragraphs : void 0,
    verticalAlign,
    textFitScale,
    textInsets,
    cornerRadius,
    lineHead,
    lineTail,
    placeholder,
    sourceShapeId: id,
    ancestorGroupIds,
    flipH: flip.flipH,
    flipV: flip.flipV,
    svgPath: customGeometry?.svgPath,
    svgViewBoxWidth: customGeometry?.svgViewBoxWidth,
    svgViewBoxHeight: customGeometry?.svgViewBoxHeight,
    animationGroup: 0,
    // default: always visible; may be overridden by animation parsing
    animationEffect: "appear"
  };
}
function inheritShapeFromTemplate(shape, template, masterTextStyles, defaultFont) {
  if (!shape) return shape;
  const mergedShape = { ...shape };
  const sourceTemplate = template ? cloneValue(template) : void 0;
  if (sourceTemplate) {
    const needsTemplateTransform = mergedShape.width === 0 && mergedShape.height === 0;
    for (const key of ["x", "y", "width", "height", "rotation"]) {
      if (needsTemplateTransform || mergedShape[key] === void 0) {
        if (sourceTemplate[key] !== void 0) mergedShape[key] = sourceTemplate[key];
      }
    }
    if (mergedShape.fill === void 0 && sourceTemplate.fill !== void 0) mergedShape.fill = sourceTemplate.fill;
    if (mergedShape.border === void 0 && sourceTemplate.border !== void 0) mergedShape.border = sourceTemplate.border;
    if (mergedShape.verticalAlign === void 0 && sourceTemplate.verticalAlign !== void 0) mergedShape.verticalAlign = sourceTemplate.verticalAlign;
    if (mergedShape.textFitScale === void 0 && sourceTemplate.textFitScale !== void 0) mergedShape.textFitScale = sourceTemplate.textFitScale;
    if (mergedShape.textInsets === void 0 && sourceTemplate.textInsets !== void 0) mergedShape.textInsets = sourceTemplate.textInsets;
    if (mergedShape.lineHead === void 0 && sourceTemplate.lineHead !== void 0) mergedShape.lineHead = sourceTemplate.lineHead;
    if (mergedShape.lineTail === void 0 && sourceTemplate.lineTail !== void 0) mergedShape.lineTail = sourceTemplate.lineTail;
    if (mergedShape.cornerRadius === void 0 && sourceTemplate.cornerRadius !== void 0) mergedShape.cornerRadius = sourceTemplate.cornerRadius;
    if (mergedShape.svgPath === void 0 && sourceTemplate.svgPath !== void 0) mergedShape.svgPath = sourceTemplate.svgPath;
    if (mergedShape.svgViewBoxWidth === void 0 && sourceTemplate.svgViewBoxWidth !== void 0) mergedShape.svgViewBoxWidth = sourceTemplate.svgViewBoxWidth;
    if (mergedShape.svgViewBoxHeight === void 0 && sourceTemplate.svgViewBoxHeight !== void 0) mergedShape.svgViewBoxHeight = sourceTemplate.svgViewBoxHeight;
  }
  if (mergedShape.paragraphs?.length) {
    mergedShape.paragraphs = mergeParagraphsFromTemplate(
      mergedShape.paragraphs,
      sourceTemplate?.paragraphs,
      masterTextStyles,
      mergedShape.placeholder || sourceTemplate?.placeholder,
      defaultFont
    );
  }
  return mergedShape;
}
function parsePicture(picNode, slideRels, themeColours, groupContext, ancestorGroupIds = []) {
  const nvPicPr = picNode["p:nvPicPr"];
  const cNvPr = nvPicPr?.["p:cNvPr"];
  const id = String(cNvPr?.["@_id"] || "");
  const name = cNvPr?.["@_name"] || "";
  const blipFill = picNode["p:blipFill"];
  const embed = extractBlipEmbed(blipFill?.["a:blip"]);
  let imageTarget;
  if (embed && slideRels[embed]) {
    imageTarget = slideRels[embed].target;
  }
  const spPr = picNode["p:spPr"];
  const xfrm = spPr?.["a:xfrm"];
  const flip = parseFlipFlags(xfrm);
  const imageCrop = parseImageCrop(blipFill);
  const transform = transformToPx(applyGroupTransform(parseShapeTransformRaw(picNode), groupContext));
  return {
    id: `pic-${id}`,
    name,
    type: "image",
    ...transform,
    imageTarget,
    imageCrop,
    flipH: flip.flipH,
    flipV: flip.flipV,
    sourceShapeId: id,
    ancestorGroupIds,
    animationGroup: 0,
    animationEffect: "appear"
  };
}
function collectDrawableNodes(spTree, slideRels, themeColours, defaultFont, themeLineStyles, groupContext, ancestorGroupIds = [], orderedTextBodyMap) {
  const shapes = [];
  if (!spTree) return shapes;
  const spNodes = spTree["p:sp"];
  if (spNodes) {
    const list = Array.isArray(spNodes) ? spNodes : [spNodes];
    for (const sp of list) {
      shapes.push(parseShape(sp, slideRels, themeColours, defaultFont, themeLineStyles, groupContext, ancestorGroupIds, orderedTextBodyMap));
    }
  }
  const picNodes = spTree["p:pic"];
  if (picNodes) {
    const list = Array.isArray(picNodes) ? picNodes : [picNodes];
    for (const pic of list) {
      shapes.push(parsePicture(pic, slideRels, themeColours, groupContext, ancestorGroupIds));
    }
  }
  const cxnSpNodes = spTree["p:cxnSp"];
  if (cxnSpNodes) {
    const list = Array.isArray(cxnSpNodes) ? cxnSpNodes : [cxnSpNodes];
    for (const cxn of list) {
      shapes.push(parseShape(cxn, slideRels, themeColours, defaultFont, themeLineStyles, groupContext, ancestorGroupIds, orderedTextBodyMap));
    }
  }
  const grpSpNodes = spTree["p:grpSp"];
  if (grpSpNodes) {
    const list = Array.isArray(grpSpNodes) ? grpSpNodes : [grpSpNodes];
    for (const grp of list) {
      const groupId = getGroupNodeId(grp);
      const nextAncestorGroupIds = groupId ? [...ancestorGroupIds, groupId] : ancestorGroupIds;
      const nextGroupContext = parseGroupContext(grp, groupContext);
      shapes.push(
        ...collectDrawableNodes(
          grp,
          slideRels,
          themeColours,
          defaultFont,
          themeLineStyles,
          nextGroupContext,
          nextAncestorGroupIds,
          orderedTextBodyMap
        )
      );
    }
  }
  const alternateContentNodes = [
    ...asArray(spTree["mc:AlternateContent"]),
    ...asArray(spTree.AlternateContent)
  ];
  for (const alternateContent of alternateContentNodes) {
    const fallback = asArray(alternateContent?.["mc:Fallback"])[0] || asArray(alternateContent?.Fallback)[0];
    const choice = asArray(alternateContent?.["mc:Choice"])[0] || asArray(alternateContent?.Choice)[0];
    const drawableContainer = fallback || choice;
    if (drawableContainer) {
      shapes.push(
        ...collectDrawableNodes(
          drawableContainer,
          slideRels,
          themeColours,
          defaultFont,
          themeLineStyles,
          groupContext,
          ancestorGroupIds,
          orderedTextBodyMap
        )
      );
    }
  }
  return shapes;
}
function parseAnimations(timingNode) {
  const animMap = /* @__PURE__ */ new Map();
  if (!timingNode) return animMap;
  const buildGroupMap = /* @__PURE__ */ new Map();
  const paragraphBuildTargets = /* @__PURE__ */ new Set();
  let nextBuildGroup = 1;
  function getBuildGroup(rawGroup) {
    const key = String(rawGroup || nextBuildGroup);
    if (!buildGroupMap.has(key)) {
      buildGroupMap.set(key, nextBuildGroup);
      nextBuildGroup += 1;
    }
    return buildGroupMap.get(key);
  }
  const buildList = timingNode["p:bldLst"];
  const buildNodes = buildList?.["p:bldP"];
  if (buildNodes) {
    const list = Array.isArray(buildNodes) ? buildNodes : [buildNodes];
    for (const build of list) {
      const target = String(build?.["@_spid"] || "");
      if (!target) continue;
      if (build?.["@_build"] === "p") {
        paragraphBuildTargets.add(target);
      }
      animMap.set(target, {
        group: getBuildGroup(build?.["@_grpId"]),
        effect: "appear",
        paragraphGroups: []
      });
    }
  }
  const tnLst = timingNode["p:tnLst"];
  if (!tnLst) return animMap;
  const rootPars = Array.isArray(tnLst["p:par"]) ? tnLst["p:par"] : tnLst["p:par"] ? [tnLst["p:par"]] : [];
  if (!rootPars.length) return animMap;
  for (const rootPar of rootPars) {
    walkForSequences(rootPar, animMap, paragraphBuildTargets);
  }
  return animMap;
}
function firstNode(node) {
  return Array.isArray(node) ? node[0] : node;
}
function walkForSequences(node, animMap, paragraphBuildTargets) {
  if (!node) return;
  const cTn = node["p:cTn"] || (Array.isArray(node) ? null : node);
  if (!cTn) return;
  const cTnObj = Array.isArray(cTn) ? cTn[0] : cTn;
  const childTnLst = firstNode(cTnObj?.["p:childTnLst"]);
  if (!childTnLst) return;
  const seqNodes = childTnLst["p:seq"];
  if (seqNodes) {
    const seqList = Array.isArray(seqNodes) ? seqNodes : [seqNodes];
    for (const seq of seqList) {
      parseClickSequence(seq, animMap, paragraphBuildTargets);
    }
  }
  const parNodes = childTnLst["p:par"];
  if (parNodes) {
    const parList = Array.isArray(parNodes) ? parNodes : [parNodes];
    for (const par of parList) {
      walkForSequences(par, animMap, paragraphBuildTargets);
    }
  }
}
function parseClickSequence(seqNode, animMap, paragraphBuildTargets) {
  if (!seqNode) return;
  const cTn = seqNode["p:cTn"];
  const cTnObj = Array.isArray(cTn) ? cTn[0] : cTn;
  if (cTnObj?.["@_nodeType"] && cTnObj["@_nodeType"] !== "mainSeq") return;
  const childTnLst = firstNode(cTnObj?.["p:childTnLst"]);
  if (!childTnLst) return;
  const parNodes = childTnLst["p:par"];
  if (!parNodes) return;
  const parList = Array.isArray(parNodes) ? parNodes : [parNodes];
  let clickIndex = 0;
  for (const par of parList) {
    const targets = /* @__PURE__ */ new Map();
    collectTargetsFromPar(par, targets);
    if (!targets.size) continue;
    clickIndex++;
    for (const [target, effect] of targets.entries()) {
      registerAnimationTarget(target, clickIndex, effect, animMap, paragraphBuildTargets);
    }
  }
}
function collectTargetsFromPar(node, targets) {
  if (!node) return;
  const cTn = node["p:cTn"];
  const cTnObj = Array.isArray(cTn) ? cTn[0] : cTn;
  const childTnLst = firstNode(cTnObj?.["p:childTnLst"]);
  if (childTnLst) {
    const parNodes = childTnLst["p:par"];
    if (parNodes) {
      const parList = Array.isArray(parNodes) ? parNodes : [parNodes];
      for (const p of parList) {
        collectTargetsFromPar(p, targets);
      }
    }
    const setNodes = childTnLst["p:set"];
    if (setNodes) {
      const list = Array.isArray(setNodes) ? setNodes : [setNodes];
      for (const s of list) {
        const target = extractAnimTarget(s);
        if (target) {
          registerCollectedTarget(target, "appear", targets);
        }
      }
    }
    const animEffectNodes = childTnLst["p:animEffect"];
    if (animEffectNodes) {
      const list = Array.isArray(animEffectNodes) ? animEffectNodes : [animEffectNodes];
      for (const ae of list) {
        const target = extractAnimTarget(ae);
        const transition = ae["@_transition"];
        const filter = ae["@_filter"];
        let effect = "fade";
        if (filter?.includes("wipe")) effect = "fly-left";
        if (transition === "out") effect = "fade";
        if (target) {
          registerCollectedTarget(target, effect, targets);
        }
      }
    }
    const animNodes = childTnLst["p:anim"];
    if (animNodes) {
      const list = Array.isArray(animNodes) ? animNodes : [animNodes];
      for (const a of list) {
        const target = extractAnimTarget(a);
        if (target) {
          registerCollectedTarget(target, "fly-left", targets);
        }
      }
    }
  }
  const directTarget = extractAnimTarget(node);
  if (directTarget) {
    registerCollectedTarget(directTarget, "appear", targets);
  }
}
function registerCollectedTarget(target, effect, targets) {
  if (!target) return;
  if (!targets.has(target)) {
    targets.set(target, effect || "appear");
  }
}
function registerAnimationTarget(target, clickGroup, effect, animMap, paragraphBuildTargets) {
  const existing = animMap.get(target);
  const nextEffect = effect || existing?.effect || "appear";
  if (paragraphBuildTargets.has(target)) {
    const paragraphGroups = Array.isArray(existing?.paragraphGroups) ? [...existing.paragraphGroups] : [];
    if (!paragraphGroups.includes(clickGroup)) {
      paragraphGroups.push(clickGroup);
    }
    animMap.set(target, {
      group: Math.min(existing?.group ?? clickGroup, clickGroup),
      effect: existing?.effect || nextEffect,
      paragraphGroups
    });
    return;
  }
  animMap.set(target, {
    group: clickGroup,
    effect: existing?.effect === "appear" ? nextEffect : existing?.effect || nextEffect,
    paragraphGroups: existing?.paragraphGroups || []
  });
}
function extractAnimTarget(node) {
  const cBhvr = node?.["p:cBhvr"];
  const tgtEl = cBhvr?.["p:tgtEl"];
  const spTgt = tgtEl?.["p:spTgt"];
  if (spTgt) {
    return spTgt["@_spid"];
  }
  const cTn = node?.["p:cTn"];
  const cTnObj = Array.isArray(cTn) ? cTn[0] : cTn;
  const stCondLst = firstNode(cTnObj?.["p:stCondLst"]);
  if (stCondLst) {
    const cond = stCondLst["p:cond"];
    const condList = Array.isArray(cond) ? cond : cond ? [cond] : [];
    for (const c of condList) {
      const sp = c?.["p:tgtEl"]?.["p:spTgt"];
      if (sp) return sp["@_spid"];
    }
  }
  return void 0;
}
function parseSlideBackground(bgNode, slideRels, themeColours) {
  if (!bgNode) return void 0;
  const bgPr = bgNode["p:bgPr"];
  if (bgPr) {
    const fill = parseFill(bgPr, slideRels, themeColours);
    if (fill) return fill;
  }
  const bgRef = bgNode["p:bgRef"];
  if (bgRef) {
    const colour = extractColour(bgRef, themeColours);
    if (colour) return { type: "solid", colour };
  }
  return void 0;
}
function parseLayoutMasterShapes(spTree, rels, themeColours, defaultFont, themeLineStyles, orderedTextBodyMap, orderedDrawableIds) {
  const shapes = [];
  const placeholders = [];
  if (!spTree) return { shapes, placeholders };
  const drawableNodes = sortShapesByDrawOrder(
    collectDrawableNodes(
      spTree,
      rels,
      themeColours,
      defaultFont,
      themeLineStyles,
      void 0,
      [],
      orderedTextBodyMap
    ),
    orderedDrawableIds
  );
  for (const shape of drawableNodes) {
    if (shape.placeholder) {
      placeholders.push(shape);
      continue;
    }
    if (shape.width > 0 || shape.height > 0) {
      shapes.push(shape);
    }
  }
  return { shapes, placeholders };
}
function parseNotes(notesXmlStr, parser) {
  if (!notesXmlStr) return void 0;
  try {
    const doc = parser.parse(notesXmlStr);
    const cSld = doc?.["p:notes"]?.["p:cSld"];
    const spTree = cSld?.["p:spTree"];
    if (!spTree) return void 0;
    const spNodes = spTree["p:sp"];
    if (!spNodes) return void 0;
    const list = Array.isArray(spNodes) ? spNodes : [spNodes];
    const textParts = [];
    for (const sp of list) {
      const phIdx = sp?.["p:nvSpPr"]?.["p:nvPr"]?.["p:ph"]?.["@_type"];
      if (phIdx === "body" || phIdx === "notes" || !phIdx) {
        const txBody = sp["p:txBody"];
        if (!txBody) continue;
        const pList = txBody["a:p"];
        if (!pList) continue;
        const pArr = Array.isArray(pList) ? pList : [pList];
        for (const p of pArr) {
          const rList = p["a:r"];
          if (!rList) continue;
          const rArr = Array.isArray(rList) ? rList : [rList];
          for (const r of rArr) {
            const text = r["a:t"];
            if (text != null) textParts.push(String(text));
          }
        }
      }
    }
    return textParts.length ? textParts.join("") : void 0;
  } catch {
    return void 0;
  }
}
async function parsePptx(filePath, deckId, getDeckDir) {
  const zip = new AdmZip(filePath);
  const parser = createXmlParser();
  const orderedParser = createXmlParser({ preserveOrder: true });
  const presXml = zip.readAsText("ppt/presentation.xml");
  const { width, height, slideRIds } = parsePresentationXml(presXml, parser);
  const presRelsXml = zip.readAsText("ppt/_rels/presentation.xml.rels");
  const presRels = parseRels(presRelsXml, parser);
  let themeData = { colours: {}, defaultFont: void 0, lineStyles: [] };
  try {
    const themeRel = Object.values(presRels).find((r) => r.type.includes("theme"));
    if (themeRel) {
      const themePath = `ppt/${themeRel.target.replace(/^\.\//, "")}`;
      const themeXml = zip.readAsText(themePath);
      themeData = parseTheme(themeXml, parser);
    }
  } catch {
  }
  const slidePaths = [];
  for (const rId of slideRIds) {
    const rel = presRels[rId];
    if (rel) {
      slidePaths.push(rel.target.replace(/^\.\//, ""));
    }
  }
  const slidesOutputDir = path.join(getDeckDir(deckId), "slides");
  await fs.mkdir(slidesOutputDir, { recursive: true });
  const mediaMap = /* @__PURE__ */ new Map();
  let mediaSeq = Date.now();
  async function extractMedia(pptxInternalTarget, slideSubDir) {
    const resolved = path.posix.normalize(path.posix.join(`ppt/${slideSubDir}`, pptxInternalTarget));
    if (mediaMap.has(resolved)) return mediaMap.get(resolved);
    const entry = zip.getEntry(resolved);
    if (!entry) return void 0;
    const ext = path.extname(resolved).toLowerCase() || ".png";
    const outName = `pptx-media-${mediaSeq}${ext}`;
    mediaSeq++;
    const outPath = path.join(slidesOutputDir, outName);
    await fs.writeFile(outPath, entry.getData());
    const relativePath = `slides/${outName}`;
    mediaMap.set(resolved, relativePath);
    return relativePath;
  }
  const slides = [];
  const layoutCache = /* @__PURE__ */ new Map();
  const masterCache = /* @__PURE__ */ new Map();
  function extractMediaFromShapes(shapes, dir) {
    for (const shape of shapes) {
      if (shape.imageTarget) {
        const resolved = path.posix.normalize(path.posix.join(`ppt/${dir}`, shape.imageTarget));
        if (!mediaMap.has(resolved)) {
          const entry = zip.getEntry(resolved);
          if (entry) {
            const ext = path.extname(resolved).toLowerCase() || ".png";
            const outName = `pptx-media-${mediaSeq}${ext}`;
            mediaSeq++;
            const outPath = path.join(slidesOutputDir, outName);
            writeFileSync(outPath, entry.getData());
            mediaMap.set(resolved, `slides/${outName}`);
          }
        }
        shape.imageRelativePath = mediaMap.get(resolved);
        shape.type = "image";
        delete shape.imageTarget;
      }
      if (shape.fill?.imageTarget) {
        const resolved = path.posix.normalize(path.posix.join(`ppt/${dir}`, shape.fill.imageTarget));
        if (!mediaMap.has(resolved)) {
          const entry = zip.getEntry(resolved);
          if (entry) {
            const ext = path.extname(resolved).toLowerCase() || ".png";
            const outName = `pptx-media-${mediaSeq}${ext}`;
            mediaSeq++;
            const outPath = path.join(slidesOutputDir, outName);
            writeFileSync(outPath, entry.getData());
            mediaMap.set(resolved, `slides/${outName}`);
          }
        }
        shape.fill.imageRelativePath = mediaMap.get(resolved);
        shape.fill.type = "image";
        delete shape.fill.imageTarget;
      }
    }
  }
  function extractMediaFromBackground(background, dir) {
    if (background?.imageTarget) {
      const resolved = path.posix.normalize(path.posix.join(`ppt/${dir}`, background.imageTarget));
      if (!mediaMap.has(resolved)) {
        const entry = zip.getEntry(resolved);
        if (entry) {
          const ext = path.extname(resolved).toLowerCase() || ".png";
          const outName = `pptx-media-${mediaSeq}${ext}`;
          mediaSeq++;
          const outPath = path.join(slidesOutputDir, outName);
          writeFileSync(outPath, entry.getData());
          mediaMap.set(resolved, `slides/${outName}`);
        }
      }
      background.imageRelativePath = mediaMap.get(resolved);
      background.type = "image";
      delete background.imageTarget;
    }
  }
  function parseLayoutFile(layoutFullPath, layoutDir, layoutBaseName) {
    if (layoutCache.has(layoutFullPath)) return layoutCache.get(layoutFullPath);
    let result = { background: void 0, shapes: [], placeholders: [], masterPath: void 0, showMasterShapes: true };
    try {
      const xml = zip.readAsText(layoutFullPath);
      const doc = parser.parse(xml);
      const { orderedTextBodyMap, orderedDrawableIds } = parseOrderedShapeMetadata(xml, orderedParser);
      const sldLayout = doc?.["p:sldLayout"];
      const cSld = sldLayout?.["p:cSld"];
      const showMasterShapes = sldLayout?.["@_showMasterSp"] !== "0" && sldLayout?.["@_showMasterSp"] !== "false";
      let layoutRels = {};
      let masterPath;
      try {
        const layoutRelsPath = `ppt/${layoutDir}/_rels/${layoutBaseName}.rels`;
        const layoutRelsXml = zip.readAsText(layoutRelsPath);
        layoutRels = parseRels(layoutRelsXml, parser);
        const masterRel = Object.values(layoutRels).find((r) => r.type.includes("slideMaster"));
        if (masterRel) {
          masterPath = path.posix.normalize(path.posix.join(`ppt/${layoutDir}`, masterRel.target));
        }
      } catch {
      }
      const bg = cSld?.["p:bg"];
      const background = parseSlideBackground(bg, layoutRels, themeData.colours);
      extractMediaFromBackground(background, layoutDir);
      const spTree = cSld?.["p:spTree"];
      const { shapes, placeholders } = parseLayoutMasterShapes(
        spTree,
        layoutRels,
        themeData.colours,
        themeData.defaultFont,
        themeData.lineStyles,
        orderedTextBodyMap,
        orderedDrawableIds
      );
      extractMediaFromShapes(shapes, layoutDir);
      extractMediaFromShapes(placeholders, layoutDir);
      result = { background, shapes, placeholders, masterPath, showMasterShapes };
    } catch {
    }
    layoutCache.set(layoutFullPath, result);
    return result;
  }
  function parseMasterFile(masterFullPath) {
    if (masterCache.has(masterFullPath)) return masterCache.get(masterFullPath);
    let result = { background: void 0, shapes: [], placeholders: [], textStyles: { title: {}, body: {}, other: {} } };
    try {
      const xml = zip.readAsText(masterFullPath);
      const doc = parser.parse(xml);
      const { orderedTextBodyMap, orderedDrawableIds } = parseOrderedShapeMetadata(xml, orderedParser);
      const sldMaster = doc?.["p:sldMaster"];
      const cSld = sldMaster?.["p:cSld"];
      const textStyles = parseMasterTextStyles(sldMaster?.["p:txStyles"], themeData.colours, themeData.defaultFont);
      const masterDir = path.posix.dirname(masterFullPath.replace(/^ppt\//, ""));
      const masterBaseName = path.posix.basename(masterFullPath);
      let masterRels = {};
      try {
        const masterRelsPath = `ppt/${masterDir}/_rels/${masterBaseName}.rels`;
        const masterRelsXml = zip.readAsText(masterRelsPath);
        masterRels = parseRels(masterRelsXml, parser);
      } catch {
      }
      const bg = cSld?.["p:bg"];
      const background = parseSlideBackground(bg, masterRels, themeData.colours);
      extractMediaFromBackground(background, masterDir);
      const spTree = cSld?.["p:spTree"];
      const { shapes, placeholders } = parseLayoutMasterShapes(
        spTree,
        masterRels,
        themeData.colours,
        themeData.defaultFont,
        themeData.lineStyles,
        orderedTextBodyMap,
        orderedDrawableIds
      );
      extractMediaFromShapes(shapes, masterDir);
      extractMediaFromShapes(placeholders, masterDir);
      result = { background, shapes, placeholders, textStyles };
    } catch {
    }
    masterCache.set(masterFullPath, result);
    return result;
  }
  for (let i = 0; i < slidePaths.length; i++) {
    const slidePath = slidePaths[i];
    const fullSlidePath = `ppt/${slidePath}`;
    const slideDir = path.posix.dirname(slidePath);
    const slideBaseName = path.posix.basename(slidePath);
    let slideXml;
    try {
      slideXml = zip.readAsText(fullSlidePath);
    } catch {
      continue;
    }
    const slideDoc = parser.parse(slideXml);
    const { orderedTextBodyMap, orderedDrawableIds } = parseOrderedShapeMetadata(slideXml, orderedParser);
    const sld = slideDoc?.["p:sld"];
    const slideRelsPath = `ppt/${slideDir}/_rels/${slideBaseName}.rels`;
    let slideRels = {};
    try {
      const slideRelsXml = zip.readAsText(slideRelsPath);
      slideRels = parseRels(slideRelsXml, parser);
    } catch {
    }
    let layoutData = { background: void 0, shapes: [], placeholders: [], masterPath: void 0, showMasterShapes: true };
    let masterData = { background: void 0, shapes: [], placeholders: [], textStyles: { title: {}, body: {}, other: {} } };
    const layoutRel = Object.values(slideRels).find((r) => r.type.includes("slideLayout"));
    if (layoutRel) {
      const layoutPath = path.posix.normalize(path.posix.join(`ppt/${slideDir}`, layoutRel.target));
      const layoutDir = path.posix.dirname(layoutPath.replace(/^ppt\//, ""));
      const layoutBaseName = path.posix.basename(layoutPath);
      layoutData = parseLayoutFile(layoutPath, layoutDir, layoutBaseName);
      if (layoutData.masterPath) {
        masterData = parseMasterFile(layoutData.masterPath);
      }
    }
    const cSld = sld?.["p:cSld"];
    const bg = cSld?.["p:bg"];
    let background = parseSlideBackground(bg, slideRels, themeData.colours);
    const useLayoutBg = !background;
    if (useLayoutBg && layoutData.background) {
      background = layoutData.background;
    }
    if (!background && masterData.background) {
      background = masterData.background;
    }
    const spTree = cSld?.["p:spTree"];
    const rawSlideShapes = sortShapesByDrawOrder(
      collectDrawableNodes(
        spTree,
        slideRels,
        themeData.colours,
        themeData.defaultFont,
        themeData.lineStyles,
        void 0,
        [],
        orderedTextBodyMap
      ),
      orderedDrawableIds
    );
    const slideShapes = rawSlideShapes.map((shape) => {
      if (!shape.placeholder) return shape;
      const masterTemplate = findPlaceholderTemplate(shape, masterData.placeholders);
      const layoutTemplate = findPlaceholderTemplate(shape, layoutData.placeholders);
      return inheritShapeFromTemplate(
        inheritShapeFromTemplate(shape, layoutTemplate, masterData.textStyles, themeData.defaultFont),
        masterTemplate,
        masterData.textStyles,
        themeData.defaultFont
      );
    }).filter((shape) => shape.width > 0 || shape.height > 0);
    for (const shape of slideShapes) {
      if (shape.imageTarget) {
        const relPath = await extractMedia(shape.imageTarget, slideDir);
        if (relPath) {
          shape.imageRelativePath = relPath;
        }
        delete shape.imageTarget;
      }
    }
    for (const shape of slideShapes) {
      if (shape.fill?.imageTarget) {
        const relPath = await extractMedia(shape.fill.imageTarget, slideDir);
        if (relPath) {
          shape.fill.imageRelativePath = relPath;
          shape.fill.type = "image";
        }
        delete shape.fill.imageTarget;
      }
    }
    if (background?.imageTarget) {
      const relPath = await extractMedia(background.imageTarget, slideDir);
      if (relPath) {
        background.imageRelativePath = relPath;
        background.type = "image";
      }
      delete background.imageTarget;
    }
    const showMasterSp = sld?.["@_showMasterSp"];
    const slideAllowsMasterShapes = showMasterSp !== "0" && showMasterSp !== "false";
    const showMasterShapes = slideAllowsMasterShapes && layoutData.showMasterShapes !== false;
    const allShapes = [];
    if (showMasterShapes) {
      for (const s of masterData.shapes) {
        allShapes.push({ ...cloneValue(s), id: `master-${s.id}-s${i}` });
      }
    }
    for (const s of layoutData.shapes) {
      allShapes.push({ ...cloneValue(s), id: `layout-${s.id}-s${i}` });
    }
    allShapes.push(...slideShapes);
    for (const shape of allShapes) {
      if (shape.paragraphs?.length) {
        shape.paragraphs = applyDefaultFontToParagraphs(shape.paragraphs, themeData.defaultFont);
      }
    }
    const timing = sld?.["p:timing"];
    const animMap = parseAnimations(timing);
    let maxAnimGroup = 0;
    for (const shape of allShapes) {
      const candidateIds = [];
      if (shape.sourceShapeId) candidateIds.push(String(shape.sourceShapeId));
      if (Array.isArray(shape.ancestorGroupIds)) {
        candidateIds.push(...[...shape.ancestorGroupIds].reverse().map(String));
      }
      const matchedId = candidateIds.find((candidateId) => animMap.has(candidateId));
      if (matchedId) {
        const anim = animMap.get(matchedId);
        shape.animationGroup = anim.group;
        shape.animationEffect = anim.effect;
        if (Array.isArray(anim.paragraphGroups) && anim.paragraphGroups.length && Array.isArray(shape.paragraphs) && shape.paragraphs.length) {
          shape.paragraphs = shape.paragraphs.map((paragraph, index) => ({
            ...paragraph,
            animationGroup: anim.paragraphGroups[Math.min(index, anim.paragraphGroups.length - 1)] ?? anim.group
          }));
          maxAnimGroup = Math.max(maxAnimGroup, ...anim.paragraphGroups);
        } else {
          if (Array.isArray(shape.paragraphs)) {
            shape.paragraphs = shape.paragraphs.map((paragraph) => ({ ...paragraph, animationGroup: anim.group }));
          }
          if (anim.group > maxAnimGroup) maxAnimGroup = anim.group;
        }
      } else if (Array.isArray(shape.paragraphs)) {
        shape.paragraphs = shape.paragraphs.map((paragraph) => ({ ...paragraph, animationGroup: 0 }));
      }
    }
    let notes;
    const notesRel = Object.values(slideRels).find((r) => r.type.includes("notesSlide"));
    if (notesRel) {
      try {
        const notesPath = `ppt/${slideDir}/${notesRel.target.replace(/^\.\//, "")}`;
        const notesXml = zip.readAsText(notesPath);
        notes = parseNotes(notesXml, parser);
      } catch {
      }
    }
    slides.push({
      slideIndex: i,
      width,
      height,
      background,
      shapes: allShapes,
      notes,
      animationStepCount: maxAnimGroup
    });
  }
  return {
    sourceFileName: path.basename(filePath),
    slides,
    theme: themeData
  };
}
export {
  ANIMATION_KEYFRAMES,
  buildSlideState,
  clampCrop,
  escHtml,
  fillToCss,
  getCroppedImageStyles,
  parsePptx,
  renderParagraph,
  renderShape,
  renderTextRun
};
//# sourceMappingURL=index.js.map
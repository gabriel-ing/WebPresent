const path = require('node:path');
const fs = require('node:fs/promises');
const { existsSync } = require('node:fs');
const { readFileAsDataUrl } = require('./utils.cjs');

function escH(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

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

function clampCrop(value) {
  const number = Number(value) || 0;
  return Math.min(0.9999, Math.max(0, number));
}

function getCroppedImageStyles(shape) {
  const crop = shape.imageCrop;
  const leftCrop = clampCrop(crop?.left);
  const topCrop = clampCrop(crop?.top);
  const rightCrop = clampCrop(crop?.right);
  const bottomCrop = clampCrop(crop?.bottom);
  const visibleWidth = Math.max(0.0001, 1 - leftCrop - rightCrop);
  const visibleHeight = Math.max(0.0001, 1 - topCrop - bottomCrop);
  const transforms = [];
  if (shape.flipH) transforms.push('scaleX(-1)');
  if (shape.flipV) transforms.push('scaleY(-1)');

  const styles = [
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

function buildPptxSlideState(slideData, animationStep, mediaUrls) {
  const slide = slideData;

  function resolveMedia(relativePath) {
    return relativePath ? mediaUrls[relativePath] || '' : '';
  }

  function fillCss(fill) {
    if (!fill || fill.type === 'none') return 'transparent';
    if (fill.type === 'solid' && fill.colour) return fill.colour;
    if (fill.type === 'image' && fill.imageRelativePath) {
      const url = resolveMedia(fill.imageRelativePath);
      return url ? `url("${escH(url)}") center/cover no-repeat` : '#ccc';
    }
    if (fill.type === 'gradient' && fill.gradientStops && fill.gradientStops.length) {
      const stops = fill.gradientStops.map((s) => `${s.colour} ${Math.round(s.position)}%`).join(', ');
      return `linear-gradient(180deg, ${stops})`;
    }
    return 'transparent';
  }

  function renderRun(run) {
    if (run.text === '\n') return '<br/>';
    const st = [];
    if (run.bold) st.push('font-weight:bold');
    if (run.italic) st.push('font-style:italic');
    if (run.underline) st.push('text-decoration:underline');
    if (run.fontSize) st.push(`font-size:${run.fontSize}pt`);
    if (run.fontFamily) st.push(`font-family:${escH(run.fontFamily)},sans-serif`);
    if (run.colour) st.push(`color:${run.colour}`);
    const text = escH(run.text);
    return st.length ? `<span style="${st.join(';')}">${text}</span>` : `<span>${text}</span>`;
  }

  function renderPara(p) {
    const st = ['margin:0', 'padding:0 0 0.15em 0'];
    if (p.alignment) st.push(`text-align:${p.alignment}`);
    if (p.level > 0) st.push(`padding-left:${p.level * 1.5}em`);
    const runs = (p.runs || []).map(renderRun).join('');
    const bullet = p.bulletType === 'bullet' ? `<span style="margin-right:0.4em">${escH(p.bulletChar || '•')}</span>` : '';
    return `<p style="${st.join(';')}">${bullet}${runs}</p>`;
  }

  function renderShape(shape) {
    if (shape.animationGroup > animationStep) return '';
    const isNew = shape.animationGroup === animationStep && animationStep > 0;
    const st = [
      'position:absolute', 'box-sizing:border-box',
      `left:${shape.x}px`, `top:${shape.y}px`,
      `width:${shape.width}px`, `height:${shape.height}px`,
      'overflow:hidden', 'word-wrap:break-word',
    ];
    if (shape.rotation) st.push(`transform:rotate(${shape.rotation}deg)`);
    if (shape.type !== 'image' && shape.type !== 'line') {
      const bg = fillCss(shape.fill);
      if (bg !== 'transparent') st.push(`background:${bg}`);
    }
    if (shape.border) st.push(`border:${shape.border.width}pt ${shape.border.style || 'solid'} ${shape.border.colour}`);
    if (shape.type === 'roundRect' && shape.cornerRadius) st.push(`border-radius:${shape.cornerRadius}px`);
    if (shape.type === 'ellipse') st.push('border-radius:50%');
    if (isNew) {
      const eff = shape.animationEffect || 'appear';
      st.push(`animation:pptx-${eff} 0.5s ease both`);
    }
    const visibleParagraphs = (shape.paragraphs || []).filter((paragraph) => (paragraph.animationGroup || 0) <= animationStep);
    if (visibleParagraphs.length) {
      st.push('display:flex', 'flex-direction:column');
      const vAlign = shape.verticalAlign || 'top';
      if (vAlign === 'middle') st.push('justify-content:center');
      else if (vAlign === 'bottom') st.push('justify-content:flex-end');
      else st.push('justify-content:flex-start');
      st.push('padding:4px 8px');
    }
    if (shape.type === 'image' && shape.imageRelativePath) {
      const url = resolveMedia(shape.imageRelativePath);
      const imageStyles = shape.imageCrop || shape.flipH || shape.flipV
        ? getCroppedImageStyles(shape)
        : 'width:100%;height:100%;object-fit:contain;display:block';
      return `<div style="${st.join(';')}"><img src="${escH(url)}" alt="" style="${imageStyles}"/></div>`;
    }
    if (shape.type === 'line') {
      const c = (shape.border && shape.border.colour) || (shape.fill && shape.fill.colour) || '#000';
      const lw = (shape.border && shape.border.width) || 1;
      st.push(`border-top:${lw}pt solid ${c}`, 'height:0');
      return `<div style="${st.join(';')}"></div>`;
    }
    const inner = visibleParagraphs.map(renderPara).join('');
    return `<div style="${st.join(';')}">${inner}</div>`;
  }

  return {
    width: slide.width || 960,
    height: slide.height || 540,
    backgroundCss: slide.background ? fillCss(slide.background) : '#ffffff',
    shapesHtml: (slide.shapes || []).map(renderShape).join('\n'),
  };
}

function buildPptxSlideHtml(slideData, animationStep, mediaUrls) {
  const state = buildPptxSlideState(slideData, animationStep, mediaUrls);
  const initialState = JSON.stringify(state);

  return `<!doctype html><html><head><meta charset="utf-8"/><style>
*{margin:0;padding:0;box-sizing:border-box}
html,body{width:100%;height:100%;overflow:hidden;background:#000}
@keyframes pptx-appear{from{opacity:0}to{opacity:1}}
@keyframes pptx-fade{from{opacity:0}to{opacity:1}}
@keyframes pptx-fly-left{from{opacity:0;transform:translateX(-60px)}to{opacity:1;transform:translateX(0)}}
@keyframes pptx-fly-right{from{opacity:0;transform:translateX(60px)}to{opacity:1;transform:translateX(0)}}
@keyframes pptx-fly-up{from{opacity:0;transform:translateY(-60px)}to{opacity:1;transform:translateY(0)}}
@keyframes pptx-fly-down{from{opacity:0;transform:translateY(60px)}to{opacity:1;transform:translateY(0)}}
.slide-root{position:absolute;left:0;top:0;transform-origin:0 0;overflow:hidden;font-family:Calibri,Arial,Helvetica,sans-serif;font-size:18pt;color:#000}
</style><script>
(function(){function fit(){var r=document.querySelector('.slide-root');if(!r)return;var sw=Number(r.dataset.slideWidth)||1,sh=Number(r.dataset.slideHeight)||1,vw=window.innerWidth||document.documentElement.clientWidth||sw,vh=window.innerHeight||document.documentElement.clientHeight||sh;if(vw<1||vh<1)return;var sc=Math.min(vw/sw,vh/sh);r.style.transform='scale('+sc+')';r.style.left=Math.round((vw-sw*sc)/2)+'px';r.style.top=Math.round((vh-sh*sc)/2)+'px';}window.__WEBPRESENT_UPDATE_PPTX=function(state){var r=document.querySelector('.slide-root');if(!r||!state)return;r.dataset.slideWidth=state.width;r.dataset.slideHeight=state.height;r.style.width=state.width+'px';r.style.height=state.height+'px';r.style.background=state.backgroundCss;r.innerHTML=state.shapesHtml||'';fit();};window.addEventListener('resize',fit);window.addEventListener('load',fit);setTimeout(fit,0);setTimeout(fit,50);setTimeout(fit,200);if(window.ResizeObserver){new ResizeObserver(fit).observe(document.documentElement);}window.__WEBPRESENT_UPDATE_PPTX(${initialState});})()
</script></head><body><div class="slide-root"></div></body></html>`;
}

async function buildPptxPresentationDocument(slideData, animationStep, deckDir) {
  const relativePaths = collectSlideMediaPaths(slideData);
  const mediaUrls = await buildMediaDataUrlMap(deckDir, relativePaths);
  return buildPptxSlideHtml(slideData, animationStep, mediaUrls);
}

async function buildPptxRuntimeUpdateScript(slideData, animationStep, deckDir, options = {}) {
  const relativePaths = collectSlideMediaPaths(slideData);
  const mediaUrls = await buildMediaDataUrlMap(deckDir, relativePaths, options);
  return `window.__WEBPRESENT_UPDATE_PPTX(${JSON.stringify(buildPptxSlideState(slideData, animationStep, mediaUrls))});`;
}

function createPptxPresentationDocumentBuilder(deckDir, options = {}) {
  const mediaCache = new Map();
  const htmlCache = new Map();
  const updateScriptCache = new Map();

  const builder = async function buildCachedPptxPresentationDocument(slideData, animationStep) {
    const cacheKey = `${slideData.slideIndex}:${animationStep}`;
    if (htmlCache.has(cacheKey)) {
      return htmlCache.get(cacheKey);
    }

    const relativePaths = collectSlideMediaPaths(slideData);
    const mediaUrls = await buildMediaDataUrlMap(deckDir, relativePaths, {
      ...options,
      cache: mediaCache,
    });
    const html = buildPptxSlideHtml(slideData, animationStep, mediaUrls);
    htmlCache.set(cacheKey, html);
    return html;
  };

  builder.buildUpdateScript = async function buildCachedPptxRuntimeUpdateScript(slideData, animationStep) {
    const cacheKey = `${slideData.slideIndex}:${animationStep}`;
    if (updateScriptCache.has(cacheKey)) {
      return updateScriptCache.get(cacheKey);
    }

    const relativePaths = collectSlideMediaPaths(slideData);
    const mediaUrls = await buildMediaDataUrlMap(deckDir, relativePaths, {
      ...options,
      cache: mediaCache,
    });
    const script = `window.__WEBPRESENT_UPDATE_PPTX(${JSON.stringify(buildPptxSlideState(slideData, animationStep, mediaUrls))});`;
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

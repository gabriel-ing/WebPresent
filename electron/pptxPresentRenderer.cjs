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
</style></head><body><div class="slide-root"></div><script>
(function(){
  function fitOverflowingText(root){
    if(!root)return;
    var boxes=root.querySelectorAll('.pptx-text-shape');
    boxes.forEach(function(box){
      var content=box.querySelector('.pptx-text-content');
      if(!content)return;
      content.style.transform='';
      content.style.width='100%';
      var styles=window.getComputedStyle(box);
      var availableHeight=box.clientHeight-parseFloat(styles.paddingTop||'0')-parseFloat(styles.paddingBottom||'0');
      if(!(availableHeight>0))return;
      var contentHeight=content.scrollHeight;
      if(contentHeight<=availableHeight+1)return;
      box.style.justifyContent='flex-start';
      var scale=availableHeight/contentHeight;
      if(!(scale>0&&scale<1))return;
      content.style.transform='scale('+scale+')';
      content.style.width=(100/scale)+'%';
    });
  }
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
    fitOverflowingText(r);
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
</script></body></html>`;
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

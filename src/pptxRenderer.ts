import { buildSlideState, ANIMATION_KEYFRAMES } from '@webpresent/pptx-engine';
import type { PptxSlideData } from '@webpresent/pptx-engine';

// ── Slide fit() script ────────────────────────────────────────────────────────

function makeFitScript(width: number, height: number): string {
  return `(function(){
  function fitOverflowingText(){
    var root=document.querySelector('.slide-root');
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
  function updateLayout(){
    fitOverflowingText();
    fit();
  }
  window.addEventListener('resize',updateLayout);
  window.addEventListener('load',updateLayout);
  setTimeout(updateLayout,0);
  setTimeout(updateLayout,50);
  setTimeout(updateLayout,200);
  if(window.ResizeObserver){new ResizeObserver(updateLayout).observe(document.documentElement);}
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

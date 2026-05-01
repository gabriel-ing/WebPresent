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

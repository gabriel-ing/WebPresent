type PptxShapeType = 'rect' | 'ellipse' | 'roundRect' | 'image' | 'line' | 'freeform';
type PptxTextRun = {
    text: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontSize?: number;
    fontFamily?: string;
    colour?: string;
};
type PptxParagraph = {
    runs: PptxTextRun[];
    alignment?: 'left' | 'center' | 'right' | 'justify';
    bulletType?: 'none' | 'bullet' | 'numbered';
    bulletChar?: string;
    level?: number;
    animationGroup?: number;
};
type PptxImageCrop = {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
};
type PptxFill = {
    type: 'solid' | 'image' | 'gradient' | 'none';
    colour?: string;
    imageRelativePath?: string;
    gradientStops?: {
        position: number;
        colour: string;
    }[];
};
type PptxBorder = {
    width: number;
    colour: string;
    style?: 'solid' | 'dashed' | 'dotted';
};
type PptxShape = {
    id: string;
    name?: string;
    type: PptxShapeType;
    x: number;
    y: number;
    width: number;
    height: number;
    rotation?: number;
    fill?: PptxFill;
    border?: PptxBorder;
    paragraphs?: PptxParagraph[];
    imageRelativePath?: string;
    imageCrop?: PptxImageCrop;
    cornerRadius?: number;
    verticalAlign?: 'top' | 'middle' | 'bottom';
    flipH?: boolean;
    flipV?: boolean;
    animationGroup: number;
    animationEffect?: 'appear' | 'fade' | 'fly-left' | 'fly-right' | 'fly-up' | 'fly-down';
};
type PptxSlideData = {
    slideIndex: number;
    width: number;
    height: number;
    background?: PptxFill;
    shapes: PptxShape[];
    notes?: string;
    animationStepCount: number;
};
type PptxDeckData = {
    sourceFileName: string;
    slides: PptxSlideData[];
    theme?: {
        colours: Record<string, string>;
        defaultFont?: string;
    };
};

/**
 * CSS @keyframe declarations for PPTX entrance animation effects.
 *
 * These are embedded in every rendered slide HTML document.
 * Extend this list as additional effect types are supported.
 */
declare const ANIMATION_KEYFRAMES = "\n@keyframes pptx-appear {\n  from { opacity: 0; }\n  to   { opacity: 1; }\n}\n@keyframes pptx-fade {\n  from { opacity: 0; }\n  to   { opacity: 1; }\n}\n@keyframes pptx-fly-left {\n  from { opacity: 0; transform: translateX(-60px); }\n  to   { opacity: 1; transform: translateX(0); }\n}\n@keyframes pptx-fly-right {\n  from { opacity: 0; transform: translateX(60px); }\n  to   { opacity: 1; transform: translateX(0); }\n}\n@keyframes pptx-fly-up {\n  from { opacity: 0; transform: translateY(-60px); }\n  to   { opacity: 1; transform: translateY(0); }\n}\n@keyframes pptx-fly-down {\n  from { opacity: 0; transform: translateY(60px); }\n  to   { opacity: 1; transform: translateY(0); }\n}\n";

/**
 * Clamps a fractional crop value to the valid range [0, 0.9999].
 * Values outside this range would produce invisible or inverted images.
 */
declare function clampCrop(value: number | undefined): number;
/**
 * Produces inline CSS for an image that needs crop and/or flip transforms.
 *
 * The parent container uses `overflow:hidden` and the image is sized/
 * positioned to show only the cropped region. Flip transforms are applied
 * via CSS `scaleX(-1)` / `scaleY(-1)`.
 */
declare function getCroppedImageStyles(shape: PptxShape): string;

/**
 * The minimal serialisable state needed to render or update a slide.
 * Passed as JSON to `window.__WEBPRESENT_UPDATE_PPTX` for in-place updates.
 */
type SlideState = {
    width: number;
    height: number;
    backgroundCss: string;
    shapesHtml: string;
};
/** Escapes a string for safe inline use in HTML text content and attribute values. */
declare function escHtml(s: string): string;
/**
 * Converts a PptxFill to a CSS background value.
 *
 * @param fill - The fill descriptor from the parsed shape or slide background.
 * @param mediaResolver - Called with a relative media path; returns a data URL
 *   or absolute URL. Defaults to returning an empty string (no media).
 */
declare function fillToCss(fill: PptxFill | undefined, mediaResolver?: (path: string) => string): string;
/** Renders a single styled text run to an HTML `<span>`. */
declare function renderTextRun(run: PptxTextRun): string;
/** Renders a single paragraph (already filtered for visibility) to a `<p>`. */
declare function renderParagraph(para: PptxParagraph): string;
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
declare function renderShape(shape: PptxShape, animationStep: number, mediaResolver?: (path: string) => string): string;
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
declare function buildSlideState(slide: PptxSlideData, animationStep: number, mediaResolver?: (path: string) => string): SlideState;

/**
 * PPTX Parser — extracts slide data from .pptx (OOXML) files.
 *
 * Uses adm-zip to read the archive and fast-xml-parser to parse XML.
 * Returns a structured PptxDeckData object that can be rendered as HTML.
 */

/**
 * Parse a .pptx file and extract structured slide data + media files.
 *
 * @param {string} pptxFilePath  Absolute path to the .pptx file.
 * @param {string} deckId        The deck ID to store extracted media in.
 * @param {function} getDeckDir  Function that returns the deck directory for a given deckId.
 * @returns {Promise<object>}    PptxDeckData
 */
declare function parsePptx(filePath: string, deckId: string, getDeckDir: (id: string) => string): Promise<PptxDeckData>;

export { ANIMATION_KEYFRAMES, type PptxBorder, type PptxDeckData, type PptxFill, type PptxImageCrop, type PptxParagraph, type PptxShape, type PptxShapeType, type PptxSlideData, type PptxTextRun, type SlideState, buildSlideState, clampCrop, escHtml, fillToCss, getCroppedImageStyles, parsePptx, renderParagraph, renderShape, renderTextRun };

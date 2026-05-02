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
    lineSpacing?: number;
    level?: number;
    animationGroup?: number;
};
type PptxImageCrop = {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
};
type PptxTextInsets = {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
};
type PptxGradientStop = {
    position: number;
    colour: string;
    opacity?: number;
};
type PptxFill = {
    type: 'solid' | 'image' | 'gradient' | 'none';
    colour?: string;
    imageRelativePath?: string;
    gradientStops?: PptxGradientStop[];
    gradientAngle?: number;
};
type PptxBorder = {
    width: number;
    colour?: string;
    style?: 'solid' | 'dashed' | 'dotted';
    gradientStops?: PptxGradientStop[];
    gradientAngle?: number;
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
    textFitScale?: number;
    textInsets?: PptxTextInsets;
    lineHead?: string;
    lineTail?: string;
    flipH?: boolean;
    flipV?: boolean;
    svgPath?: string;
    svgViewBoxWidth?: number;
    svgViewBoxHeight?: number;
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
        lineStyles?: Array<{
            width?: number;
            style?: 'solid' | 'dashed' | 'dotted';
        }>;
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

type SlideState = {
    width: number;
    height: number;
    backgroundCss: string;
    shapesHtml: string;
};
declare function escHtml(s: string): string;
declare function fillToCss(fill: PptxFill | undefined, mediaResolver?: (path: string) => string): string;
declare function renderTextRun(run: PptxTextRun, fontScale?: number): string;
declare function renderParagraph(para: PptxParagraph, fontScale?: number): string;
declare function renderShape(shape: PptxShape, animationStep: number, mediaResolver?: (path: string) => string): string;
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

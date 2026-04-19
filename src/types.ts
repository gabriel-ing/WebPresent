export type StepType = 'web' | 'slide' | 'pptx-slide';

export type SlideMediaKind = 'image' | 'video';

export type SlideRef = {
  id: string;
  relativePath: string;
  sourceFileName?: string;
  mediaKind?: SlideMediaKind;
};

// ── PPTX Slide Types ─────────────────────────────────────────────────────────

export type PptxShapeType = 'rect' | 'ellipse' | 'roundRect' | 'image' | 'line' | 'freeform';

export type PptxTextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  colour?: string;
};

export type PptxParagraph = {
  runs: PptxTextRun[];
  alignment?: 'left' | 'center' | 'right' | 'justify';
  bulletType?: 'none' | 'bullet' | 'numbered';
  bulletChar?: string;
  level?: number;
  animationGroup?: number;
};

export type PptxImageCrop = {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
};

export type PptxFill = {
  type: 'solid' | 'image' | 'gradient' | 'none';
  colour?: string;
  imageRelativePath?: string;
  gradientStops?: { position: number; colour: string }[];
};

export type PptxBorder = {
  width: number;
  colour: string;
  style?: 'solid' | 'dashed' | 'dotted';
};

export type PptxShape = {
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

export type PptxSlideData = {
  slideIndex: number;
  width: number;
  height: number;
  background?: PptxFill;
  shapes: PptxShape[];
  notes?: string;
  animationStepCount: number;
};

export type PptxDeckData = {
  sourceFileName: string;
  slides: PptxSlideData[];
  theme?: {
    colours: Record<string, string>;
    defaultFont?: string;
  };
};

// ── Presentation Types ───────────────────────────────────────────────────────

export type PresentationStep = {
  id: string;
  type: StepType;
  title?: string;
  notes?: string;
  groupId?: string;
  url?: string;
  webZoom?: number;
  slideRef?: SlideRef;
  pptxSlideData?: PptxSlideData;
  pptxAnimationStep?: number;
};

export type Presentation = {
  id: string;
  title: string;
  createdAt: string;
  updatedAt: string;
  items: PresentationStep[];
};

export type SlideImportMode = 'separate' | 'grouped';

export type DisplayInfo = {
  id: number;
  label: string;
  width: number;
  height: number;
};

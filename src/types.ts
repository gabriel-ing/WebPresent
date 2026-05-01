import type { PptxSlideData } from '@webpresent/pptx-engine';

export type {
  PptxShapeType,
  PptxTextRun,
  PptxParagraph,
  PptxImageCrop,
  PptxFill,
  PptxBorder,
  PptxShape,
  PptxSlideData,
  PptxDeckData,
} from '@webpresent/pptx-engine';

export type StepType = 'web' | 'slide' | 'pptx-slide';

export type SlideMediaKind = 'image' | 'video';

export type SlideRef = {
  id: string;
  relativePath: string;
  sourceFileName?: string;
  mediaKind?: SlideMediaKind;
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

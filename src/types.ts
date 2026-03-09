export type StepType = 'web' | 'slide';

export type SlideRef = {
  id: string;
  relativePath: string;
  sourceFileName?: string;
};

export type PresentationStep = {
  id: string;
  type: StepType;
  title?: string;
  notes?: string;
  groupId?: string;
  url?: string;
  slideRef?: SlideRef;
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

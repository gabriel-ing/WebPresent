import type { DisplayInfo, Presentation, SlideImportMode, SlideRef } from './types';

type StartPresentationParams = {
  deckId: string;
  startIndex: number;
  displayId?: number;
};

type ImportSlidesParams = {
  deckId: string;
  filePaths: string[];
  mode: SlideImportMode;
};

type ResolveSlideParams = {
  deckId: string;
  relativePath: string;
};

export type PresentApi = {
  deckGetCurrent: () => Promise<Presentation>;
  deckCreate: () => Promise<Presentation>;
  deckSave: (presentation: Presentation) => Promise<void>;
  deckImport: () => Promise<Presentation | null>;
  deckExport: (deckId: string) => Promise<void>;
  pickSlideFiles: () => Promise<string[]>;
  pickSlideDirectory: () => Promise<string[]>;
  importSlides: (params: ImportSlidesParams) => Promise<SlideRef[]>;
  resolveSlideUrl: (params: ResolveSlideParams) => Promise<string | null>;
  resolveSlideDataUrl: (params: ResolveSlideParams) => Promise<string | null>;
  startPresentation: (params: StartPresentationParams) => Promise<void>;
  stopPresentation: () => Promise<void>;
  getDisplays: () => Promise<DisplayInfo[]>;
  openExternal: (url: string) => Promise<void>;
};

declare global {
  interface Window {
    presentApi?: PresentApi;
  }
}

export {};

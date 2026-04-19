import type { PresentationStep, SlideRef } from './types';
import { guessTitleFromUrl } from './utils';

const VIDEO_EXTENSIONS = new Set(['.mp4', '.webm', '.mov', '.m4v', '.ogv', '.ogg']);

export function getSlideMediaKind(slideRef?: SlideRef): 'image' | 'video' {
  if (slideRef?.mediaKind) return slideRef.mediaKind;
  const source = slideRef?.relativePath || slideRef?.sourceFileName || '';
  const extension = source.includes('.') ? source.slice(source.lastIndexOf('.')).toLowerCase() : '';
  return VIDEO_EXTENSIONS.has(extension) ? 'video' : 'image';
}

export function getStepTitle(step: PresentationStep): string {
  if (step.title?.trim()) return step.title;
  if (step.type === 'web') {
    return step.url ? guessTitleFromUrl(step.url) : 'Untitled web step';
  }
  if (step.type === 'pptx-slide') {
    const slideNum = step.pptxSlideData ? step.pptxSlideData.slideIndex + 1 : '?';
    const animStep = step.pptxAnimationStep || 0;
    return animStep > 0 ? `Slide ${slideNum} (build ${animStep})` : `Slide ${slideNum}`;
  }
  return step.slideRef?.sourceFileName || step.slideRef?.id || 'Untitled slide step';
}

export function insertAfterSelected(
  items: PresentationStep[],
  selectedId: string | null,
  incoming: PresentationStep[],
): PresentationStep[] {
  if (!selectedId) return [...items, ...incoming];
  const selectedIndex = items.findIndex((item) => item.id === selectedId);
  if (selectedIndex === -1) return [...items, ...incoming];
  return [...items.slice(0, selectedIndex + 1), ...incoming, ...items.slice(selectedIndex + 1)];
}

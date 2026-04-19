import { useEffect, useState } from 'react';
import type { Presentation, PresentationStep } from '../types';

/**
 * Manages a cache of media data URLs for the current presentation.
 *
 * Resolves both plain SlideRef media and PPTX shape/background images
 * using a single unified effect.
 *
 * Cache key format: `"<deckId>|<relativePath>"`
 */
export function useSlideUrls(
  presentation: Presentation | null,
  selectedStep: PresentationStep | null,
): Record<string, string> {
  const [slideUrls, setSlideUrls] = useState<Record<string, string>>({});

  // Clear URL cache when the deck changes
  const presentationId = presentation?.id ?? null;
  useEffect(() => {
    setSlideUrls({});
  }, [presentationId]);

  useEffect(() => {
    if (!presentation || !selectedStep || !window.presentApi) return;

    const deckId = presentation.id;
    const pathsToResolve: string[] = [];

    // Plain slide
    if (selectedStep.slideRef?.relativePath) {
      pathsToResolve.push(selectedStep.slideRef.relativePath);
    }

    // PPTX slide — collect all embedded image paths
    if (selectedStep.pptxSlideData) {
      const slide = selectedStep.pptxSlideData;
      if (slide.background?.imageRelativePath) {
        pathsToResolve.push(slide.background.imageRelativePath);
      }
      for (const shape of slide.shapes) {
        if (shape.imageRelativePath) pathsToResolve.push(shape.imageRelativePath);
        if (shape.fill?.imageRelativePath) pathsToResolve.push(shape.fill.imageRelativePath);
      }
    }

    const unresolved = pathsToResolve.filter((p) => !slideUrls[`${deckId}|${p}`]);
    if (!unresolved.length) return;

    let mounted = true;
    async function resolveAll() {
      const results: Record<string, string> = {};
      for (const relativePath of unresolved) {
        const resolved = await window.presentApi!.resolveSlideDataUrl({ deckId, relativePath });
        if (!mounted) return;
        if (resolved) results[`${deckId}|${relativePath}`] = resolved;
      }
      if (Object.keys(results).length) {
        setSlideUrls((prev) => ({ ...prev, ...results }));
      }
    }

    void resolveAll();
    return () => { mounted = false; };
  }, [presentation, selectedStep]);

  return slideUrls;
}

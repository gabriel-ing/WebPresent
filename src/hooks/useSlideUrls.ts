import { useEffect, useState } from 'react';
import type { Presentation, PresentationStep } from '../types';

/**
 * Resolves preview media for the selected step only.
 *
 * The preview pane uses blob URLs instead of data URLs so large slide assets
 * do not balloon into multiple base64-encoded copies in memory.
 */

function collectStepAssetPaths(selectedStep: PresentationStep | null): string[] {
  if (!selectedStep) return [];

  const paths = new Set<string>();

  if (selectedStep.slideRef?.relativePath) {
    paths.add(selectedStep.slideRef.relativePath);
  }

  if (selectedStep.pptxSlideData) {
    const slide = selectedStep.pptxSlideData;
    if (slide.background?.imageRelativePath) {
      paths.add(slide.background.imageRelativePath);
    }
    for (const shape of slide.shapes) {
      if (shape.imageRelativePath) paths.add(shape.imageRelativePath);
      if (shape.fill?.imageRelativePath) paths.add(shape.fill.imageRelativePath);
    }
  }

  return Array.from(paths);
}

function toUint8Array(buffer: Uint8Array | ArrayBuffer | number[]): Uint8Array {
  if (buffer instanceof Uint8Array) return buffer;
  if (buffer instanceof ArrayBuffer) return new Uint8Array(buffer);
  return Uint8Array.from(buffer);
}

function toBlobPart(buffer: Uint8Array | ArrayBuffer | number[]): ArrayBuffer {
  const bytes = toUint8Array(buffer);
  const copy = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(copy).set(bytes);
  return copy;
}

export function useSlideUrls(
  presentation: Presentation | null,
  selectedStep: PresentationStep | null,
): Record<string, string> {
  const [slideUrls, setSlideUrls] = useState<Record<string, string>>({});

  const deckId = presentation?.id ?? '';
  const pathsToResolve = collectStepAssetPaths(selectedStep);
  const assetSignature = `${deckId}|${pathsToResolve.join('|')}`;

  useEffect(() => {
    setSlideUrls({});
    if (!deckId || !pathsToResolve.length || !window.presentApi?.resolveSlidePreviewAsset) return;

    let mounted = true;
    const createdUrls: string[] = [];

    async function resolveAll() {
      const results: Record<string, string> = {};
      for (const relativePath of pathsToResolve) {
        const asset = await window.presentApi!.resolveSlidePreviewAsset({ deckId, relativePath });
        if (!mounted) return;
        if (!asset) continue;

        const blobUrl = URL.createObjectURL(
          new Blob([toBlobPart(asset.buffer)], { type: asset.mimeType || 'application/octet-stream' }),
        );
        createdUrls.push(blobUrl);
        results[`${deckId}|${relativePath}`] = blobUrl;
      }
      if (!mounted) {
        createdUrls.forEach((url) => URL.revokeObjectURL(url));
        return;
      }
      setSlideUrls(results);
    }

    void resolveAll();
    return () => {
      mounted = false;
      createdUrls.forEach((url) => URL.revokeObjectURL(url));
    };
  }, [assetSignature]);

  return slideUrls;
}

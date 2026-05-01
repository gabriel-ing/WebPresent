import type { PptxShape } from './types';

/**
 * Clamps a fractional crop value to the valid range [0, 0.9999].
 * Values outside this range would produce invisible or inverted images.
 */
export function clampCrop(value: number | undefined): number {
  const n = Number(value) || 0;
  return Math.min(0.9999, Math.max(0, n));
}

/**
 * Produces inline CSS for an image that needs crop and/or flip transforms.
 *
 * The parent container uses `overflow:hidden` and the image is sized/
 * positioned to show only the cropped region. Flip transforms are applied
 * via CSS `scaleX(-1)` / `scaleY(-1)`.
 */
export function getCroppedImageStyles(shape: PptxShape): string {
  const leftCrop = clampCrop(shape.imageCrop?.left);
  const topCrop = clampCrop(shape.imageCrop?.top);
  const rightCrop = clampCrop(shape.imageCrop?.right);
  const bottomCrop = clampCrop(shape.imageCrop?.bottom);
  const visibleWidth = Math.max(0.0001, 1 - leftCrop - rightCrop);
  const visibleHeight = Math.max(0.0001, 1 - topCrop - bottomCrop);

  const transforms: string[] = [];
  if (shape.flipH) transforms.push('scaleX(-1)');
  if (shape.flipV) transforms.push('scaleY(-1)');

  const styles: string[] = [
    'position:absolute',
    `left:${(-leftCrop / visibleWidth) * 100}%`,
    `top:${(-topCrop / visibleHeight) * 100}%`,
    `width:${(1 / visibleWidth) * 100}%`,
    `height:${(1 / visibleHeight) * 100}%`,
    'display:block',
  ];

  if (transforms.length) {
    styles.push(`transform:${transforms.join(' ')}`);
    styles.push('transform-origin:center center');
  }

  return styles.join(';');
}

/**
 * Shared utility functions for the Electron main process.
 */

const crypto = require('node:crypto');
const fs = require('node:fs/promises');
const path = require('node:path');
const { PNG } = require('pngjs');
const UTIF = require('utif');

// ── ID / Time ─────────────────────────────────────────────────────────────────

function nowIso() {
  return new Date().toISOString();
}

function createId(prefix) {
  return `${prefix}-${crypto.randomUUID()}`;
}

// ── File extension helpers ─────────────────────────────────────────────────────

const MIME_MAP = new Map([
  ['.png', 'image/png'],
  ['.jpg', 'image/jpeg'],
  ['.jpeg', 'image/jpeg'],
  ['.gif', 'image/gif'],
  ['.svg', 'image/svg+xml'],
  ['.tif', 'image/tiff'],
  ['.tiff', 'image/tiff'],
  ['.webp', 'image/webp'],
  ['.mp4', 'video/mp4'],
  ['.m4v', 'video/mp4'],
  ['.webm', 'video/webm'],
  ['.ogv', 'video/ogg'],
  ['.ogg', 'video/ogg'],
  ['.mov', 'video/quicktime'],
]);

const SLIDE_COPY_EXTENSIONS = new Set([
  '.png', '.jpg', '.jpeg', '.gif', '.svg',
  '.tif', '.tiff', '.webp',
  '.mp4', '.webm', '.mov', '.m4v', '.ogv', '.ogg',
]);

const VIDEO_EXTENSIONS = new Set(['.mp4', '.webm', '.mov', '.m4v', '.ogv', '.ogg']);

/**
 * Returns the extension to use when copying a slide file.
 * Falls back to '.png' for unknown types.
 */
function extensionFor(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return SLIDE_COPY_EXTENSIONS.has(ext) ? ext : '.png';
}

function mimeTypeFromPath(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return MIME_MAP.get(ext) || 'application/octet-stream';
}

function mediaKindFromPath(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return VIDEO_EXTENSIONS.has(ext) ? 'video' : 'image';
}

function bufferToDataUrl(buffer, mimeType) {
  return `data:${mimeType};base64,${buffer.toString('base64')}`;
}

function normalizeTiffDimension(value) {
  if (Array.isArray(value)) return normalizeTiffDimension(value[0]);
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

function convertTiffBufferToPngBuffer(buffer) {
  const ifds = UTIF.decode(buffer);
  const frame = Array.isArray(ifds) ? ifds[0] : null;
  if (!frame) return null;

  UTIF.decodeImage(buffer, frame);

  const width = normalizeTiffDimension(frame.width ?? frame.t256);
  const height = normalizeTiffDimension(frame.height ?? frame.t257);
  if (!width || !height) return null;

  const rgba = Buffer.from(UTIF.toRGBA8(frame));
  const png = new PNG({ width, height });
  rgba.copy(png.data);

  return PNG.sync.write(png);
}

function convertTiffBufferToPngDataUrl(buffer) {
  const pngBuffer = convertTiffBufferToPngBuffer(buffer);
  return pngBuffer ? bufferToDataUrl(pngBuffer, 'image/png') : null;
}

async function readFileForPreviewAsset(filePath, options = {}) {
  const readFile = options.readFile || fs.readFile;
  const buffer = await readFile(filePath);
  const mimeType = mimeTypeFromPath(filePath);

  if (mimeType === 'image/tiff') {
    try {
      const pngBuffer = convertTiffBufferToPngBuffer(buffer);
      if (pngBuffer) {
        return {
          buffer: pngBuffer,
          mimeType: 'image/png',
        };
      }
    } catch {
      // Fall back to the original bytes if conversion fails.
    }
  }

  return { buffer, mimeType };
}

async function readFileAsDataUrl(filePath, options = {}) {
  const asset = await readFileForPreviewAsset(filePath, options);
  return bufferToDataUrl(asset.buffer, asset.mimeType);
}

module.exports = {
  nowIso,
  createId,
  extensionFor,
  mimeTypeFromPath,
  mediaKindFromPath,
  readFileForPreviewAsset,
  readFileAsDataUrl,
  SLIDE_COPY_EXTENSIONS,
};

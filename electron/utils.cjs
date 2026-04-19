/**
 * Shared utility functions for the Electron main process.
 */

const crypto = require('node:crypto');
const path = require('node:path');

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

module.exports = {
  nowIso,
  createId,
  extensionFor,
  mimeTypeFromPath,
  mediaKindFromPath,
  SLIDE_COPY_EXTENSIONS,
};

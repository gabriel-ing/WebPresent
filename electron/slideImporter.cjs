/**
 * Slide file import — copies media files into a deck's slides directory,
 * returns SlideRef objects that can be stored in the deck JSON.
 */

const path = require('node:path');
const fs = require('node:fs/promises');
const { dialog } = require('electron');
const { extensionFor, mediaKindFromPath, SLIDE_COPY_EXTENSIONS } = require('./utils.cjs');
const { getDeckDir } = require('./deckManager.cjs');
const { sortSlidePaths } = require('./fileSortUtils.cjs');

const DIRECTORY_IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.svg', '.tif', '.tiff']);

// ── Sequence numbering ────────────────────────────────────────────────────────

async function nextSlideNumber(deckId) {
  const slidesDir = path.join(getDeckDir(deckId), 'slides');
  await fs.mkdir(slidesDir, { recursive: true });
  const names = await fs.readdir(slidesDir);
  let max = 0;
  for (const name of names) {
    const match = name.match(/^slide-(\d{4})\./i);
    if (!match) continue;
    const number = Number.parseInt(match[1], 10);
    if (number > max) max = number;
  }
  return max + 1;
}

// ── Import ────────────────────────────────────────────────────────────────────

/**
 * Copies media files into the deck slides directory and returns SlideRef objects.
 * @param {string} deckId
 * @param {string[]} filePaths  Absolute paths to media files
 * @returns {Promise<SlideRef[]>}
 */
async function importSlidesToDeck(deckId, filePaths) {
  const slidesDir = path.join(getDeckDir(deckId), 'slides');
  await fs.mkdir(slidesDir, { recursive: true });
  const refs = [];
  let seq = await nextSlideNumber(deckId);
  const orderedPaths = sortSlidePaths(filePaths);
  for (const filePath of orderedPaths) {
    const ext = extensionFor(filePath);
    const slideBase = `slide-${String(seq).padStart(4, '0')}`;
    const fileName = `${slideBase}${ext}`;
    const destination = path.join(slidesDir, fileName);
    await fs.copyFile(filePath, destination);
    refs.push({
      id: slideBase,
      relativePath: path.posix.join('slides', fileName),
      sourceFileName: path.basename(filePath),
      mediaKind: mediaKindFromPath(filePath),
    });
    seq += 1;
  }
  return refs;
}

// ── Dialog helpers ────────────────────────────────────────────────────────────

async function pickSlideFiles(ownerWindow) {
  const result = await dialog.showOpenDialog(ownerWindow ?? undefined, {
    title: 'Select slide image or video files',
    properties: ['openFile', 'multiSelections'],
    filters: [
      {
        name: 'Media',
        extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'svg', 'tif', 'tiff', 'mp4', 'webm', 'mov', 'm4v', 'ogv', 'ogg'],
      },
    ],
  });
  return result.canceled ? [] : result.filePaths;
}

async function pickDirectoryImageFiles(ownerWindow) {
  const result = await dialog.showOpenDialog(ownerWindow ?? undefined, {
    title: 'Select slide image folder',
    properties: ['openDirectory'],
  });
  if (result.canceled || !result.filePaths[0]) return [];

  const directoryPath = result.filePaths[0];
  const entries = await fs.readdir(directoryPath, { withFileTypes: true });
  const files = entries
    .filter((entry) => entry.isFile())
    .map((entry) => path.join(directoryPath, entry.name))
    .filter((filePath) => DIRECTORY_IMAGE_EXTENSIONS.has(path.extname(filePath).toLowerCase()));

  return sortSlidePaths(files);
}

async function pickPptxFile(ownerWindow) {
  const result = await dialog.showOpenDialog(ownerWindow ?? undefined, {
    title: 'Select PowerPoint file',
    properties: ['openFile'],
    filters: [{ name: 'PowerPoint', extensions: ['pptx'] }],
  });
  if (result.canceled || !result.filePaths[0]) return null;
  return result.filePaths[0];
}

module.exports = {
  importSlidesToDeck,
  pickSlideFiles,
  pickDirectoryImageFiles,
  pickPptxFile,
};

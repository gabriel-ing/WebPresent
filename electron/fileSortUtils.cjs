/**
 * Pure filename-sorting helpers for slide files.
 * No Electron dependency — safe to use in tests and renderers.
 */

const path = require('node:path');

const naturalFileNameCollator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });

/**
 * Extracts a numeric slide number from filenames like "slide1", "Slide 02", etc.
 * Returns null if the filename doesn't match the pattern.
 */
function extractSlideNumber(filePath) {
  const name = path.basename(filePath, path.extname(filePath));
  const match = name.match(/^slide\s*([0-9]+)$/i);
  return match ? Number.parseInt(match[1], 10) : null;
}

function compareSlideFilePaths(a, b) {
  const numberA = extractSlideNumber(a);
  const numberB = extractSlideNumber(b);
  if (numberA !== null && numberB !== null && numberA !== numberB) return numberA - numberB;
  if (numberA !== null && numberB === null) return -1;
  if (numberA === null && numberB !== null) return 1;
  return naturalFileNameCollator.compare(path.basename(a), path.basename(b));
}

/**
 * Returns a new array sorted so that "slide<N>" names come first in numeric
 * order, then everything else in locale-aware alphanumeric order.
 */
function sortSlidePaths(filePaths) {
  return [...filePaths].sort(compareSlideFilePaths);
}

module.exports = { extractSlideNumber, sortSlidePaths };

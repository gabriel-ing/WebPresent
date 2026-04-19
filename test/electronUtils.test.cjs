const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

const { extensionFor, mimeTypeFromPath, mediaKindFromPath } = require('../electron/utils.cjs');
const { extractSlideNumber, sortSlidePaths } = require('../electron/fileSortUtils.cjs');

// ── extensionFor ─────────────────────────────────────────────────────────────

test('extensionFor returns correct extension for known types', () => {
  assert.equal(extensionFor('/slides/photo.jpg'), '.jpg');
  assert.equal(extensionFor('/slides/clip.mp4'), '.mp4');
  assert.equal(extensionFor('/slides/diagram.svg'), '.svg');
});

test('extensionFor falls back to .png for unknown types', () => {
  assert.equal(extensionFor('/slides/document.pdf'), '.png');
  assert.equal(extensionFor('/slides/noextension'), '.png');
});

// ── mimeTypeFromPath ──────────────────────────────────────────────────────────

test('mimeTypeFromPath returns correct MIME type', () => {
  assert.equal(mimeTypeFromPath('/x/image.png'), 'image/png');
  assert.equal(mimeTypeFromPath('/x/image.jpg'), 'image/jpeg');
  assert.equal(mimeTypeFromPath('/x/image.jpeg'), 'image/jpeg');
  assert.equal(mimeTypeFromPath('/x/file.webp'), 'image/webp');
  assert.equal(mimeTypeFromPath('/x/video.mp4'), 'video/mp4');
  assert.equal(mimeTypeFromPath('/x/video.m4v'), 'video/mp4');
  assert.equal(mimeTypeFromPath('/x/video.webm'), 'video/webm');
  assert.equal(mimeTypeFromPath('/x/video.mov'), 'video/quicktime');
  assert.equal(mimeTypeFromPath('/x/video.ogg'), 'video/ogg');
  assert.equal(mimeTypeFromPath('/x/unknown.xyz'), 'application/octet-stream');
});

// ── mediaKindFromPath ─────────────────────────────────────────────────────────

test('mediaKindFromPath identifies video extensions', () => {
  assert.equal(mediaKindFromPath('/x/clip.mp4'), 'video');
  assert.equal(mediaKindFromPath('/x/clip.webm'), 'video');
  assert.equal(mediaKindFromPath('/x/clip.mov'), 'video');
  assert.equal(mediaKindFromPath('/x/clip.m4v'), 'video');
});

test('mediaKindFromPath identifies image extensions', () => {
  assert.equal(mediaKindFromPath('/x/photo.png'), 'image');
  assert.equal(mediaKindFromPath('/x/photo.jpg'), 'image');
  assert.equal(mediaKindFromPath('/x/diagram.svg'), 'image');
});

// ── extractSlideNumber ────────────────────────────────────────────────────────

test('extractSlideNumber returns number for matching filenames', () => {
  assert.equal(extractSlideNumber('/path/slide1.png'), 1);
  assert.equal(extractSlideNumber('/path/Slide10.jpg'), 10);
  assert.equal(extractSlideNumber('/path/slide 03.png'), 3);
  assert.equal(extractSlideNumber('/path/SLIDE42.png'), 42);
});

test('extractSlideNumber returns null for non-matching filenames', () => {
  assert.equal(extractSlideNumber('/path/photo.png'), null);
  assert.equal(extractSlideNumber('/path/slide_extra.png'), null);
  assert.equal(extractSlideNumber('/path/myslide1.png'), null);
});

// ── sortSlidePaths ────────────────────────────────────────────────────────────

test('sortSlidePaths orders slide<N> files numerically', () => {
  const input = ['/p/slide10.png', '/p/slide2.png', '/p/slide1.png'];
  const result = sortSlidePaths(input);
  assert.deepEqual(result, ['/p/slide1.png', '/p/slide2.png', '/p/slide10.png']);
});

test('sortSlidePaths keeps slide<N> before non-slide names', () => {
  const input = ['/p/alpha.png', '/p/slide1.png', '/p/beta.png'];
  const result = sortSlidePaths(input);
  assert.equal(result[0], '/p/slide1.png');
});

test('sortSlidePaths does not mutate the original array', () => {
  const input = ['/p/slide2.png', '/p/slide1.png'];
  const copy = [...input];
  sortSlidePaths(input);
  assert.deepEqual(input, copy);
});

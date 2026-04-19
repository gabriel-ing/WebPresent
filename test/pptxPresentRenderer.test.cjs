const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const os = require('node:os');
const { mkdtemp, mkdir, rm, writeFile } = require('node:fs/promises');
const { parsePptx } = require('../electron/pptxParser.cjs');

const {
  buildPptxPresentationDocument,
  buildPptxRuntimeUpdateScript,
  createPptxPresentationDocumentBuilder,
} = require('../electron/pptxPresentRenderer.cjs');

const ONE_BY_ONE_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnRsl0AAAAASUVORK5CYII=';

const workspaceRoot = path.resolve(__dirname, '..');
const fixturesDir = path.join(workspaceRoot, 'features', 'pptx');

async function parseFixture(fileName) {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-render-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  await mkdir(getDeckDir(deckId), { recursive: true });

  try {
    const deck = await parsePptx(path.join(fixturesDir, fileName), deckId, getDeckDir);
    return { deck, tempRoot, deckId, getDeckDir };
  } catch (error) {
    await rm(tempRoot, { recursive: true, force: true });
    throw error;
  }
}

test('fullscreen PPTX HTML inlines slide images as data URLs', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');
  const slidesDir = path.join(deckDir, 'slides');

  try {
    await mkdir(slidesDir, { recursive: true });
    await writeFile(path.join(slidesDir, 'image.png'), Buffer.from(ONE_BY_ONE_PNG_BASE64, 'base64'));

    const slide = {
      slideIndex: 0,
      width: 960,
      height: 540,
      shapes: [
        {
          id: 'pic-1',
          type: 'image',
          x: 0,
          y: 0,
          width: 200,
          height: 100,
          imageRelativePath: 'slides/image.png',
          animationGroup: 0,
          animationEffect: 'appear',
        },
      ],
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);

    assert.match(html, /data:image\/png;base64,/);
    assert.doesNotMatch(html, /file:\/\//);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX HTML applies image crop styles and paragraph-build visibility', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');
  const slidesDir = path.join(deckDir, 'slides');

  try {
    await mkdir(slidesDir, { recursive: true });
    await writeFile(path.join(slidesDir, 'image.png'), Buffer.from(ONE_BY_ONE_PNG_BASE64, 'base64'));

    const slide = {
      slideIndex: 0,
      width: 960,
      height: 540,
      shapes: [
        {
          id: 'pic-1',
          type: 'image',
          x: 0,
          y: 0,
          width: 200,
          height: 100,
          imageRelativePath: 'slides/image.png',
          imageCrop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.1 },
          animationGroup: 0,
          animationEffect: 'appear',
        },
        {
          id: 'shape-1',
          type: 'rect',
          x: 0,
          y: 120,
          width: 200,
          height: 120,
          animationGroup: 1,
          animationEffect: 'appear',
          paragraphs: [
            { runs: [{ text: 'First bullet' }], animationGroup: 1 },
            { runs: [{ text: 'Second bullet' }], animationGroup: 2 },
          ],
        },
      ],
      animationStepCount: 2,
    };

    const htmlAtStepOne = await buildPptxPresentationDocument(slide, 1, deckDir);
    const htmlAtStepTwo = await buildPptxPresentationDocument(slide, 2, deckDir);

    assert.match(htmlAtStepOne, /left:-16\.6666666666666/);
    assert.match(htmlAtStepOne, /top:-28\.57142857142857/);
    assert.match(htmlAtStepOne, /width:166\.666666666666/);
    assert.match(htmlAtStepOne, /height:142\.857142857142/);
    assert.match(htmlAtStepOne, /First bullet/);
    assert.doesNotMatch(htmlAtStepOne, /Second bullet/);
    assert.match(htmlAtStepTwo, /Second bullet/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('cached PPTX document builder reuses media file reads across steps', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');
  const slidesDir = path.join(deckDir, 'slides');

  try {
    await mkdir(slidesDir, { recursive: true });
    await writeFile(path.join(slidesDir, 'image.png'), Buffer.from(ONE_BY_ONE_PNG_BASE64, 'base64'));

    let readCount = 0;
    const builder = createPptxPresentationDocumentBuilder(deckDir, {
      readFile: async (...args) => {
        readCount += 1;
        return require('node:fs/promises').readFile(...args);
      },
    });

    const slide = {
      slideIndex: 0,
      width: 960,
      height: 540,
      shapes: [
        {
          id: 'pic-1',
          type: 'image',
          x: 0,
          y: 0,
          width: 200,
          height: 100,
          imageRelativePath: 'slides/image.png',
          animationGroup: 0,
          animationEffect: 'appear',
        },
      ],
      animationStepCount: 1,
    };

    await builder(slide, 0);
    await builder(slide, 1);

    assert.equal(readCount, 1, 'expected repeated renders to reuse the same inlined media data URL');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX runtime helper supports in-place slide updates between builds', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');
  const slidesDir = path.join(deckDir, 'slides');

  try {
    await mkdir(slidesDir, { recursive: true });
    await writeFile(path.join(slidesDir, 'image.png'), Buffer.from(ONE_BY_ONE_PNG_BASE64, 'base64'));

    const slide = {
      slideIndex: 0,
      width: 960,
      height: 540,
      shapes: [
        {
          id: 'shape-1',
          type: 'rect',
          x: 0,
          y: 120,
          width: 200,
          height: 120,
          animationGroup: 1,
          animationEffect: 'appear',
          paragraphs: [
            { runs: [{ text: 'First bullet' }], animationGroup: 1 },
            { runs: [{ text: 'Second bullet' }], animationGroup: 2 },
          ],
        },
      ],
      animationStepCount: 2,
    };

    const html = await buildPptxPresentationDocument(slide, 1, deckDir);
    const updateScript = await buildPptxRuntimeUpdateScript(slide, 2, deckDir);

    assert.match(html, /window\.__WEBPRESENT_UPDATE_PPTX/);
    assert.match(updateScript, /window\.__WEBPRESENT_UPDATE_PPTX/);
    assert.match(updateScript, /Second bullet/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX HTML converts TIFF slide assets to browser-safe PNG data URLs', async () => {
  const { deck, tempRoot, deckId, getDeckDir } = await parseFixture('MMC_simpliPyTEM_presentation.pptx');

  try {
    const slideEleven = deck.slides[10];

    assert.ok(
      slideEleven.shapes.some((shape) => /\.tiff?$/i.test(shape.imageRelativePath || '')),
      'expected the fixture slide to include TIFF-backed images',
    );

    const html = await buildPptxPresentationDocument(slideEleven, 0, getDeckDir(deckId));

    assert.match(html, /data:image\/png;base64,/);
    assert.doesNotMatch(
      html,
      /data:image\/tiff;base64,/,
      'expected TIFF-backed slide assets to be converted before embedding in browser HTML',
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});
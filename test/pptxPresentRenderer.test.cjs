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

test('PPTX HTML styles bullet markers using the paragraph typography', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');

  try {
    await mkdir(deckDir, { recursive: true });

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
          width: 400,
          height: 120,
          animationGroup: 0,
          paragraphs: [
            {
              bulletType: 'bullet',
              bulletChar: '•',
              runs: [
                {
                  text: 'Learn something new',
                  fontSize: 28,
                  fontFamily: 'Gotham Book',
                  colour: '#000000',
                },
              ],
            },
          ],
        },
      ],
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);

    assert.match(html, /margin-right:0\.4em;font-size:28pt;font-family:Gotham Book,sans-serif;color:#000000|font-size:28pt;font-family:Gotham Book,sans-serif;color:#000000;margin-right:0\.4em/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX HTML keeps empty bullet spacer paragraphs without rendering a bullet marker', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');

  try {
    await mkdir(deckDir, { recursive: true });

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
          width: 500,
          height: 220,
          animationGroup: 0,
          paragraphs: [
            {
              bulletType: 'bullet',
              runs: [
                {
                  text: 'First bullet',
                  fontSize: 20,
                  fontFamily: 'Verdana',
                },
              ],
            },
            {
              bulletType: 'bullet',
              runs: [
                {
                  text: '',
                  fontSize: 20,
                  fontFamily: 'Verdana',
                },
              ],
            },
            {
              bulletType: 'bullet',
              runs: [
                {
                  text: 'Second bullet',
                  fontSize: 20,
                  fontFamily: 'Verdana',
                },
              ],
            },
          ],
        },
      ],
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);
    const bulletCount = (html.match(/>•<\/span>/g) || []).length;

    assert.equal(bulletCount, 2);
    assert.match(html, /&nbsp;/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX HTML renders straight connector arrows as SVG lines', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');

  try {
    await mkdir(deckDir, { recursive: true });

    const slide = {
      slideIndex: 0,
      width: 960,
      height: 540,
      shapes: [
        {
          id: 'shape-connector',
          type: 'line',
          x: 100,
          y: 120,
          width: 200,
          height: 100,
          flipV: true,
          lineTail: 'triangle',
          border: {
            width: 6,
            colour: '#000000',
            style: 'solid',
          },
          animationGroup: 0,
        },
      ],
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);

    assert.match(html, /marker-end=\\"url\(#shape-connector-tail\)\\"/);
    assert.match(html, /<line[^>]*x1=\\"0\\"[^>]*y1=\\"100\\"[^>]*x2=\\"200\\"[^>]*y2=\\"0\\"/);
    assert.doesNotMatch(html, /border-top:/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX HTML honors explicit zero text insets on text boxes', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');

  try {
    await mkdir(deckDir, { recursive: true });

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
          width: 320,
          height: 40,
          animationGroup: 0,
          textInsets: { left: 0, top: 0, right: 0, bottom: 0 },
          paragraphs: [
            {
              runs: [
                {
                  text: 'What the agent sees:',
                  fontSize: 16,
                  fontFamily: 'Verdana',
                  bold: true,
                },
              ],
            },
          ],
        },
      ],
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);

    assert.match(html, /padding:0px 0px 0px 0px/);
    assert.match(html, /What the agent sees:/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX HTML applies paragraph line spacing and omits trailing empty paragraphs', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');

  try {
    await mkdir(deckDir, { recursive: true });

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
          width: 400,
          height: 220,
          animationGroup: 0,
          paragraphs: [
            {
              lineSpacing: 2,
              runs: [
                {
                  text: 'Schedule item',
                  fontSize: 20,
                  fontFamily: 'Gotham Book',
                  colour: '#000000',
                },
              ],
            },
            {
              bulletType: 'bullet',
              runs: [
                {
                  text: '',
                  fontSize: 18,
                  fontFamily: 'Gotham Book',
                  colour: '#000000',
                },
              ],
            },
          ],
        },
      ],
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);

    assert.match(html, /line-height:2/);
    assert.match(html, /Schedule item/);
    assert.doesNotMatch(html, /•/);
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

test('fullscreen PPTX HTML applies the initial state after the slide root exists', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');

  try {
    await mkdir(deckDir, { recursive: true });

    const slide = {
      slideIndex: 0,
      width: 400,
      height: 225,
      shapes: [
        {
          id: 'shape-1',
          type: 'rect',
          x: 10,
          y: 20,
          width: 100,
          height: 50,
          animationGroup: 0,
          animationEffect: 'appear',
        },
      ],
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);
    const slideRootIndex = html.indexOf('<div class="slide-root"></div>');
    const initialUpdateIndex = html.lastIndexOf('window.__WEBPRESENT_UPDATE_PPTX(');

    assert.ok(slideRootIndex !== -1, 'expected the fullscreen document to include the slide root');
    assert.ok(
      initialUpdateIndex > slideRootIndex,
      'expected the initial PPTX state to be applied after the slide root exists',
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX HTML scales text runs when the shape requests norm autofit', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');

  try {
    await mkdir(deckDir, { recursive: true });

    const slide = {
      slideIndex: 0,
      width: 400,
      height: 225,
      shapes: [
        {
          id: 'shape-1',
          type: 'rect',
          x: 0,
          y: 0,
          width: 200,
          height: 100,
          textFitScale: 0.9,
          paragraphs: [
            {
              runs: [
                {
                  text: 'Scaled title',
                  fontSize: 44,
                  fontFamily: 'Verdana',
                  colour: '#00B2A9',
                  bold: true,
                },
              ],
            },
          ],
          animationGroup: 0,
          animationEffect: 'appear',
        },
      ],
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);

    assert.match(html, /font-size:39\.6pt/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX HTML preserves configured gradient angles and stop ordering', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');

  try {
    await mkdir(deckDir, { recursive: true });

    const slide = {
      slideIndex: 0,
      width: 400,
      height: 225,
      shapes: [],
      background: {
        type: 'gradient',
        gradientAngle: 45,
        gradientStops: [
          { position: 100, colour: '#2F2A95' },
          { position: 0, colour: '#FFFFFF' },
          { position: 62, colour: '#000000' },
        ],
      },
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);

    assert.match(
      html,
      /"backgroundCss":"linear-gradient\(45deg, #FFFFFF 0%, #000000 62%, #2F2A95 100%\)"/,
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('PPTX HTML renders freeform geometry as SVG instead of axis-aligned boxes', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-present-'));
  const deckDir = path.join(tempRoot, 'deck');

  try {
    await mkdir(deckDir, { recursive: true });

    const slide = {
      slideIndex: 0,
      width: 320,
      height: 180,
      shapes: [
        {
          id: 'freeform-1',
          type: 'freeform',
          x: 20,
          y: 10,
          width: 100,
          height: 120,
          svgPath: 'M 0 0 L 100 20 L 100 120 L 0 90 Z',
          svgViewBoxWidth: 100,
          svgViewBoxHeight: 120,
          fill: {
            type: 'gradient',
            gradientAngle: 210,
            gradientStops: [
              { position: 0, colour: '#92C0E9' },
              { position: 100, colour: '#2F2A95' },
            ],
          },
          border: {
            width: 1.5,
            gradientAngle: 180,
            gradientStops: [
              { position: 0, colour: '#00B2A9' },
              { position: 100, colour: '#006B68' },
            ],
          },
          animationGroup: 0,
          animationEffect: 'appear',
        },
      ],
      animationStepCount: 0,
    };

    const html = await buildPptxPresentationDocument(slide, 0, deckDir);

    assert.match(html, /<svg[^>]+viewBox=\\"0 0 100 120\\"/);
    assert.match(html, /<path[^>]+d=\\"M 0 0 L 100 20 L 100 120 L 0 90 Z\\"/);
    assert.match(html, /<linearGradient/);
    assert.doesNotMatch(html, /background:linear-gradient\(210deg/);
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
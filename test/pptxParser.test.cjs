const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const os = require('node:os');
const { mkdtemp, rm, mkdir } = require('node:fs/promises');
const AdmZip = require('adm-zip');

const { parsePptx } = require('../electron/pptxParser.cjs');

const workspaceRoot = path.resolve(__dirname, '..');
const fixturesDir = path.join(workspaceRoot, 'features', 'pptx');

function countSlidePictures(pptxPath) {
  const zip = new AdmZip(pptxPath);
  return zip
    .getEntries()
    .filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry.entryName))
    .reduce((total, entry) => {
      const xml = entry.getData().toString('utf8');
      const matches = xml.match(/<p:pic(?=[\s>])/g);
      return total + (matches ? matches.length : 0);
    }, 0);
}

async function parseFixture(fileName) {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-test-'));
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

test('parser extracts grouped slide pictures as image shapes', async () => {
  const fileName = 'CQC - DV Portfolio_GI.pptx';
  const expectedPictures = countSlidePictures(path.join(fixturesDir, fileName));
  const { deck, tempRoot } = await parseFixture(fileName);

  try {
    const parsedPictures = deck.slides
      .flatMap((slide) => slide.shapes)
      .filter((shape) => shape.type === 'image' && shape.imageRelativePath);

    assert.ok(expectedPictures > 0, 'fixture should contain slide pictures');
    assert.ok(
      parsedPictures.length >= expectedPictures,
      `expected at least ${expectedPictures} rendered image shapes, got ${parsedPictures.length}`,
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser turns click-build timing into animation steps', async () => {
  const { deck, tempRoot } = await parseFixture('GabrielIng_slideForGRC.pptx');

  try {
    assert.equal(deck.slides.length, 1, 'fixture should contain exactly one slide');
    const [slide] = deck.slides;

    assert.ok(slide.animationStepCount >= 1, 'expected at least one click-build animation step');
    assert.ok(
      slide.shapes.some((shape) => shape.animationGroup >= 1),
      'expected at least one parsed shape to be assigned to an animation group',
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser extracts slide-number fields as text runs', async () => {
  const { deck, tempRoot } = await parseFixture('1-CHEMA1_intro_regression.pptx');

  try {
    const slide = deck.slides[1];
    const slideNumberShape = slide.shapes.find((shape) => shape.placeholder?.type === 'sldNum');

    assert.ok(slideNumberShape, 'expected a slide-number placeholder shape');
    assert.equal(
      slideNumberShape.paragraphs?.[0]?.runs?.[0]?.text,
      '2',
      'expected the slide-number field text to be preserved',
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser keeps picture crop metadata for cropped images', async () => {
  const { deck, tempRoot } = await parseFixture('CQC - DV Portfolio_GI.pptx');

  try {
    const slide = deck.slides[1];
    const croppedImage = slide.shapes.find((shape) => shape.type === 'image' && shape.sourceShapeId === '5');

    assert.ok(croppedImage, 'expected the cropped picture shape to be parsed');
    assert.equal(croppedImage.imageCrop?.left, 0.00586);
    assert.equal(croppedImage.imageCrop?.top, 0);
    assert.equal(croppedImage.imageCrop?.right, 0);
    assert.equal(croppedImage.imageCrop?.bottom, 0);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser ignores interactive timing sequences and preserves paragraph build order', async () => {
  const { deck, tempRoot } = await parseFixture('MMC_simpliPyTEM_presentation.pptx');

  try {
    const slideTwo = deck.slides[1];
    const paragraphBuildShape = slideTwo.shapes.find((shape) => shape.sourceShapeId === '204');

    assert.ok(paragraphBuildShape, 'expected the paragraph-build text shape to be parsed');
    assert.equal(paragraphBuildShape.animationGroup, 1, 'paragraph builds should start on their first click');
    assert.deepEqual(
      paragraphBuildShape.paragraphs?.map((paragraph) => paragraph.animationGroup),
      [1, 2, 5, 6, 7],
      'expected each paragraph to keep its click order',
    );

    const slideThree = deck.slides[2];
    const groupByShapeId = Object.fromEntries(
      ['225', '223', '221', '224', '222'].map((shapeId) => [
        shapeId,
        slideThree.shapes.find((shape) => shape.sourceShapeId === shapeId)?.animationGroup,
      ]),
    );

    assert.deepEqual(groupByShapeId, {
      '221': 2,
      '222': 3,
      '223': 2,
      '224': 3,
      '225': 1,
    });

    const secondParagraphBuildShape = slideThree.shapes.find((shape) => shape.sourceShapeId === '219');
    assert.deepEqual(
      secondParagraphBuildShape.paragraphs?.map((paragraph) => paragraph.animationGroup),
      [4, 5, 6, 7],
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser respects layout master-shape suppression to avoid duplicate banner graphics', async () => {
  const { deck, tempRoot } = await parseFixture('MMC_simpliPyTEM_presentation.pptx');

  try {
    const slideTwo = deck.slides[1];
    const brandingShapes = slideTwo.shapes.filter((shape) => shape.name === 'UCL Branding');

    assert.equal(
      brandingShapes.length,
      2,
      'expected the slide to render the layout banner only once instead of duplicating master and layout graphics',
    );
    assert.deepEqual(
      brandingShapes.map((shape) => shape.sourceShapeId).sort(),
      ['183', '184'],
      'expected duplicated master banner shapes to be suppressed when the layout disables master shapes',
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser keeps alternate-content fallback shapes so equation text boxes still render', async () => {
  const { deck, tempRoot } = await parseFixture('1-CHEMA1_intro_regression.pptx');

  try {
    const slideFour = deck.slides[3];
    const equationShape = slideFour.shapes.find((shape) => shape.sourceShapeId === '307207');

    assert.ok(equationShape, 'expected the alternate-content equation text box to be preserved');
    assert.equal(equationShape.fill?.type, 'image', 'expected the fallback rendering to keep the equation box visible');
    assert.ok(equationShape.fill?.imageRelativePath, 'expected the fallback equation graphic to be extracted');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});
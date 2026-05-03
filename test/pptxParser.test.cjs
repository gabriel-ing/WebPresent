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

async function createSyntheticPptx(pptxPath, options = {}) {
  const titleRuns = options.titleRuns || ['Welcome to the 1st ', 'InterSystems READY Hackathon'];
  const titleFillXml = options.titleFillXml || '<a:solidFill><a:srgbClr val="00B2A9"/></a:solidFill>';
  const extraSlideShapesXml = options.extraSlideShapesXml || '';
  const themeFmtSchemeXml = options.themeFmtSchemeXml || '    <a:fmtScheme name="Synthetic Format"/>';
  const layoutTitleStyleXml =
    options.layoutTitleStyleXml || `            <a:lvl1pPr algn="ctr">
              <a:defRPr sz="4400">
                ${titleFillXml}
                <a:latin typeface="Verdana"/>
              </a:defRPr>
            </a:lvl1pPr>`;
  const masterTitleStyleXml =
    options.masterTitleStyleXml || `      <a:lvl1pPr algn="l">
        <a:defRPr sz="2800">
          <a:solidFill><a:srgbClr val="222222"/></a:solidFill>
          <a:latin typeface="Gotham Book"/>
        </a:defRPr>
      </a:lvl1pPr>`;
  const titleRunsXml = titleRuns
    .map(
      (text) => `            <a:r>\n              <a:rPr lang="en-US"/>\n              <a:t>${text}</a:t>\n            </a:r>`,
    )
    .join('\n');
  const zip = new AdmZip();

  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
</Types>`,
      'utf8',
    ),
  );

  zip.addFile(
    'ppt/presentation.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`,
      'utf8',
    ),
  );

  zip.addFile(
    'ppt/_rels/presentation.xml.rels',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`,
      'utf8',
    ),
  );

  zip.addFile(
    'ppt/theme/theme1.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Synthetic Theme">
  <a:themeElements>
    <a:clrScheme name="Synthetic">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="222222"/></a:dk2>
      <a:lt2><a:srgbClr val="F3F3F3"/></a:lt2>
      <a:accent1><a:srgbClr val="00B2A9"/></a:accent1>
      <a:accent2><a:srgbClr val="2F2A95"/></a:accent2>
      <a:accent3><a:srgbClr val="F5A623"/></a:accent3>
      <a:accent4><a:srgbClr val="6B7280"/></a:accent4>
      <a:accent5><a:srgbClr val="10B981"/></a:accent5>
      <a:accent6><a:srgbClr val="0EA5E9"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Synthetic Fonts">
      <a:majorFont>
        <a:latin typeface="Aptos Display"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Aptos"/>
      </a:minorFont>
    </a:fontScheme>
  ${themeFmtSchemeXml}
  </a:themeElements>
</a:theme>`,
      'utf8',
    ),
  );

  zip.addFile(
    'ppt/slides/slide1.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="3352800" y="2717321"/>
            <a:ext cx="5486400" cy="1831796"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr><a:normAutofit fontScale="90000"/></a:bodyPr>
          <a:lstStyle/>
          <a:p>
${titleRunsXml}
          </a:p>
        </p:txBody>
      </p:sp>
${extraSlideShapesXml}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`,
      'utf8',
    ),
  );

  zip.addFile(
    'ppt/slides/_rels/slide1.xml.rels',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
      'utf8',
    ),
  );

  zip.addFile(
    'ppt/slideLayouts/slideLayout1.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Centered Title Placeholder"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="546100" y="1524001"/>
            <a:ext cx="7023100" cy="2317750"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="mid"/>
          <a:lstStyle>
${layoutTitleStyleXml}
          </a:lstStyle>
          <a:p>
            <a:pPr lvl="0"/>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>Layout title</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="10" name="Logo"/>
          <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip>
            <a:extLst>
              <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
                <asvg:svgBlip r:embed="rId2"/>
              </a:ext>
            </a:extLst>
          </a:blip>
          <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="570977" y="459740"/>
            <a:ext cx="2219325" cy="591820"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>
      <p:grpSp>
        <p:nvGrpSpPr>
          <p:cNvPr id="19" name="Decorative Group"/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="11430001" y="476250"/>
            <a:ext cx="213430" cy="571500"/>
            <a:chOff x="4249" y="1683"/>
            <a:chExt cx="695" cy="1861"/>
          </a:xfrm>
        </p:grpSpPr>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="20" name="Grouped Mark"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="4249" y="1824"/>
              <a:ext cx="463" cy="1720"/>
            </a:xfrm>
            <a:solidFill><a:srgbClr val="2F2A95"/></a:solidFill>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
        </p:sp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`,
      'utf8',
    ),
  );

  zip.addFile(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
</Relationships>`,
      'utf8',
    ),
  );

  zip.addFile(
    'ppt/slideMasters/slideMaster1.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree/></p:cSld>
  <p:txStyles>
    <p:titleStyle>
${masterTitleStyleXml}
    </p:titleStyle>
  </p:txStyles>
</p:sldMaster>`,
      'utf8',
    ),
  );

  zip.addFile(
    'ppt/media/image1.svg',
    Buffer.from(
      `<svg xmlns="http://www.w3.org/2000/svg" width="225" height="60" viewBox="0 0 225 60"><rect width="225" height="60" fill="#ffffff"/><text x="10" y="38" fill="#2F2A95" font-size="28">InterSystems</text></svg>`,
      'utf8',
    ),
  );

  zip.writeZip(pptxPath);
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

test('parser resolves layout SVG pictures and ctrTitle placeholder styles for title slides', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'layout-svg-placeholder.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath);

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];

    const titleShape = slide.shapes.find((shape) => shape.sourceShapeId === '2');
    assert.ok(titleShape, 'expected the title placeholder shape to be parsed');
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[0]?.fontSize, 44);
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[0]?.colour, '#00B2A9');

    const logoShape = slide.shapes.find((shape) => shape.type === 'image' && shape.name === 'Logo');
    assert.ok(logoShape, 'expected the layout SVG picture to be parsed');
    assert.ok(logoShape.imageRelativePath, 'expected the layout SVG picture to be extracted');
    assert.match(logoShape.imageRelativePath, /^slides\/pptx-media-\d+\.svg$/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves normAutofit fontScale on title placeholders', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'layout-svg-placeholder.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath);

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const titleShape = slide.shapes.find((shape) => shape.sourceShapeId === '2');

    assert.ok(titleShape, 'expected the title placeholder shape to be parsed');
    assert.equal(titleShape.textFitScale, 0.9);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves explicit bodyPr text insets on text boxes', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'text-insets.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, {
      extraSlideShapesXml: `      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Inset textbox"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="914400"/>
            <a:ext cx="3048000" cy="457200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:spAutoFit/></a:bodyPr>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1600" b="1"/>
              <a:t>What the agent sees:</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>`,
    });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const textShape = slide.shapes.find((shape) => shape.sourceShapeId === '3');

    assert.ok(textShape, 'expected the text box shape to be parsed');
    assert.deepEqual(textShape.textInsets, { left: 0, top: 0, right: 0, bottom: 0 });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves inline line-break positions within a paragraph', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'inline-line-breaks.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, {
      extraSlideShapesXml: `      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="52" name="Line break textbox"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="1828800" y="1828800"/>
            <a:ext cx="3657600" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="2000" b="1"/>
              <a:t>Multi-step workflows</a:t>
            </a:r>
            <a:br>
              <a:rPr lang="en-US" sz="2000"/>
            </a:br>
            <a:r>
              <a:rPr lang="en-US" sz="2000"/>
              <a:t>(e.g. research → summarize)</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>`,
    });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const textShape = deck.slides[0].shapes.find((shape) => shape.sourceShapeId === '52');

    assert.ok(textShape, 'expected the line-break text box to be parsed');
    assert.deepEqual(
      textShape.paragraphs?.[0]?.runs?.map((run) => run.text),
      ['Multi-step workflows', '\n', '(e.g. research → summarize)'],
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser maps straight connector arrows to line shapes with end markers', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'straight-connector-arrow.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, {
      extraSlideShapesXml: `      <p:cxnSp>
        <p:nvCxnSpPr>
          <p:cNvPr id="50" name="Straight Arrow Connector 1"/>
          <p:cNvCxnSpPr><a:cxnSpLocks/></p:cNvCxnSpPr>
          <p:nvPr/>
        </p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm flipV="1">
            <a:off x="1828800" y="1828800"/>
            <a:ext cx="3657600" cy="1828800"/>
          </a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="76200">
            <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
            <a:tailEnd type="triangle"/>
          </a:ln>
        </p:spPr>
      </p:cxnSp>`,
    });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const connector = deck.slides[0].shapes.find((shape) => shape.sourceShapeId === '50');

    assert.ok(connector, 'expected the connector shape to be parsed');
    assert.equal(connector.type, 'line');
    assert.equal(connector.lineTail, 'triangle');
    assert.equal(connector.flipV, true);
    assert.equal(connector.border?.colour, '#000000');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves draw order across mixed shape node types', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'mixed-draw-order.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, {
      extraSlideShapesXml: `      <p:cxnSp>
        <p:nvCxnSpPr>
          <p:cNvPr id="60" name="Connector under label"/>
          <p:cNvCxnSpPr><a:cxnSpLocks/></p:cNvCxnSpPr>
          <p:nvPr/>
        </p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="1828800" y="1828800"/>
            <a:ext cx="3048000" cy="0"/>
          </a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="38100"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
        </p:spPr>
      </p:cxnSp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="61" name="Label above connector"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="2286000" y="1600200"/>
            <a:ext cx="2286000" cy="457200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="2000"/>
              <a:t>Visible label</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>`,
    });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const orderedShapeIds = deck.slides[0].shapes
      .filter((shape) => shape.sourceShapeId === '60' || shape.sourceShapeId === '61')
      .map((shape) => shape.sourceShapeId);

    assert.deepEqual(orderedShapeIds, ['60', '61']);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser applies luminance modifiers on scheme colours', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'scheme-colour-luminance-modifier.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, {
      extraSlideShapesXml: `      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="51" name="Luminance text box"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="1828800" y="1828800"/>
            <a:ext cx="3657600" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="2000">
                <a:solidFill>
                  <a:schemeClr val="bg2"><a:lumMod val="25000"/></a:schemeClr>
                </a:solidFill>
              </a:rPr>
              <a:t>Dimmed text</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>`,
    });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const textShape = deck.slides[0].shapes.find((shape) => shape.sourceShapeId === '51');

    assert.ok(textShape, 'expected the luminance text box to be parsed');
    assert.equal(textShape.paragraphs?.[0]?.runs?.[0]?.colour, '#3D3D3D');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves trailing whitespace between adjacent text runs', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'layout-svg-placeholder.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath);

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const titleShape = slide.shapes.find((shape) => shape.sourceShapeId === '2');

    assert.ok(titleShape, 'expected the title placeholder shape to be parsed');
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[0]?.text, 'Welcome to the 1st ');
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[1]?.text, 'InterSystems READY Hackathon');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves leading zeros in text runs', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'layout-svg-placeholder-leading-zero.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, { titleRuns: ['01'] });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const titleShape = slide.shapes.find((shape) => shape.sourceShapeId === '2');

    assert.ok(titleShape, 'expected the title placeholder shape to be parsed');
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[0]?.text, '01');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves inherited placeholder alignment from layout text styles', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'layout-centered-title-alignment.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath);

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const titleShape = slide.shapes.find((shape) => shape.sourceShapeId === '2');

    assert.ok(titleShape, 'expected the title placeholder shape to be parsed');
    assert.equal(titleShape.paragraphs?.[0]?.alignment, 'center');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser resets layout title emphasis when local defRPr omits bold', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'layout-title-emphasis-reset.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, {
      masterTitleStyleXml: `      <a:lvl1pPr algn="l">
        <a:defRPr sz="2800" b="1">
          <a:solidFill><a:srgbClr val="222222"/></a:solidFill>
          <a:latin typeface="Gotham Book"/>
        </a:defRPr>
      </a:lvl1pPr>`,
      layoutTitleStyleXml: `            <a:lvl1pPr algn="ctr">
              <a:defRPr sz="4400">
                <a:solidFill><a:srgbClr val="00B2A9"/></a:solidFill>
                <a:latin typeface="Verdana"/>
              </a:defRPr>
            </a:lvl1pPr>`,
    });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const titleShape = deck.slides[0].shapes.find((shape) => shape.sourceShapeId === '2');

    assert.ok(titleShape, 'expected the title placeholder shape to be parsed');
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[0]?.bold, false);
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[0]?.fontSize, 44);
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[0]?.colour, '#00B2A9');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves paragraph line spacing percentages', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'paragraph-line-spacing.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, {
      extraSlideShapesXml: `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="30" name="Body Placeholder"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="3352800" y="3657600"/>
            <a:ext cx="5486400" cy="1828800"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:lnSpc><a:spcPct val="200000"/></a:lnSpc></a:pPr>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>Agenda item</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>`,
    });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const bodyShape = slide.shapes.find((shape) => shape.sourceShapeId === '30');

    assert.ok(bodyShape, 'expected the body placeholder shape to be parsed');
    assert.equal(bodyShape.paragraphs?.[0]?.lineSpacing, 2);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser resolves tx2 scheme colours through the theme palette', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'layout-scheme-colour-title.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, {
      titleFillXml: '<a:solidFill><a:schemeClr val="tx2"/></a:solidFill>',
    });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const titleShape = slide.shapes.find((shape) => shape.sourceShapeId === '2');

    assert.ok(titleShape, 'expected the title placeholder shape to be parsed');
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[0]?.colour, '#222222');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves grouped layout shapes with small child coordinate systems', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'layout-svg-placeholder-grouped-layout-shape.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath);

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const groupedMark = slide.shapes.find((shape) => shape.id.startsWith('layout-') && shape.sourceShapeId === '20');

    assert.ok(groupedMark, 'expected the grouped layout shape to be rendered');
    assert.ok(groupedMark.width > 0, 'expected grouped layout shape width to remain non-zero');
    assert.ok(groupedMark.height > 0, 'expected grouped layout shape height to remain non-zero');
    assert.equal(groupedMark.fill?.colour, '#2F2A95');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser preserves inherited placeholder font families over theme fallback fonts', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'layout-svg-placeholder.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath);

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const titleShape = slide.shapes.find((shape) => shape.sourceShapeId === '2');

    assert.ok(titleShape, 'expected the title placeholder shape to be parsed');
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[0]?.fontFamily, 'Verdana');
    assert.equal(titleShape.paragraphs?.[0]?.runs?.[1]?.fontFamily, 'Verdana');
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

test('parser does not invent black borders for textbox lines without explicit fill', async () => {
  const { deck, tempRoot } = await parseFixture('MMC_simpliPyTEM_presentation.pptx');

  try {
    const slideOne = deck.slides[0];
    const subtitleShape = slideOne.shapes.find((shape) => shape.sourceShapeId === '199');
    const documentationShape = slideOne.shapes.find((shape) => shape.sourceShapeId === '201');

    assert.ok(subtitleShape, 'expected the subtitle text box to be parsed');
    assert.ok(documentationShape, 'expected the documentation text box to be parsed');
    assert.equal(
      subtitleShape.border,
      undefined,
      'expected inherited textbox defaults without line fill to stay borderless',
    );
    assert.equal(
      documentationShape.border,
      undefined,
      'expected textbox lines without explicit fill to stay borderless',
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser inherits rectangle border width from theme line styles when a:ln omits @w', async () => {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'webpresent-pptx-synthetic-'));
  const deckId = 'deck-test';
  const getDeckDir = (id) => path.join(tempRoot, id);
  const pptxPath = path.join(tempRoot, 'theme-line-style-border-width.pptx');

  try {
    await mkdir(getDeckDir(deckId), { recursive: true });
    await createSyntheticPptx(pptxPath, {
      themeFmtSchemeXml: `    <a:fmtScheme name="Synthetic Format">
      <a:lnStyleLst>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
    </a:fmtScheme>`,
      extraSlideShapesXml: `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="30" name="Border Box"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="914400"/>
            <a:ext cx="2743200" cy="1828800"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln>
            <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
          </a:ln>
        </p:spPr>
        <p:style>
          <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
          <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
          <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
          <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
        </p:style>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p/>
        </p:txBody>
      </p:sp>`,
    });

    const deck = await parsePptx(pptxPath, deckId, getDeckDir);
    const slide = deck.slides[0];
    const borderBox = slide.shapes.find((shape) => shape.sourceShapeId === '30');

    assert.ok(borderBox, 'expected the rectangle shape to be parsed');
    assert.equal(borderBox.border?.colour, '#000000');
    assert.equal(borderBox.border?.width, 1.5);
    assert.equal(borderBox.border?.style, 'solid');
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test('parser ignores media-call timing when computing visible click builds', async () => {
  const { deck, tempRoot } = await parseFixture('MMC_simpliPyTEM_presentation.pptx');

  try {
    const slideTen = deck.slides[9];
    const delayedRevealShape = slideTen.shapes.find((shape) => shape.sourceShapeId === '481');

    assert.ok(delayedRevealShape, 'expected the delayed reveal callout to be parsed');
    assert.equal(
      slideTen.animationStepCount,
      1,
      'expected media playback timing to be ignored so only visible click builds remain',
    );
    assert.equal(
      delayedRevealShape.animationGroup,
      1,
      'expected the visible callout reveal to be assigned to the first actual click build',
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});
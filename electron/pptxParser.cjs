/**
 * PPTX Parser — extracts slide data from .pptx (OOXML) files.
 *
 * Uses adm-zip to read the archive and fast-xml-parser to parse XML.
 * Returns a structured PptxDeckData object that can be rendered as HTML.
 */

const AdmZip = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');
const path = require('node:path');
const fs = require('node:fs/promises');

// ── Constants ────────────────────────────────────────────────────────────────

const EMU_PER_PT = 12700;
const EMU_PER_PX = 9525; // 96 dpi: 914400 EMU/inch ÷ 96 px/inch

function emuToPx(emu) {
  return Math.round((Number(emu) || 0) / EMU_PER_PX);
}

function emuToPt(emu) {
  return Math.round(((Number(emu) || 0) / EMU_PER_PT) * 10) / 10;
}

/** Half-point to pt (font sizes in OOXML are in half-points × 100). */
function hundredthPtToPt(val) {
  return Math.round(((Number(val) || 0) / 100) * 10) / 10;
}

function pickDefined(...values) {
  for (const value of values) {
    if (value !== undefined && value !== null) return value;
  }
  return undefined;
}

function cloneValue(value) {
  if (value === undefined) return undefined;
  return JSON.parse(JSON.stringify(value));
}

function asArray(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function hasAnyValue(obj) {
  return Boolean(obj) && Object.values(obj).some((value) => value !== undefined);
}

// ── XML parser factory ───────────────────────────────────────────────────────

function createXmlParser() {
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    allowBooleanAttributes: true,
    parseAttributeValue: false, // keep as strings to avoid precision loss
    isArray: (name) => {
      // Force arrays for elements that can repeat
      const alwaysArray = new Set([
        'p:sp',
        'p:pic',
        'p:grpSp',
        'p:graphicFrame',
        'p:cxnSp',
        'a:p',
        'a:r',
        'a:br',
        'a:gs',
        'Relationship',
        'a:buNone',
        'a:buChar',
        'a:buAutoNum',
        'p:childTnLst',
        'p:par',
        'p:seq',
        'p:anim',
        'p:animEffect',
        'p:set',
        'p:cTn',
        'p:stCondLst',
        'p:cond',
        'p:sldId',
        'mc:AlternateContent',
        'mc:Choice',
        'mc:Fallback',
      ]);
      return alwaysArray.has(name);
    },
  });
}

// ── Colour helpers ───────────────────────────────────────────────────────────

/** OOXML stores colours as RRGGBB (no #). Normalise to #RRGGBB. */
function normColour(val) {
  if (!val) return undefined;
  const s = String(val).replace(/^#/, '');
  if (/^[0-9a-fA-F]{6}$/.test(s)) return `#${s}`;
  return undefined;
}

const SCHEME_COLOUR_MAP = {
  dk1: 'tx1',
  lt1: 'bg1',
  dk2: 'tx2',
  lt2: 'bg2',
};

function resolveSchemeColour(schemeKey, themeColours) {
  if (!themeColours) return undefined;
  const mapped = SCHEME_COLOUR_MAP[schemeKey] || schemeKey;
  return themeColours[mapped] || themeColours[schemeKey] || undefined;
}

// ── Theme parsing ────────────────────────────────────────────────────────────

function parseTheme(xmlStr, parser) {
  if (!xmlStr) return { colours: {}, defaultFont: undefined };
  const doc = parser.parse(xmlStr);
  const theme = doc?.['a:theme'];
  const elements = theme?.['a:themeElements'];
  const colours = {};
  let defaultFont;

  // Colour scheme
  const clrScheme = elements?.['a:clrScheme'];
  if (clrScheme) {
    for (const [key, val] of Object.entries(clrScheme)) {
      if (key.startsWith('@_')) continue;
      const srgb = val?.['a:srgbClr']?.['@_val'];
      const sysClr = val?.['a:sysClr']?.['@_lastClr'] || val?.['a:sysClr']?.['@_val'];
      const hex = srgb || sysClr;
      if (hex) {
        colours[key.replace('a:', '')] = `#${hex}`;
      }
    }
  }

  // Font scheme
  const fontScheme = elements?.['a:fontScheme'];
  const majorFont = fontScheme?.['a:majorFont']?.['a:latin']?.['@_typeface'];
  const minorFont = fontScheme?.['a:minorFont']?.['a:latin']?.['@_typeface'];
  defaultFont = minorFont || majorFont || undefined;

  return { colours, defaultFont };
}

// ── Relationship parsing ─────────────────────────────────────────────────────

function parseRels(xmlStr, parser) {
  if (!xmlStr) return {};
  const doc = parser.parse(xmlStr);
  const rels = {};
  const relationships = doc?.Relationships?.Relationship;
  if (!relationships) return rels;
  const list = Array.isArray(relationships) ? relationships : [relationships];
  for (const r of list) {
    if (r?.['@_Id'] && r?.['@_Target']) {
      rels[r['@_Id']] = {
        target: r['@_Target'],
        type: r['@_Type'] || '',
      };
    }
  }
  return rels;
}

// ── Presentation.xml parsing (slide order + dimensions) ──────────────────────

function parsePresentationXml(xmlStr, parser) {
  const doc = parser.parse(xmlStr);
  const pres = doc?.['p:presentation'];
  const sldSz = pres?.['p:sldSz'];
  const width = emuToPx(sldSz?.['@_cx']);
  const height = emuToPx(sldSz?.['@_cy']);

  // Slide list (rIds)
  const sldIdLst = pres?.['p:sldIdLst']?.['p:sldId'];
  const slideRIds = [];
  if (sldIdLst) {
    const list = Array.isArray(sldIdLst) ? sldIdLst : [sldIdLst];
    for (const s of list) {
      const rId = s?.['@_r:id'];
      if (rId) slideRIds.push(rId);
    }
  }

  return { width: width || 960, height: height || 540, slideRIds };
}

// ── Fill parsing ─────────────────────────────────────────────────────────────

function parseFill(spPr, slideRels, themeColours) {
  if (!spPr) return undefined;

  // Solid fill
  const solidFill = spPr['a:solidFill'];
  if (solidFill) {
    const colour = extractColour(solidFill, themeColours);
    if (colour) return { type: 'solid', colour };
  }

  // Image fill (blipFill on the shape properties)
  const blipFill = spPr['a:blipFill'];
  if (blipFill) {
    const embed = blipFill?.['a:blip']?.['@_r:embed'];
    if (embed && slideRels[embed]) {
      return { type: 'image', imageTarget: slideRels[embed].target };
    }
  }

  // No fill
  if (spPr['a:noFill'] !== undefined) return { type: 'none' };

  // Gradient fill (simplified — take first and last stops)
  const gradFill = spPr['a:gradFill'];
  if (gradFill) {
    const gsLst = gradFill['a:gsLst']?.['a:gs'];
    if (gsLst) {
      const list = Array.isArray(gsLst) ? gsLst : [gsLst];
      const stops = list.map((gs) => ({
        position: Number(gs['@_pos'] || 0) / 1000, // OOXML positions are in 1/1000ths of %
        colour: extractColour(gs, themeColours) || '#000000',
      }));
      if (stops.length) return { type: 'gradient', gradientStops: stops };
    }
  }

  return undefined;
}

function extractColour(node, themeColours) {
  if (!node) return undefined;
  const srgb = node['a:srgbClr'];
  if (srgb) return normColour(srgb['@_val']);

  const schm = node['a:schemeClr'];
  if (schm) {
    const val = schm['@_val'];
    return resolveSchemeColour(val, themeColours);
  }

  return undefined;
}

// ── Border parsing ───────────────────────────────────────────────────────────

function parseBorder(spPr, themeColours) {
  const ln = spPr?.['a:ln'];
  if (!ln) return undefined;
  if (ln['a:noFill'] !== undefined) return undefined;

  const solidFill = ln['a:solidFill'];
  if (!solidFill) return undefined;

  const widthEmu = Number(ln['@_w'] || 0);
  if (!widthEmu) return undefined;
  const widthPt = emuToPt(widthEmu);

  const colour = extractColour(solidFill, themeColours);
  if (!colour) return undefined;

  const prstDash = ln['a:prstDash']?.['@_val'];
  let style = 'solid';
  if (prstDash === 'dash' || prstDash === 'lgDash') style = 'dashed';
  if (prstDash === 'dot' || prstDash === 'sysDot') style = 'dotted';

  return { width: widthPt, colour, style };
}

// ── Text parsing ─────────────────────────────────────────────────────────────

function parseRunStyleDefaults(rPr, themeColours, defaultFont) {
  const runDefaults = {};
  if (!rPr) {
    if (defaultFont) runDefaults.fontFamily = defaultFont;
    return runDefaults;
  }

  if (rPr['@_b'] !== undefined) {
    runDefaults.bold = rPr['@_b'] === '1' || rPr['@_b'] === 'true';
  }
  if (rPr['@_i'] !== undefined) {
    runDefaults.italic = rPr['@_i'] === '1' || rPr['@_i'] === 'true';
  }
  if (rPr['@_u'] !== undefined) {
    runDefaults.underline = rPr['@_u'] !== 'none';
  }

  const sz = rPr['@_sz'];
  if (sz) runDefaults.fontSize = hundredthPtToPt(sz);

  const typeface = rPr['a:latin']?.['@_typeface'];
  if (typeface) runDefaults.fontFamily = typeface;
  else if (defaultFont) runDefaults.fontFamily = defaultFont;

  const solidFill = rPr['a:solidFill'];
  if (solidFill) {
    runDefaults.colour = extractColour(solidFill, themeColours);
  }

  return runDefaults;
}

function mergeRunDefaults(base, override) {
  if (!base && !override) return undefined;
  const merged = {
    bold: pickDefined(override?.bold, base?.bold),
    italic: pickDefined(override?.italic, base?.italic),
    underline: pickDefined(override?.underline, base?.underline),
    fontSize: pickDefined(override?.fontSize, base?.fontSize),
    fontFamily: pickDefined(override?.fontFamily, base?.fontFamily),
    colour: pickDefined(override?.colour, base?.colour),
  };
  return hasAnyValue(merged) ? merged : undefined;
}

function applyRunDefaults(run, defaults, defaultFont) {
  const merged = { ...run };
  if (defaults) {
    for (const key of ['bold', 'italic', 'underline', 'fontSize', 'fontFamily', 'colour']) {
      if (merged[key] === undefined && defaults[key] !== undefined) {
        merged[key] = defaults[key];
      }
    }
  }
  if (merged.fontFamily === undefined && defaultFont) {
    merged.fontFamily = defaultFont;
  }
  return merged;
}

function parseParagraphDefaults(pPrNode, themeColours, defaultFont, level) {
  if (!pPrNode) return undefined;

  const defaults = {};
  if (level !== undefined) defaults.level = level;

  const algn = pPrNode['@_algn'];
  if (algn === 'ctr') defaults.alignment = 'center';
  else if (algn === 'r') defaults.alignment = 'right';
  else if (algn === 'just') defaults.alignment = 'justify';
  else if (algn === 'l') defaults.alignment = 'left';

  if (pPrNode['a:buChar']) {
    defaults.bulletType = 'bullet';
    defaults.bulletChar = pPrNode['a:buChar']['@_char'] || undefined;
  } else if (pPrNode['a:buAutoNum']) {
    defaults.bulletType = 'numbered';
  } else if (pPrNode['a:buNone'] !== undefined) {
    defaults.bulletType = 'none';
  }

  const runDefaults = mergeRunDefaults(
    parseRunStyleDefaults(pPrNode['a:endParaRPr'], themeColours, defaultFont),
    parseRunStyleDefaults(pPrNode['a:defRPr'], themeColours, defaultFont),
  );
  if (runDefaults) defaults.runDefaults = runDefaults;

  return hasAnyValue(defaults) ? defaults : undefined;
}

function mergeParagraphDefaults(base, override) {
  if (!base && !override) return undefined;
  const merged = {
    alignment: pickDefined(override?.alignment, base?.alignment),
    bulletType: pickDefined(override?.bulletType, base?.bulletType),
    bulletChar: pickDefined(override?.bulletChar, base?.bulletChar),
    level: pickDefined(override?.level, base?.level),
    runDefaults: mergeRunDefaults(base?.runDefaults, override?.runDefaults),
  };
  return hasAnyValue(merged) ? merged : undefined;
}

function parseListStyleDefaults(styleNode, themeColours, defaultFont) {
  const levels = {};
  if (!styleNode) return levels;

  const defaultLevel = parseParagraphDefaults(styleNode['a:defPPr'], themeColours, defaultFont, 0);
  if (defaultLevel) levels[0] = defaultLevel;

  for (let level = 0; level < 9; level++) {
    const node = styleNode[`a:lvl${level + 1}pPr`];
    const parsed = parseParagraphDefaults(node, themeColours, defaultFont, level);
    if (parsed) {
      levels[level] = mergeParagraphDefaults(levels[level], parsed);
    }
  }

  return levels;
}

function mergeTextStyleMaps(base, override) {
  const merged = {};
  for (let level = 0; level < 9; level++) {
    const mergedLevel = mergeParagraphDefaults(base?.[level], override?.[level]);
    if (mergedLevel) merged[level] = mergedLevel;
  }
  return merged;
}

function resolveParagraphDefaults(textStyleMap, level) {
  return textStyleMap?.[level] || textStyleMap?.[0];
}

function mergeParagraphRunDefaults(paragraphs, textStyleMap, defaultFont) {
  if (!paragraphs?.length) return paragraphs;

  return paragraphs.map((paragraph, index) => {
    const paragraphDefaults = resolveParagraphDefaults(textStyleMap, paragraph.level || 0);
    const templateParagraph = undefined;
    const mergedRuns = (paragraph.runs || []).map((run) => applyRunDefaults(run, paragraphDefaults?.runDefaults, defaultFont));
    return {
      ...paragraph,
      alignment: pickDefined(paragraph.alignment, paragraphDefaults?.alignment),
      bulletType: pickDefined(paragraph.bulletType, paragraphDefaults?.bulletType),
      bulletChar: pickDefined(paragraph.bulletChar, paragraphDefaults?.bulletChar),
      level: pickDefined(paragraph.level, paragraphDefaults?.level, 0),
      runs: mergedRuns.length ? mergedRuns : (templateParagraph ? templateParagraph.runs : mergedRuns),
    };
  });
}

function parseMasterTextStyles(txStyles, themeColours, defaultFont) {
  return {
    title: parseListStyleDefaults(txStyles?.['p:titleStyle'], themeColours, defaultFont),
    body: parseListStyleDefaults(txStyles?.['p:bodyStyle'], themeColours, defaultFont),
    other: parseListStyleDefaults(txStyles?.['p:otherStyle'], themeColours, defaultFont),
  };
}

function getPlaceholderInfo(node) {
  const ph =
    node?.['p:nvSpPr']?.['p:nvPr']?.['p:ph'] ||
    node?.['p:nvPicPr']?.['p:nvPr']?.['p:ph'] ||
    node?.['p:nvCxnSpPr']?.['p:nvPr']?.['p:ph'];

  if (!ph) return undefined;

  return {
    type: ph['@_type'] || 'body',
    idx: ph['@_idx'] != null ? String(ph['@_idx']) : undefined,
    orient: ph['@_orient'] || undefined,
    size: ph['@_sz'] || undefined,
  };
}

function scorePlaceholderMatch(source, candidate) {
  if (!source || !candidate) return -1;

  const looseIndexTypes = new Set(['dt', 'ftr', 'hdr', 'sldNum']);
  const allowsLooseIndexMatch = looseIndexTypes.has(source.type) && looseIndexTypes.has(candidate.type);

  let score = 0;
  if (source.idx !== undefined || candidate.idx !== undefined) {
    if (source.idx === candidate.idx) score += 12;
    else if (source.idx !== undefined && candidate.idx !== undefined && !allowsLooseIndexMatch) return -1;
  }

  if (source.type !== undefined || candidate.type !== undefined) {
    if (source.type === candidate.type) {
      score += 8;
    } else if (source.type && candidate.type) {
      const compatibleBodyTypes = new Set(['body', 'obj', 'subTitle']);
      if (compatibleBodyTypes.has(source.type) && compatibleBodyTypes.has(candidate.type)) {
        score += 3;
      } else {
        return -1;
      }
    }
  }

  if (source.orient !== undefined || candidate.orient !== undefined) {
    if (source.orient === candidate.orient) score += 2;
    else if (source.orient !== undefined && candidate.orient !== undefined) return -1;
  }

  if (source.size !== undefined || candidate.size !== undefined) {
    if (source.size === candidate.size) score += 1;
    else if (source.size !== undefined && candidate.size !== undefined) return -1;
  }

  return score;
}

function findPlaceholderTemplate(shape, templates) {
  if (!shape?.placeholder || !templates?.length) return undefined;

  let bestMatch;
  let bestScore = -1;
  for (const template of templates) {
    const score = scorePlaceholderMatch(shape.placeholder, template?.placeholder);
    if (score > bestScore) {
      bestScore = score;
      bestMatch = template;
    }
  }

  return bestScore >= 0 ? bestMatch : undefined;
}

function getTextStyleMapForPlaceholder(placeholder, masterTextStyles) {
  if (!masterTextStyles || !placeholder) return undefined;

  switch (placeholder.type) {
    case 'title':
    case 'ctrTitle':
      return masterTextStyles.title;
    case 'body':
    case 'obj':
    case 'subTitle':
      return masterTextStyles.body;
    default:
      return masterTextStyles.other;
  }
}

function getTemplateRunDefaults(paragraphs, index) {
  if (!paragraphs?.length) return undefined;
  const templateParagraph = paragraphs[index] || paragraphs[paragraphs.length - 1];
  if (!templateParagraph?.runs?.length) return undefined;
  const candidate = templateParagraph.runs.find((run) => run.text !== '\n') || templateParagraph.runs[0];
  if (!candidate) return undefined;
  const { text: _text, ...defaults } = candidate;
  return hasAnyValue(defaults) ? defaults : undefined;
}

function mergeParagraphsFromTemplate(paragraphs, templateParagraphs, masterTextStyles, placeholder, defaultFont) {
  if (!paragraphs?.length) return paragraphs;

  const masterStyleMap = getTextStyleMapForPlaceholder(placeholder, masterTextStyles);
  return paragraphs.map((paragraph, index) => {
    const templateParagraph = templateParagraphs?.[index] || templateParagraphs?.[templateParagraphs.length - 1];
    const paragraphDefaults = mergeParagraphDefaults(
      resolveParagraphDefaults(masterStyleMap, paragraph.level || 0),
      templateParagraph
        ? {
            alignment: templateParagraph.alignment,
            bulletType: templateParagraph.bulletType,
            bulletChar: templateParagraph.bulletChar,
            level: templateParagraph.level,
            runDefaults: getTemplateRunDefaults(templateParagraphs, index),
          }
        : undefined,
    );

    return {
      ...paragraph,
      alignment: pickDefined(paragraph.alignment, paragraphDefaults?.alignment),
      bulletType: pickDefined(paragraph.bulletType, paragraphDefaults?.bulletType),
      bulletChar: pickDefined(paragraph.bulletChar, paragraphDefaults?.bulletChar),
      level: pickDefined(paragraph.level, paragraphDefaults?.level, 0),
      runs: (paragraph.runs || []).map((run) => applyRunDefaults(run, paragraphDefaults?.runDefaults, defaultFont)),
    };
  });
}

function parseTextBody(txBody, themeColours, defaultFont, inheritedTextStyles) {
  if (!txBody) return [];
  const paragraphs = [];
  const pList = txBody['a:p'];
  if (!pList) return [];
  const pArray = Array.isArray(pList) ? pList : [pList];
  const localTextStyles = parseListStyleDefaults(txBody['a:lstStyle'], themeColours, defaultFont);
  const textStyleMap = mergeTextStyleMaps(inheritedTextStyles, localTextStyles);

  for (const p of pArray) {
    const pPr = p['a:pPr'];
    const level = Number(pPr?.['@_lvl'] || 0);
    const paragraphDefaults = mergeParagraphDefaults(
      resolveParagraphDefaults(textStyleMap, level),
      parseParagraphDefaults(pPr, themeColours, defaultFont, level),
    );

    let alignment = paragraphDefaults?.alignment || 'left';
    let bulletType = paragraphDefaults?.bulletType || 'none';
    let bulletChar = paragraphDefaults?.bulletChar;

    const runs = [];
    const rList = p['a:r'];
    const fldList = p['a:fld'];

    const paragraphRunDefaults = mergeRunDefaults(
      paragraphDefaults?.runDefaults,
      parseRunStyleDefaults(pPr?.['a:defRPr'], themeColours, defaultFont),
    );

    const emptyParagraphRunDefaults = mergeRunDefaults(
      paragraphRunDefaults,
      parseRunStyleDefaults(p['a:endParaRPr'], themeColours, defaultFont),
    );

    if (rList) {
      const rArray = Array.isArray(rList) ? rList : [rList];
      for (const r of rArray) {
        const rPr = r['a:rPr'];
        const text = r['a:t'] != null ? String(r['a:t']) : '';
        const runDefaults = mergeRunDefaults(paragraphRunDefaults, parseRunStyleDefaults(rPr, themeColours, defaultFont));
        runs.push(applyRunDefaults({ text }, runDefaults, defaultFont));
      }
    }

    if (fldList) {
      const fldArray = Array.isArray(fldList) ? fldList : [fldList];
      for (const fld of fldArray) {
        const rPr = fld['a:rPr'];
        const text = fld['a:t'] != null ? String(fld['a:t']) : '';
        const runDefaults = mergeRunDefaults(paragraphRunDefaults, parseRunStyleDefaults(rPr, themeColours, defaultFont));
        runs.push(applyRunDefaults({ text }, runDefaults, defaultFont));
      }
    }

    // Handle line breaks
    const brList = p['a:br'];
    if (brList) {
      const brArray = Array.isArray(brList) ? brList : [brList];
      for (const _br of brArray) {
        runs.push({ text: '\n' });
      }
    }

    // If no runs but has endParaRPr, it's an empty paragraph (still emit it for spacing)
    if (runs.length === 0) {
      runs.push(applyRunDefaults({ text: '' }, emptyParagraphRunDefaults, defaultFont));
    }

    paragraphs.push({ runs, alignment, bulletType, bulletChar, level, animationGroup: 0 });
  }

  return paragraphs;
}

// ── Shape parsing ────────────────────────────────────────────────────────────

function parseShapeTransform(spNode) {
  // Try xfrm in spPr first, then in grpSpPr
  const xfrm =
    spNode?.['p:spPr']?.['a:xfrm'] ||
    spNode?.['p:grpSpPr']?.['a:xfrm'] ||
    spNode?.['p:spPr']?.['a:off']?.['..'] || // shouldn't happen but defensive
    null;

  if (!xfrm) {
    // Try direct off/ext in spPr
    const off = spNode?.['p:spPr']?.['a:off'];
    const ext = spNode?.['p:spPr']?.['a:ext'];
    if (off || ext) {
      return {
        x: emuToPx(off?.['@_x']),
        y: emuToPx(off?.['@_y']),
        width: emuToPx(ext?.['@_cx']),
        height: emuToPx(ext?.['@_cy']),
        rotation: 0,
      };
    }
    return { x: 0, y: 0, width: 0, height: 0, rotation: 0 };
  }

  const off = xfrm['a:off'];
  const ext = xfrm['a:ext'];
  const rot = Number(xfrm['@_rot'] || 0) / 60000; // 60000ths of a degree → degrees

  return {
    x: emuToPx(off?.['@_x']),
    y: emuToPx(off?.['@_y']),
    width: emuToPx(ext?.['@_cx']),
    height: emuToPx(ext?.['@_cy']),
    rotation: rot,
  };
}

function applyGroupTransform(transform, groupContext) {
  if (!groupContext) return transform;

  const scaleX = groupContext.childWidth ? groupContext.width / groupContext.childWidth : 1;
  const scaleY = groupContext.childHeight ? groupContext.height / groupContext.childHeight : 1;

  return {
    x: Math.round(groupContext.x + (transform.x - groupContext.childX) * scaleX),
    y: Math.round(groupContext.y + (transform.y - groupContext.childY) * scaleY),
    width: Math.round(transform.width * scaleX),
    height: Math.round(transform.height * scaleY),
    rotation: (transform.rotation || 0) + (groupContext.rotation || 0),
  };
}

function parseGroupContext(grpNode, parentGroupContext) {
  const xfrm = grpNode?.['p:grpSpPr']?.['a:xfrm'];
  if (!xfrm) return parentGroupContext;

  const ownContext = {
    x: emuToPx(xfrm?.['a:off']?.['@_x']),
    y: emuToPx(xfrm?.['a:off']?.['@_y']),
    width: emuToPx(xfrm?.['a:ext']?.['@_cx']),
    height: emuToPx(xfrm?.['a:ext']?.['@_cy']),
    childX: emuToPx(xfrm?.['a:chOff']?.['@_x']),
    childY: emuToPx(xfrm?.['a:chOff']?.['@_y']),
    childWidth: emuToPx(xfrm?.['a:chExt']?.['@_cx']) || emuToPx(xfrm?.['a:ext']?.['@_cx']),
    childHeight: emuToPx(xfrm?.['a:chExt']?.['@_cy']) || emuToPx(xfrm?.['a:ext']?.['@_cy']),
    rotation: Number(xfrm?.['@_rot'] || 0) / 60000,
  };

  if (!parentGroupContext) {
    return ownContext;
  }

  const transformedBounds = applyGroupTransform(ownContext, parentGroupContext);
  return {
    ...ownContext,
    x: transformedBounds.x,
    y: transformedBounds.y,
    width: transformedBounds.width,
    height: transformedBounds.height,
    rotation: transformedBounds.rotation,
  };
}

function getGroupNodeId(grpNode) {
  return String(grpNode?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id'] || '');
}

function getPresetGeometry(spPr) {
  const prst = spPr?.['a:prstGeom']?.['@_prst'];
  if (!prst) return 'rect';
  if (prst === 'ellipse') return 'ellipse';
  if (prst === 'roundRect') return 'roundRect';
  if (prst === 'line') return 'line';
  return 'rect';
}

function parseFlipFlags(xfrm) {
  if (!xfrm) return { flipH: false, flipV: false };
  return {
    flipH: xfrm['@_flipH'] === '1' || xfrm['@_flipH'] === 'true',
    flipV: xfrm['@_flipV'] === '1' || xfrm['@_flipV'] === 'true',
  };
}

function parseImageCrop(blipFill) {
  const srcRect = blipFill?.['a:srcRect'];
  if (!srcRect) return undefined;

  const crop = {
    left: Number(srcRect['@_l'] || 0) / 100000,
    top: Number(srcRect['@_t'] || 0) / 100000,
    right: Number(srcRect['@_r'] || 0) / 100000,
    bottom: Number(srcRect['@_b'] || 0) / 100000,
  };

  return Object.values(crop).some((value) => value > 0) ? crop : undefined;
}

function parseShape(spNode, slideRels, themeColours, defaultFont, groupContext, ancestorGroupIds = []) {
  const nvSpPr = spNode['p:nvSpPr'] || spNode['p:nvCxnSpPr'];
  const cNvPr = nvSpPr?.['p:cNvPr'];
  const id = String(cNvPr?.['@_id'] || '');
  const name = cNvPr?.['@_name'] || '';
  const placeholder = getPlaceholderInfo(spNode);

  const spPr = spNode['p:spPr'];
  const transform = applyGroupTransform(parseShapeTransform(spNode), groupContext);
  const shapeType = getPresetGeometry(spPr);
  const fill = parseFill(spPr, slideRels, themeColours);
  const border = parseBorder(spPr, themeColours);
  const flip = parseFlipFlags(spPr?.['a:xfrm']);

  const txBody = spNode['p:txBody'];
  const paragraphs = parseTextBody(txBody, themeColours, defaultFont);

  // Vertical text alignment from bodyPr
  let verticalAlign;
  const bodyPr = txBody?.['a:bodyPr'];
  if (bodyPr) {
    const anchor = bodyPr['@_anchor'];
    if (anchor === 'ctr' || anchor === 'mid') verticalAlign = 'middle';
    else if (anchor === 'b') verticalAlign = 'bottom';
    else if (anchor === 't') verticalAlign = 'top';
  }

  let cornerRadius;
  if (shapeType === 'roundRect') {
    // Default corner radius ~16.67% of min dimension
    cornerRadius = Math.round(Math.min(transform.width, transform.height) * 0.167);
  }

  return {
    id: `shape-${id}`,
    name,
    type: shapeType,
    ...transform,
    fill,
    border,
    paragraphs: paragraphs.length ? paragraphs : undefined,
    verticalAlign,
    cornerRadius,
    placeholder,
    sourceShapeId: id,
    ancestorGroupIds,
    flipH: flip.flipH,
    flipV: flip.flipV,
    animationGroup: 0, // default: always visible; may be overridden by animation parsing
    animationEffect: 'appear',
  };
}

function inheritShapeFromTemplate(shape, template, masterTextStyles, defaultFont) {
  if (!shape) return shape;

  const mergedShape = { ...shape };
  const sourceTemplate = template ? cloneValue(template) : undefined;

  if (sourceTemplate) {
    const needsTemplateTransform = mergedShape.width === 0 && mergedShape.height === 0;
    for (const key of ['x', 'y', 'width', 'height', 'rotation']) {
      if (needsTemplateTransform || mergedShape[key] === undefined) {
        if (sourceTemplate[key] !== undefined) mergedShape[key] = sourceTemplate[key];
      }
    }
    if (mergedShape.fill === undefined && sourceTemplate.fill !== undefined) mergedShape.fill = sourceTemplate.fill;
    if (mergedShape.border === undefined && sourceTemplate.border !== undefined) mergedShape.border = sourceTemplate.border;
    if (mergedShape.verticalAlign === undefined && sourceTemplate.verticalAlign !== undefined) mergedShape.verticalAlign = sourceTemplate.verticalAlign;
    if (mergedShape.cornerRadius === undefined && sourceTemplate.cornerRadius !== undefined) mergedShape.cornerRadius = sourceTemplate.cornerRadius;
  }

  if (mergedShape.paragraphs?.length) {
    mergedShape.paragraphs = mergeParagraphsFromTemplate(
      mergedShape.paragraphs,
      sourceTemplate?.paragraphs,
      masterTextStyles,
      mergedShape.placeholder || sourceTemplate?.placeholder,
      defaultFont,
    );
  }

  return mergedShape;
}

function parsePicture(picNode, slideRels, themeColours, groupContext, ancestorGroupIds = []) {
  const nvPicPr = picNode['p:nvPicPr'];
  const cNvPr = nvPicPr?.['p:cNvPr'];
  const id = String(cNvPr?.['@_id'] || '');
  const name = cNvPr?.['@_name'] || '';

  const blipFill = picNode['p:blipFill'];
  const embed = blipFill?.['a:blip']?.['@_r:embed'];
  let imageTarget;
  if (embed && slideRels[embed]) {
    imageTarget = slideRels[embed].target;
  }

  const spPr = picNode['p:spPr'];
  const xfrm = spPr?.['a:xfrm'];
  const off = xfrm?.['a:off'];
  const ext = xfrm?.['a:ext'];
  const rot = Number(xfrm?.['@_rot'] || 0) / 60000;
  const flip = parseFlipFlags(xfrm);
  const imageCrop = parseImageCrop(blipFill);
  const transform = applyGroupTransform({
    x: emuToPx(off?.['@_x']),
    y: emuToPx(off?.['@_y']),
    width: emuToPx(ext?.['@_cx']),
    height: emuToPx(ext?.['@_cy']),
    rotation: rot,
  }, groupContext);

  return {
    id: `pic-${id}`,
    name,
    type: 'image',
    ...transform,
    imageTarget,
    imageCrop,
    flipH: flip.flipH,
    flipV: flip.flipV,
    sourceShapeId: id,
    ancestorGroupIds,
    animationGroup: 0,
    animationEffect: 'appear',
  };
}

function collectDrawableNodes(spTree, slideRels, themeColours, defaultFont, groupContext, ancestorGroupIds = []) {
  const shapes = [];
  if (!spTree) return shapes;

  const spNodes = spTree['p:sp'];
  if (spNodes) {
    const list = Array.isArray(spNodes) ? spNodes : [spNodes];
    for (const sp of list) {
      shapes.push(parseShape(sp, slideRels, themeColours, defaultFont, groupContext, ancestorGroupIds));
    }
  }

  const picNodes = spTree['p:pic'];
  if (picNodes) {
    const list = Array.isArray(picNodes) ? picNodes : [picNodes];
    for (const pic of list) {
      shapes.push(parsePicture(pic, slideRels, themeColours, groupContext, ancestorGroupIds));
    }
  }

  const cxnSpNodes = spTree['p:cxnSp'];
  if (cxnSpNodes) {
    const list = Array.isArray(cxnSpNodes) ? cxnSpNodes : [cxnSpNodes];
    for (const cxn of list) {
      shapes.push(parseShape(cxn, slideRels, themeColours, defaultFont, groupContext, ancestorGroupIds));
    }
  }

  const grpSpNodes = spTree['p:grpSp'];
  if (grpSpNodes) {
    const list = Array.isArray(grpSpNodes) ? grpSpNodes : [grpSpNodes];
    for (const grp of list) {
      const groupId = getGroupNodeId(grp);
      const nextAncestorGroupIds = groupId ? [...ancestorGroupIds, groupId] : ancestorGroupIds;
      const nextGroupContext = parseGroupContext(grp, groupContext);
      shapes.push(...collectDrawableNodes(grp, slideRels, themeColours, defaultFont, nextGroupContext, nextAncestorGroupIds));
    }
  }

  const alternateContentNodes = [
    ...asArray(spTree['mc:AlternateContent']),
    ...asArray(spTree.AlternateContent),
  ];
  for (const alternateContent of alternateContentNodes) {
    const fallback = asArray(alternateContent?.['mc:Fallback'])[0] || asArray(alternateContent?.Fallback)[0];
    const choice = asArray(alternateContent?.['mc:Choice'])[0] || asArray(alternateContent?.Choice)[0];
    const drawableContainer = fallback || choice;
    if (drawableContainer) {
      shapes.push(...collectDrawableNodes(drawableContainer, slideRels, themeColours, defaultFont, groupContext, ancestorGroupIds));
    }
  }

  return shapes;
}

// ── Animation parsing ────────────────────────────────────────────────────────

/**
 * Parses the <p:timing> element to figure out which shapes appear on which
 * click. Returns a Map<shapeTargetId, { group: number, effect: string }>.
 *
 * OOXML animation model is deeply nested; we take a practical approach:
 * look for sequences of <p:par> inside <p:seq> (the click sequence),
 * and within each par look for target shape references.
 */
function parseAnimations(timingNode) {
  const animMap = new Map(); // spTarget (string ID) → { group, effect }
  if (!timingNode) return animMap;

  const buildGroupMap = new Map();
  const paragraphBuildTargets = new Set();
  let nextBuildGroup = 1;

  function getBuildGroup(rawGroup) {
    const key = String(rawGroup || nextBuildGroup);
    if (!buildGroupMap.has(key)) {
      buildGroupMap.set(key, nextBuildGroup);
      nextBuildGroup += 1;
    }
    return buildGroupMap.get(key);
  }

  const buildList = timingNode['p:bldLst'];
  const buildNodes = buildList?.['p:bldP'];
  if (buildNodes) {
    const list = Array.isArray(buildNodes) ? buildNodes : [buildNodes];
    for (const build of list) {
      const target = String(build?.['@_spid'] || '');
      if (!target) continue;
      if (build?.['@_build'] === 'p') {
        paragraphBuildTargets.add(target);
      }
      animMap.set(target, {
        group: getBuildGroup(build?.['@_grpId']),
        effect: 'appear',
        paragraphGroups: [],
      });
    }
  }

  // Navigate into the timing tree
  const tnLst = timingNode['p:tnLst'];
  if (!tnLst) return animMap;

  const rootPars = Array.isArray(tnLst['p:par']) ? tnLst['p:par'] : tnLst['p:par'] ? [tnLst['p:par']] : [];
  if (!rootPars.length) return animMap;

  for (const rootPar of rootPars) {
    walkForSequences(rootPar, animMap, paragraphBuildTargets);
  }

  return animMap;
}

function firstNode(node) {
  return Array.isArray(node) ? node[0] : node;
}

function walkForSequences(node, animMap, paragraphBuildTargets) {
  if (!node) return;

  const cTn = node['p:cTn'] || (Array.isArray(node) ? null : node);
  if (!cTn) return;

  const cTnObj = Array.isArray(cTn) ? cTn[0] : cTn;

  const childTnLst = firstNode(cTnObj?.['p:childTnLst']);
  if (!childTnLst) return;

  // Look for seq (click sequence container)
  const seqNodes = childTnLst['p:seq'];
  if (seqNodes) {
    const seqList = Array.isArray(seqNodes) ? seqNodes : [seqNodes];
    for (const seq of seqList) {
      parseClickSequence(seq, animMap, paragraphBuildTargets);
    }
  }

  // Recurse into par nodes
  const parNodes = childTnLst['p:par'];
  if (parNodes) {
    const parList = Array.isArray(parNodes) ? parNodes : [parNodes];
    for (const par of parList) {
      walkForSequences(par, animMap, paragraphBuildTargets);
    }
  }
}

function parseClickSequence(seqNode, animMap, paragraphBuildTargets) {
  if (!seqNode) return;

  const cTn = seqNode['p:cTn'];
  const cTnObj = Array.isArray(cTn) ? cTn[0] : cTn;
  if (cTnObj?.['@_nodeType'] && cTnObj['@_nodeType'] !== 'mainSeq') return;
  const childTnLst = firstNode(cTnObj?.['p:childTnLst']);
  if (!childTnLst) return;

  const parNodes = childTnLst['p:par'];
  if (!parNodes) return;
  const parList = Array.isArray(parNodes) ? parNodes : [parNodes];

  let clickIndex = 0;
  for (const par of parList) {
    const targets = new Map();
    collectTargetsFromPar(par, targets);
    if (!targets.size) continue;
    clickIndex++;
    for (const [target, effect] of targets.entries()) {
      registerAnimationTarget(target, clickIndex, effect, animMap, paragraphBuildTargets);
    }
  }
}

function collectTargetsFromPar(node, targets) {
  if (!node) return;
  const cTn = node['p:cTn'];
  const cTnObj = Array.isArray(cTn) ? cTn[0] : cTn;

  const childTnLst = firstNode(cTnObj?.['p:childTnLst']);
  if (childTnLst) {
    const parNodes = childTnLst['p:par'];
    if (parNodes) {
      const parList = Array.isArray(parNodes) ? parNodes : [parNodes];
      for (const p of parList) {
        collectTargetsFromPar(p, targets);
      }
    }

    const setNodes = childTnLst['p:set'];
    if (setNodes) {
      const list = Array.isArray(setNodes) ? setNodes : [setNodes];
      for (const s of list) {
        const target = extractAnimTarget(s);
        if (target) {
          registerCollectedTarget(target, 'appear', targets);
        }
      }
    }

    const animEffectNodes = childTnLst['p:animEffect'];
    if (animEffectNodes) {
      const list = Array.isArray(animEffectNodes) ? animEffectNodes : [animEffectNodes];
      for (const ae of list) {
        const target = extractAnimTarget(ae);
        const transition = ae['@_transition'];
        const filter = ae['@_filter'];
        let effect = 'fade';
        if (filter?.includes('wipe')) effect = 'fly-left';
        if (transition === 'out') effect = 'fade'; // exit animations
        if (target) {
          registerCollectedTarget(target, effect, targets);
        }
      }
    }

    const animNodes = childTnLst['p:anim'];
    if (animNodes) {
      const list = Array.isArray(animNodes) ? animNodes : [animNodes];
      for (const a of list) {
        const target = extractAnimTarget(a);
        if (target) {
          registerCollectedTarget(target, 'fly-left', targets);
        }
      }
    }
  }

  const directTarget = extractAnimTarget(node);
  if (directTarget) {
    registerCollectedTarget(directTarget, 'appear', targets);
  }
}

function registerCollectedTarget(target, effect, targets) {
  if (!target) return;
  if (!targets.has(target)) {
    targets.set(target, effect || 'appear');
  }
}

function registerAnimationTarget(target, clickGroup, effect, animMap, paragraphBuildTargets) {
  const existing = animMap.get(target);
  const nextEffect = effect || existing?.effect || 'appear';

  if (paragraphBuildTargets.has(target)) {
    const paragraphGroups = Array.isArray(existing?.paragraphGroups) ? [...existing.paragraphGroups] : [];
    if (!paragraphGroups.includes(clickGroup)) {
      paragraphGroups.push(clickGroup);
    }
    animMap.set(target, {
      group: Math.min(existing?.group ?? clickGroup, clickGroup),
      effect: existing?.effect || nextEffect,
      paragraphGroups,
    });
    return;
  }

  animMap.set(target, {
    group: clickGroup,
    effect: existing?.effect === 'appear' ? nextEffect : existing?.effect || nextEffect,
    paragraphGroups: existing?.paragraphGroups || [],
  });
}

function extractAnimTarget(node) {
  const cBhvr = node?.['p:cBhvr'];
  const tgtEl = cBhvr?.['p:tgtEl'];
  const spTgt = tgtEl?.['p:spTgt'];
  if (spTgt) {
    return spTgt['@_spid'];
  }
  // Also check inside cTn → stCondLst → cond → tgtEl
  const cTn = node?.['p:cTn'];
  const cTnObj = Array.isArray(cTn) ? cTn[0] : cTn;
  const stCondLst = firstNode(cTnObj?.['p:stCondLst']);
  if (stCondLst) {
    const cond = stCondLst['p:cond'];
    const condList = Array.isArray(cond) ? cond : cond ? [cond] : [];
    for (const c of condList) {
      const sp = c?.['p:tgtEl']?.['p:spTgt'];
      if (sp) return sp['@_spid'];
    }
  }
  return undefined;
}

// ── Slide background parsing ─────────────────────────────────────────────────

function parseSlideBackground(bgNode, slideRels, themeColours) {
  if (!bgNode) return undefined;

  const bgPr = bgNode['p:bgPr'];
  if (bgPr) {
    const fill = parseFill(bgPr, slideRels, themeColours);
    if (fill) return fill;
  }

  const bgRef = bgNode['p:bgRef'];
  if (bgRef) {
    const colour = extractColour(bgRef, themeColours);
    if (colour) return { type: 'solid', colour };
  }

  return undefined;
}

// ── Layout / Master shape & background parsing ──────────────────────────────

/**
 * Parse shapes from a slide layout or slide master spTree.
 * These are "background" shapes that render behind slide content.
 * Placeholder shapes are collected for inheritance only and are not rendered
 * directly, since PowerPoint/Keynote treat them as template slots.
 */
function parseLayoutMasterShapes(spTree, rels, themeColours, defaultFont) {
  const shapes = [];
  const placeholders = [];
  if (!spTree) return { shapes, placeholders };

  const drawableNodes = collectDrawableNodes(spTree, rels, themeColours, defaultFont);
  for (const shape of drawableNodes) {
    if (shape.placeholder) {
      placeholders.push(shape);
      continue;
    }
    if (shape.width > 0 || shape.height > 0) {
      shapes.push(shape);
    }
  }

  return { shapes, placeholders };
}

/**
 * Check if a txBody has any non-empty text runs.
 */
function hasNonEmptyText(txBody) {
  if (!txBody) return false;
  const pList = txBody['a:p'];
  if (!pList) return false;
  const pArray = Array.isArray(pList) ? pList : [pList];
  for (const p of pArray) {
    const rList = p['a:r'];
    if (!rList) continue;
    const rArray = Array.isArray(rList) ? rList : [rList];
    for (const r of rArray) {
      const text = r['a:t'];
      if (text != null && String(text).trim().length > 0) return true;
    }
  }
  return false;
}

// ── Slide notes parsing ──────────────────────────────────────────────────────

function parseNotes(notesXmlStr, parser) {
  if (!notesXmlStr) return undefined;
  try {
    const doc = parser.parse(notesXmlStr);
    const cSld = doc?.['p:notes']?.['p:cSld'];
    const spTree = cSld?.['p:spTree'];
    if (!spTree) return undefined;
    const spNodes = spTree['p:sp'];
    if (!spNodes) return undefined;
    const list = Array.isArray(spNodes) ? spNodes : [spNodes];

    const textParts = [];
    for (const sp of list) {
      // Look for the notes placeholder (type 12 = body for notes)
      const phIdx = sp?.['p:nvSpPr']?.['p:nvPr']?.['p:ph']?.['@_type'];
      if (phIdx === 'body' || phIdx === 'notes' || !phIdx) {
        const txBody = sp['p:txBody'];
        if (!txBody) continue;
        const pList = txBody['a:p'];
        if (!pList) continue;
        const pArr = Array.isArray(pList) ? pList : [pList];
        for (const p of pArr) {
          const rList = p['a:r'];
          if (!rList) continue;
          const rArr = Array.isArray(rList) ? rList : [rList];
          for (const r of rArr) {
            const text = r['a:t'];
            if (text != null) textParts.push(String(text));
          }
        }
      }
    }
    return textParts.length ? textParts.join('') : undefined;
  } catch {
    return undefined;
  }
}

// ── Main parse function ──────────────────────────────────────────────────────

/**
 * Parse a .pptx file and extract structured slide data + media files.
 *
 * @param {string} pptxFilePath  Absolute path to the .pptx file.
 * @param {string} deckId        The deck ID to store extracted media in.
 * @param {function} getDeckDir  Function that returns the deck directory for a given deckId.
 * @returns {Promise<object>}    PptxDeckData
 */
async function parsePptx(pptxFilePath, deckId, getDeckDir) {
  const zip = new AdmZip(pptxFilePath);
  const parser = createXmlParser();

  // ── Read presentation.xml ──
  const presXml = zip.readAsText('ppt/presentation.xml');
  const { width, height, slideRIds } = parsePresentationXml(presXml, parser);

  // ── Read presentation rels ──
  const presRelsXml = zip.readAsText('ppt/_rels/presentation.xml.rels');
  const presRels = parseRels(presRelsXml, parser);

  // ── Read theme ──
  let themeData = { colours: {}, defaultFont: undefined };
  try {
    const themeRel = Object.values(presRels).find((r) => r.type.includes('theme'));
    if (themeRel) {
      const themePath = `ppt/${themeRel.target.replace(/^\.\//, '')}`;
      const themeXml = zip.readAsText(themePath);
      themeData = parseTheme(themeXml, parser);
    }
  } catch {
    // Theme parsing is optional
  }

  // ── Resolve slide paths from rels ──
  const slidePaths = [];
  for (const rId of slideRIds) {
    const rel = presRels[rId];
    if (rel) {
      slidePaths.push(rel.target.replace(/^\.\//, ''));
    }
  }

  // ── Prepare media output directory ──
  const slidesOutputDir = path.join(getDeckDir(deckId), 'slides');
  await fs.mkdir(slidesOutputDir, { recursive: true });

  // Track media extraction: pptx-internal-path → relative path in deck
  const mediaMap = new Map();
  let mediaSeq = Date.now(); // simple unique suffix

  async function extractMedia(pptxInternalTarget, slideSubDir) {
    // pptxInternalTarget is relative to the slide, e.g. "../media/image1.png"
    const resolved = path.posix.normalize(path.posix.join(`ppt/${slideSubDir}`, pptxInternalTarget));
    if (mediaMap.has(resolved)) return mediaMap.get(resolved);

    const entry = zip.getEntry(resolved);
    if (!entry) return undefined;

    const ext = path.extname(resolved).toLowerCase() || '.png';
    const outName = `pptx-media-${mediaSeq}${ext}`;
    mediaSeq++;
    const outPath = path.join(slidesOutputDir, outName);
    await fs.writeFile(outPath, entry.getData());
    const relativePath = `slides/${outName}`;
    mediaMap.set(resolved, relativePath);
    return relativePath;
  }

  // ── Parse each slide ──
  const slides = [];

  // Cache for parsed layouts and masters so we don't re-parse them for every slide
  const layoutCache = new Map(); // layoutPath → { background, shapes, placeholders, masterPath }
  const masterCache = new Map(); // masterPath → { background, shapes, placeholders, textStyles }

  /**
   * Extract media from shapes in a layout or master context.
   * Resolves imageTarget fields to extracted file paths.
   */
  function extractMediaFromShapes(shapes, dir) {
    for (const shape of shapes) {
      if (shape.imageTarget) {
        const resolved = path.posix.normalize(path.posix.join(`ppt/${dir}`, shape.imageTarget));
        if (!mediaMap.has(resolved)) {
          const entry = zip.getEntry(resolved);
          if (entry) {
            const ext = path.extname(resolved).toLowerCase() || '.png';
            const outName = `pptx-media-${mediaSeq}${ext}`;
            mediaSeq++;
            const outPath = path.join(slidesOutputDir, outName);
            require('node:fs').writeFileSync(outPath, entry.getData());
            mediaMap.set(resolved, `slides/${outName}`);
          }
        }
        shape.imageRelativePath = mediaMap.get(resolved);
        shape.type = 'image';
        delete shape.imageTarget;
      }
      if (shape.fill?.imageTarget) {
        const resolved = path.posix.normalize(path.posix.join(`ppt/${dir}`, shape.fill.imageTarget));
        if (!mediaMap.has(resolved)) {
          const entry = zip.getEntry(resolved);
          if (entry) {
            const ext = path.extname(resolved).toLowerCase() || '.png';
            const outName = `pptx-media-${mediaSeq}${ext}`;
            mediaSeq++;
            const outPath = path.join(slidesOutputDir, outName);
            require('node:fs').writeFileSync(outPath, entry.getData());
            mediaMap.set(resolved, `slides/${outName}`);
          }
        }
        shape.fill.imageRelativePath = mediaMap.get(resolved);
        shape.fill.type = 'image';
        delete shape.fill.imageTarget;
      }
    }
  }

  function extractMediaFromBackground(background, dir) {
    if (background?.imageTarget) {
      const resolved = path.posix.normalize(path.posix.join(`ppt/${dir}`, background.imageTarget));
      if (!mediaMap.has(resolved)) {
        const entry = zip.getEntry(resolved);
        if (entry) {
          const ext = path.extname(resolved).toLowerCase() || '.png';
          const outName = `pptx-media-${mediaSeq}${ext}`;
          mediaSeq++;
          const outPath = path.join(slidesOutputDir, outName);
          require('node:fs').writeFileSync(outPath, entry.getData());
          mediaMap.set(resolved, `slides/${outName}`);
        }
      }
      background.imageRelativePath = mediaMap.get(resolved);
      background.type = 'image';
      delete background.imageTarget;
    }
  }

  /**
   * Parse a slide layout XML, returning background + shapes.
   * Also returns the master path referenced by the layout.
   */
  function parseLayoutFile(layoutFullPath, layoutDir, layoutBaseName) {
    if (layoutCache.has(layoutFullPath)) return layoutCache.get(layoutFullPath);

    let result = { background: undefined, shapes: [], placeholders: [], masterPath: undefined, showMasterShapes: true };
    try {
      const xml = zip.readAsText(layoutFullPath);
      const doc = parser.parse(xml);
      const sldLayout = doc?.['p:sldLayout'];
      const cSld = sldLayout?.['p:cSld'];
      const showMasterShapes = sldLayout?.['@_showMasterSp'] !== '0' && sldLayout?.['@_showMasterSp'] !== 'false';

      // Parse layout rels first
      let layoutRels = {};
      let masterPath;
      try {
        const layoutRelsPath = `ppt/${layoutDir}/_rels/${layoutBaseName}.rels`;
        const layoutRelsXml = zip.readAsText(layoutRelsPath);
        layoutRels = parseRels(layoutRelsXml, parser);
        const masterRel = Object.values(layoutRels).find((r) => r.type.includes('slideMaster'));
        if (masterRel) {
          masterPath = path.posix.normalize(path.posix.join(`ppt/${layoutDir}`, masterRel.target));
        }
      } catch {
        // No rels file
      }

      // Layout background
      const bg = cSld?.['p:bg'];
      const background = parseSlideBackground(bg, layoutRels, themeData.colours);
      extractMediaFromBackground(background, layoutDir);

      // Layout shapes (using layout rels for image resolution)
      const spTree = cSld?.['p:spTree'];
      const { shapes, placeholders } = parseLayoutMasterShapes(spTree, layoutRels, themeData.colours, themeData.defaultFont);
      extractMediaFromShapes(shapes, layoutDir);
      extractMediaFromShapes(placeholders, layoutDir);

      result = { background, shapes, placeholders, masterPath, showMasterShapes };
    } catch {
      // Layout parsing failure is not fatal
    }
    layoutCache.set(layoutFullPath, result);
    return result;
  }

  /**
   * Parse a slide master XML, returning background + shapes.
   */
  function parseMasterFile(masterFullPath) {
    if (masterCache.has(masterFullPath)) return masterCache.get(masterFullPath);

    let result = { background: undefined, shapes: [], placeholders: [], textStyles: { title: {}, body: {}, other: {} } };
    try {
      const xml = zip.readAsText(masterFullPath);
      const doc = parser.parse(xml);
      const sldMaster = doc?.['p:sldMaster'];
      const cSld = sldMaster?.['p:cSld'];
      const textStyles = parseMasterTextStyles(sldMaster?.['p:txStyles'], themeData.colours, themeData.defaultFont);

      // Master rels
      const masterDir = path.posix.dirname(masterFullPath.replace(/^ppt\//, ''));
      const masterBaseName = path.posix.basename(masterFullPath);

      let masterRels = {};
      try {
        const masterRelsPath = `ppt/${masterDir}/_rels/${masterBaseName}.rels`;
        const masterRelsXml = zip.readAsText(masterRelsPath);
        masterRels = parseRels(masterRelsXml, parser);
      } catch {
        // No rels
      }

      // Master background
      const bg = cSld?.['p:bg'];
      const background = parseSlideBackground(bg, masterRels, themeData.colours);
      extractMediaFromBackground(background, masterDir);

      // Master shapes
      const spTree = cSld?.['p:spTree'];
      const { shapes, placeholders } = parseLayoutMasterShapes(spTree, masterRels, themeData.colours, themeData.defaultFont);
      extractMediaFromShapes(shapes, masterDir);
      extractMediaFromShapes(placeholders, masterDir);

      result = { background, shapes, placeholders, textStyles };
    } catch {
      // Master parsing failure is not fatal
    }
    masterCache.set(masterFullPath, result);
    return result;
  }

  for (let i = 0; i < slidePaths.length; i++) {
    const slidePath = slidePaths[i];
    const fullSlidePath = `ppt/${slidePath}`;
    const slideDir = path.posix.dirname(slidePath); // e.g. "slides"
    const slideBaseName = path.posix.basename(slidePath); // e.g. "slide1.xml"

    let slideXml;
    try {
      slideXml = zip.readAsText(fullSlidePath);
    } catch {
      continue;
    }
    const slideDoc = parser.parse(slideXml);
    const sld = slideDoc?.['p:sld'];

    // ── Slide rels ──
    const slideRelsPath = `ppt/${slideDir}/_rels/${slideBaseName}.rels`;
    let slideRels = {};
    try {
      const slideRelsXml = zip.readAsText(slideRelsPath);
      slideRels = parseRels(slideRelsXml, parser);
    } catch {
      // No rels file
    }

    // ── Find slide layout and master ──
    let layoutData = { background: undefined, shapes: [], placeholders: [], masterPath: undefined, showMasterShapes: true };
    let masterData = { background: undefined, shapes: [], placeholders: [], textStyles: { title: {}, body: {}, other: {} } };

    const layoutRel = Object.values(slideRels).find((r) => r.type.includes('slideLayout'));
    if (layoutRel) {
      const layoutPath = path.posix.normalize(path.posix.join(`ppt/${slideDir}`, layoutRel.target));
      const layoutDir = path.posix.dirname(layoutPath.replace(/^ppt\//, ''));
      const layoutBaseName = path.posix.basename(layoutPath);
      layoutData = parseLayoutFile(layoutPath, layoutDir, layoutBaseName);

      if (layoutData.masterPath) {
        masterData = parseMasterFile(layoutData.masterPath);
      }
    }

    // ── Background (slide → layout → master fallback chain) ──
    const cSld = sld?.['p:cSld'];
    const bg = cSld?.['p:bg'];
    let background = parseSlideBackground(bg, slideRels, themeData.colours);

    // Check if slide explicitly says "use master background" via showMasterSp or no bg
    const useLayoutBg = !background;
    if (useLayoutBg && layoutData.background) {
      background = layoutData.background;
    }
    if (!background && masterData.background) {
      background = masterData.background;
    }

    // ── Shapes from the slide itself ──
    const spTree = cSld?.['p:spTree'];
    const rawSlideShapes = collectDrawableNodes(spTree, slideRels, themeData.colours, themeData.defaultFont);
    const slideShapes = rawSlideShapes
      .map((shape) => {
        if (!shape.placeholder) return shape;
        const masterTemplate = findPlaceholderTemplate(shape, masterData.placeholders);
        const layoutTemplate = findPlaceholderTemplate(shape, layoutData.placeholders);
        return inheritShapeFromTemplate(
          inheritShapeFromTemplate(shape, masterTemplate, masterData.textStyles, themeData.defaultFont),
          layoutTemplate,
          masterData.textStyles,
          themeData.defaultFont,
        );
      })
      .filter((shape) => shape.width > 0 || shape.height > 0);

    // Extract any images referenced directly by picture shapes
    for (const shape of slideShapes) {
      if (shape.imageTarget) {
        const relPath = await extractMedia(shape.imageTarget, slideDir);
        if (relPath) {
          shape.imageRelativePath = relPath;
        }
        delete shape.imageTarget;
      }
    }

    // Extract any images referenced in shape fills
    for (const shape of slideShapes) {
      if (shape.fill?.imageTarget) {
        const relPath = await extractMedia(shape.fill.imageTarget, slideDir);
        if (relPath) {
          shape.fill.imageRelativePath = relPath;
          shape.fill.type = 'image';
        }
        delete shape.fill.imageTarget;
      }
    }

    // Extract background image if present
    if (background?.imageTarget) {
      const relPath = await extractMedia(background.imageTarget, slideDir);
      if (relPath) {
        background.imageRelativePath = relPath;
        background.type = 'image';
      }
      delete background.imageTarget;
    }

    // ── Merge layout/master shapes behind slide shapes ──
    // Master shapes go first (farthest back), then layout, then slide
    const showMasterSp = sld?.['@_showMasterSp'];
    const slideAllowsMasterShapes = showMasterSp !== '0' && showMasterSp !== 'false';
    const showMasterShapes = slideAllowsMasterShapes && layoutData.showMasterShapes !== false;

    const allShapes = [];
    if (showMasterShapes) {
      // Deep-clone master shapes so they don't share references across slides
      for (const s of masterData.shapes) {
        allShapes.push({ ...cloneValue(s), id: `master-${s.id}-s${i}` });
      }
    }
    for (const s of layoutData.shapes) {
      allShapes.push({ ...cloneValue(s), id: `layout-${s.id}-s${i}` });
    }
    allShapes.push(...slideShapes);

    // ── Animations ──
    const timing = sld?.['p:timing'];
    const animMap = parseAnimations(timing);

    // Apply animation groups to shapes
    let maxAnimGroup = 0;
    for (const shape of allShapes) {
      const candidateIds = [];
      if (shape.sourceShapeId) candidateIds.push(String(shape.sourceShapeId));
      if (Array.isArray(shape.ancestorGroupIds)) {
        candidateIds.push(...[...shape.ancestorGroupIds].reverse().map(String));
      }

      const matchedId = candidateIds.find((candidateId) => animMap.has(candidateId));
      if (matchedId) {
        const anim = animMap.get(matchedId);
        shape.animationGroup = anim.group;
        shape.animationEffect = anim.effect;
        if (Array.isArray(anim.paragraphGroups) && anim.paragraphGroups.length && Array.isArray(shape.paragraphs) && shape.paragraphs.length) {
          shape.paragraphs = shape.paragraphs.map((paragraph, index) => ({
            ...paragraph,
            animationGroup: anim.paragraphGroups[Math.min(index, anim.paragraphGroups.length - 1)] ?? anim.group,
          }));
          maxAnimGroup = Math.max(maxAnimGroup, ...anim.paragraphGroups);
        } else {
          if (Array.isArray(shape.paragraphs)) {
            shape.paragraphs = shape.paragraphs.map((paragraph) => ({ ...paragraph, animationGroup: anim.group }));
          }
          if (anim.group > maxAnimGroup) maxAnimGroup = anim.group;
        }
      } else if (Array.isArray(shape.paragraphs)) {
        shape.paragraphs = shape.paragraphs.map((paragraph) => ({ ...paragraph, animationGroup: 0 }));
      }
    }

    // ── Notes ──
    let notes;
    const notesRel = Object.values(slideRels).find((r) => r.type.includes('notesSlide'));
    if (notesRel) {
      try {
        const notesPath = `ppt/${slideDir}/${notesRel.target.replace(/^\.\//, '')}`;
        const notesXml = zip.readAsText(notesPath);
        notes = parseNotes(notesXml, parser);
      } catch {
        // Notes parsing is optional
      }
    }

    slides.push({
      slideIndex: i,
      width,
      height,
      background,
      shapes: allShapes,
      notes,
      animationStepCount: maxAnimGroup,
    });
  }

  return {
    sourceFileName: path.basename(pptxFilePath),
    slides,
    theme: themeData,
  };
}

module.exports = { parsePptx };

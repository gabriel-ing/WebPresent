/**
 * Thin compatibility shim — delegates to @webpresent/pptx-engine.
 *
 * All PPTX parsing logic now lives in packages/pptx-engine/src/parser.ts.
 * This file exists only so that any existing require('./pptxParser.cjs')
 * calls continue to work without changes.
 */
const { parsePptx } = require('@webpresent/pptx-engine');

module.exports = { parsePptx };

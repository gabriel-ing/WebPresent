/**
 * Deck management — CRUD operations, config persistence, import/export.
 *
 * All deck data lives under `<userData>/decks/<deckId>/deck.json` with
 * slide media under `<userData>/decks/<deckId>/slides/`.
 */

const path = require('node:path');
const fs = require('node:fs/promises');
const { existsSync } = require('node:fs');
const AdmZip = require('adm-zip');
const { app, dialog } = require('electron');
const { nowIso, createId } = require('./utils.cjs');

// ── Path helpers ─────────────────────────────────────────────────────────────

function getDecksRoot() {
  return path.join(app.getPath('userData'), 'decks');
}

function getConfigPath() {
  return path.join(app.getPath('userData'), 'config.json');
}

function getDeckDir(deckId) {
  return path.join(getDecksRoot(), deckId);
}

function getDeckJsonPath(deckId) {
  return path.join(getDeckDir(deckId), 'deck.json');
}

async function ensureDeckRoot() {
  await fs.mkdir(getDecksRoot(), { recursive: true });
}

// ── Config ───────────────────────────────────────────────────────────────────

async function readConfig() {
  const configPath = getConfigPath();
  if (!existsSync(configPath)) return {};
  try {
    const raw = await fs.readFile(configPath, 'utf8');
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

async function writeConfig(config) {
  await fs.mkdir(path.dirname(getConfigPath()), { recursive: true });
  await fs.writeFile(getConfigPath(), JSON.stringify(config, null, 2), 'utf8');
}

async function saveLastOpenedDeckId(deckId) {
  const current = await readConfig();
  current.lastOpenedDeckId = deckId;
  await writeConfig(current);
}

// ── CRUD ─────────────────────────────────────────────────────────────────────

async function createDeck() {
  await ensureDeckRoot();
  const deckId = createId('deck');
  const now = nowIso();
  const deck = {
    id: deckId,
    title: 'Untitled deck',
    createdAt: now,
    updatedAt: now,
    items: [],
  };
  await fs.mkdir(path.join(getDeckDir(deckId), 'slides'), { recursive: true });
  await fs.writeFile(getDeckJsonPath(deckId), JSON.stringify(deck, null, 2), 'utf8');
  await saveLastOpenedDeckId(deckId);
  return deck;
}

async function loadDeck(deckId) {
  const raw = await fs.readFile(getDeckJsonPath(deckId), 'utf8');
  return JSON.parse(raw);
}

async function saveDeck(presentation) {
  const deckDir = getDeckDir(presentation.id);
  await fs.mkdir(path.join(deckDir, 'slides'), { recursive: true });
  const next = { ...presentation, updatedAt: nowIso() };
  await fs.writeFile(getDeckJsonPath(next.id), JSON.stringify(next, null, 2), 'utf8');
  await saveLastOpenedDeckId(next.id);
}

async function getOrCreateInitialDeck() {
  await ensureDeckRoot();
  const config = await readConfig();
  const existingDeckId = config.lastOpenedDeckId;
  if (existingDeckId && existsSync(getDeckJsonPath(existingDeckId))) {
    return loadDeck(existingDeckId);
  }
  return createDeck();
}

// ── Import / Export ──────────────────────────────────────────────────────────

async function exportDeck(deckId, ownerWindow) {
  const deck = await loadDeck(deckId);
  const defaultName = `${deck.title || deck.id}.presentdeck`;
  const result = await dialog.showSaveDialog(ownerWindow ?? undefined, {
    title: 'Export Deck',
    defaultPath: defaultName,
    filters: [{ name: 'Present Deck', extensions: ['presentdeck', 'zip'] }],
  });
  if (result.canceled || !result.filePath) return;

  const deckDir = getDeckDir(deckId);
  const slidesDir = path.join(deckDir, 'slides');
  const zip = new AdmZip();
  zip.addLocalFile(path.join(deckDir, 'deck.json'), '', 'deck.json');
  if (existsSync(slidesDir)) {
    const slideFiles = await fs.readdir(slidesDir);
    for (const slideFile of slideFiles) {
      zip.addLocalFile(path.join(slidesDir, slideFile), 'slides', slideFile);
    }
  }
  zip.writeZip(result.filePath);
}

async function importDeck(ownerWindow) {
  const result = await dialog.showOpenDialog(ownerWindow ?? undefined, {
    title: 'Import Deck',
    properties: ['openFile'],
    filters: [{ name: 'Present Deck', extensions: ['presentdeck', 'zip'] }],
  });
  if (result.canceled || !result.filePaths[0]) return null;

  const zip = new AdmZip(result.filePaths[0]);
  const entries = zip.getEntries().map((entry) => entry.entryName);
  if (!entries.includes('deck.json')) {
    throw new Error('Invalid deck file. Missing deck.json');
  }

  const importedId = createId('deck');
  const importedDir = getDeckDir(importedId);
  await fs.mkdir(path.join(importedDir, 'slides'), { recursive: true });
  zip.extractAllTo(importedDir, true);

  const rawDeck = await fs.readFile(path.join(importedDir, 'deck.json'), 'utf8');
  const parsed = JSON.parse(rawDeck);
  const now = nowIso();
  parsed.id = importedId;
  parsed.createdAt = parsed.createdAt || now;
  parsed.updatedAt = now;
  await fs.writeFile(path.join(importedDir, 'deck.json'), JSON.stringify(parsed, null, 2), 'utf8');
  await saveLastOpenedDeckId(importedId);
  return parsed;
}

module.exports = {
  getDeckDir,
  getDeckJsonPath,
  createDeck,
  loadDeck,
  saveDeck,
  getOrCreateInitialDeck,
  exportDeck,
  importDeck,
};

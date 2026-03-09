const { app, BrowserWindow, dialog, ipcMain, screen, shell } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const { existsSync } = require('node:fs');
const { pathToFileURL } = require('node:url');
const crypto = require('node:crypto');
const AdmZip = require('adm-zip');

const isDev = Boolean(process.env.ELECTRON_RENDERER_URL);

let editorWindow = null;
let presentationWindow = null;
let currentPresentationDeck = null;
let currentPresentationIndex = 0;

function nowIso() {
  return new Date().toISOString();
}

function createId(prefix) {
  return `${prefix}-${crypto.randomUUID()}`;
}

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

async function readConfig() {
  const configPath = getConfigPath();
  if (!existsSync(configPath)) {
    return {};
  }
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
  const deckJson = await fs.readFile(getDeckJsonPath(deckId), 'utf8');
  return JSON.parse(deckJson);
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

function extensionFor(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.jpg' || ext === '.jpeg' || ext === '.png' || ext === '.gif' || ext === '.webp') {
    return ext;
  }
  return '.png';
}

function mimeTypeFromPath(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.png') {
    return 'image/png';
  }
  if (ext === '.jpg' || ext === '.jpeg') {
    return 'image/jpeg';
  }
  if (ext === '.gif') {
    return 'image/gif';
  }
  if (ext === '.webp') {
    return 'image/webp';
  }
  return 'application/octet-stream';
}

async function nextSlideNumber(deckId) {
  const slidesDir = path.join(getDeckDir(deckId), 'slides');
  await fs.mkdir(slidesDir, { recursive: true });
  const names = await fs.readdir(slidesDir);
  let max = 0;
  for (const name of names) {
    const match = name.match(/^slide-(\d{4})\./i);
    if (!match) {
      continue;
    }
    const number = Number.parseInt(match[1], 10);
    if (number > max) {
      max = number;
    }
  }
  return max + 1;
}

async function importSlidesToDeck(deckId, filePaths) {
  const slidesDir = path.join(getDeckDir(deckId), 'slides');
  await fs.mkdir(slidesDir, { recursive: true });
  const refs = [];
  let seq = await nextSlideNumber(deckId);
  for (const filePath of filePaths) {
    const ext = extensionFor(filePath);
    const slideBase = `slide-${String(seq).padStart(4, '0')}`;
    const fileName = `${slideBase}${ext}`;
    const destination = path.join(slidesDir, fileName);
    await fs.copyFile(filePath, destination);
    refs.push({
      id: slideBase,
      relativePath: path.posix.join('slides', fileName),
      sourceFileName: path.basename(filePath),
    });
    seq += 1;
  }
  return refs;
}

async function exportDeck(deckId) {
  const deck = await loadDeck(deckId);
  const defaultName = `${deck.title || deck.id}.presentdeck`;
  const result = await dialog.showSaveDialog(editorWindow ?? undefined, {
    title: 'Export Deck',
    defaultPath: defaultName,
    filters: [{ name: 'Present Deck', extensions: ['presentdeck', 'zip'] }],
  });
  if (result.canceled || !result.filePath) {
    return;
  }

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

async function importDeck() {
  const result = await dialog.showOpenDialog(editorWindow ?? undefined, {
    title: 'Import Deck',
    properties: ['openFile'],
    filters: [{ name: 'Present Deck', extensions: ['presentdeck', 'zip'] }],
  });
  if (result.canceled || !result.filePaths[0]) {
    return null;
  }

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

function buildSlideHtml(title, fileUrl, missingReason) {
  const safeTitle = title ? String(title) : 'Slide';
  if (missingReason) {
    return `<!doctype html><html><head><meta charset="utf-8"/><title>${safeTitle}</title><style>html,body{margin:0;background:#000;color:#fff;height:100%;font-family:Segoe UI,Arial,sans-serif}.wrap{height:100%;display:flex;align-items:center;justify-content:center;text-align:center;padding:2rem}</style></head><body><div class="wrap">${missingReason}</div></body></html>`;
  }
  return `<!doctype html><html><head><meta charset="utf-8"/><title>${safeTitle}</title><style>html,body{margin:0;background:#000;height:100%;overflow:hidden}img{width:100%;height:100%;object-fit:contain;display:block}</style></head><body><img src="${fileUrl}" alt="${safeTitle}"/></body></html>`;
}

async function readSlideAsDataUrl(deckId, relativePath) {
  const absolutePath = path.join(getDeckDir(deckId), relativePath);
  if (!existsSync(absolutePath)) {
    return null;
  }
  const mimeType = mimeTypeFromPath(absolutePath);
  const buffer = await fs.readFile(absolutePath);
  return `data:${mimeType};base64,${buffer.toString('base64')}`;
}

async function showPresentationStep(index) {
  if (!presentationWindow || !currentPresentationDeck) {
    return;
  }
  if (index < 0 || index >= currentPresentationDeck.items.length) {
    return;
  }
  currentPresentationIndex = index;
  const step = currentPresentationDeck.items[index];

  if (step.type === 'web' && step.url) {
    await presentationWindow.loadURL(step.url);
    return;
  }

  if (step.type === 'slide' && step.slideRef?.relativePath) {
    const slideDataUrl = await readSlideAsDataUrl(currentPresentationDeck.id, step.slideRef.relativePath);
    const html = slideDataUrl
      ? buildSlideHtml(step.title, slideDataUrl)
      : buildSlideHtml(step.title, '', `Missing slide image: ${step.slideRef.relativePath}`);
    await presentationWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    return;
  }

  const html = buildSlideHtml(step.title, '', 'Invalid step.');
  await presentationWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
}

async function stopPresentation() {
  if (presentationWindow && !presentationWindow.isDestroyed()) {
    presentationWindow.close();
  }
  presentationWindow = null;
  currentPresentationDeck = null;
  currentPresentationIndex = 0;
  if (editorWindow && !editorWindow.isDestroyed()) {
    editorWindow.focus();
  }
}

async function goToNextPresentationStep() {
  if (!currentPresentationDeck) {
    return;
  }
  const maxIndex = currentPresentationDeck.items.length - 1;
  if (currentPresentationIndex < maxIndex) {
    await showPresentationStep(currentPresentationIndex + 1);
  }
}

async function goToPreviousPresentationStep() {
  if (currentPresentationIndex > 0) {
    await showPresentationStep(currentPresentationIndex - 1);
  }
}

async function startPresentation(deckId, startIndex = 0, displayId) {
  const deck = await loadDeck(deckId);
  if (!deck.items.length) {
    throw new Error('Deck has no steps to present.');
  }

  currentPresentationDeck = deck;
  currentPresentationIndex = Math.min(Math.max(0, startIndex), deck.items.length - 1);

  if (presentationWindow && !presentationWindow.isDestroyed()) {
    presentationWindow.close();
  }

  const displays = screen.getAllDisplays();
  const selectedDisplay =
    (typeof displayId === 'number' ? displays.find((d) => d.id === displayId) : undefined) || screen.getPrimaryDisplay();

  presentationWindow = new BrowserWindow({
    x: selectedDisplay.bounds.x,
    y: selectedDisplay.bounds.y,
    width: selectedDisplay.bounds.width,
    height: selectedDisplay.bounds.height,
    backgroundColor: '#000000',
    fullscreen: true,
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  presentationWindow.on('closed', () => {
    presentationWindow = null;
    currentPresentationDeck = null;
    currentPresentationIndex = 0;
  });

  presentationWindow.webContents.on('before-input-event', (event, input) => {
    if (input.type !== 'keyDown') {
      return;
    }
    if (input.key === 'ArrowRight' || input.key === ' ') {
      event.preventDefault();
      void goToNextPresentationStep();
      return;
    }
    if (input.key === 'ArrowLeft') {
      event.preventDefault();
      void goToPreviousPresentationStep();
      return;
    }
    if (input.key === 'Escape') {
      event.preventDefault();
      void stopPresentation();
    }
  });

  await showPresentationStep(currentPresentationIndex);
}

function createEditorWindow() {
  editorWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    backgroundColor: '#f3f4f6',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  if (isDev) {
    editorWindow.loadURL(process.env.ELECTRON_RENDERER_URL);
  } else {
    editorWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }

  editorWindow.on('closed', () => {
    editorWindow = null;
    void stopPresentation();
  });
}

ipcMain.handle('deck:getCurrent', async () => getOrCreateInitialDeck());
ipcMain.handle('deck:create', async () => createDeck());
ipcMain.handle('deck:save', async (_event, presentation) => {
  await saveDeck(presentation);
});
ipcMain.handle('deck:import', async () => importDeck());
ipcMain.handle('deck:export', async (_event, deckId) => {
  await exportDeck(deckId);
});
ipcMain.handle('deck:pickSlideFiles', async () => {
  const result = await dialog.showOpenDialog(editorWindow ?? undefined, {
    title: 'Select slide image files',
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif'] }],
  });
  return result.canceled ? [] : result.filePaths;
});
ipcMain.handle('deck:importSlides', async (_event, params) => {
  return importSlidesToDeck(params.deckId, params.filePaths || []);
});
ipcMain.handle('deck:resolveSlideUrl', async (_event, params) => {
  const fullPath = path.join(getDeckDir(params.deckId), params.relativePath);
  if (!existsSync(fullPath)) {
    return null;
  }
  return pathToFileURL(fullPath).href;
});
ipcMain.handle('deck:resolveSlideDataUrl', async (_event, params) => {
  return readSlideAsDataUrl(params.deckId, params.relativePath);
});
ipcMain.handle('presentation:start', async (_event, params) => {
  await startPresentation(params.deckId, params.startIndex, params.displayId);
});
ipcMain.handle('presentation:stop', async () => {
  await stopPresentation();
});
ipcMain.handle('system:getDisplays', async () => {
  return screen.getAllDisplays().map((display, index) => ({
    id: display.id,
    label: display.label || `Display ${index + 1}`,
    width: display.bounds.width,
    height: display.bounds.height,
  }));
});
ipcMain.handle('system:openExternal', async (_event, url) => {
  if (!url) {
    return;
  }
  await shell.openExternal(url);
});

app.whenReady().then(() => {
  createEditorWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createEditorWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

const { app, BrowserWindow, dialog, globalShortcut, ipcMain, screen, shell } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const { existsSync } = require('node:fs');
const { pathToFileURL } = require('node:url');
const crypto = require('node:crypto');
const AdmZip = require('adm-zip');

const isDev = Boolean(process.env.ELECTRON_RENDERER_URL);
const DIRECTORY_IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.svg', '.tif', '.tiff']);
const naturalFileNameCollator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
const MIN_WEB_ZOOM_PERCENT = 25;
const MAX_WEB_ZOOM_PERCENT = 300;

let editorWindow = null;
let presentationWindow = null;
let currentPresentationDeck = null;
let currentPresentationIndex = 0;
let requestedPresentationIndex = 0;
let presentationNavigationQueue = Promise.resolve();

const PRESENTATION_SHORTCUT_NEXT = 'Shift+Right';
const PRESENTATION_SHORTCUT_PREVIOUS = 'Shift+Left';
const PRESENTATION_SHORTCUT_EXIT = 'Escape';

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
  if (
    ext === '.jpg' ||
    ext === '.jpeg' ||
    ext === '.png' ||
    ext === '.gif' ||
    ext === '.svg' ||
    ext === '.tif' ||
    ext === '.tiff' ||
    ext === '.webp' ||
    ext === '.mp4' ||
    ext === '.webm' ||
    ext === '.mov' ||
    ext === '.m4v' ||
    ext === '.ogv' ||
    ext === '.ogg'
  ) {
    return ext;
  }
  return '.png';
}

function extractSlideNumber(filePath) {
  const name = path.basename(filePath, path.extname(filePath));
  const match = name.match(/^slide\s*([0-9]+)$/i);
  if (!match) {
    return null;
  }
  return Number.parseInt(match[1], 10);
}

function compareSlideFilePaths(a, b) {
  const numberA = extractSlideNumber(a);
  const numberB = extractSlideNumber(b);

  if (numberA !== null && numberB !== null && numberA !== numberB) {
    return numberA - numberB;
  }
  if (numberA !== null && numberB === null) {
    return -1;
  }
  if (numberA === null && numberB !== null) {
    return 1;
  }

  return naturalFileNameCollator.compare(path.basename(a), path.basename(b));
}

function sortSlidePaths(filePaths) {
  return [...filePaths].sort(compareSlideFilePaths);
}

function mediaKindFromPath(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.mp4' || ext === '.webm' || ext === '.mov' || ext === '.m4v' || ext === '.ogv' || ext === '.ogg') {
    return 'video';
  }
  return 'image';
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
  if (ext === '.svg') {
    return 'image/svg+xml';
  }
  if (ext === '.tif' || ext === '.tiff') {
    return 'image/tiff';
  }
  if (ext === '.webp') {
    return 'image/webp';
  }
  if (ext === '.mp4' || ext === '.m4v') {
    return 'video/mp4';
  }
  if (ext === '.webm') {
    return 'video/webm';
  }
  if (ext === '.ogv' || ext === '.ogg') {
    return 'video/ogg';
  }
  if (ext === '.mov') {
    return 'video/quicktime';
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
  const orderedPaths = sortSlidePaths(filePaths);
  for (const filePath of orderedPaths) {
    const ext = extensionFor(filePath);
    const slideBase = `slide-${String(seq).padStart(4, '0')}`;
    const fileName = `${slideBase}${ext}`;
    const destination = path.join(slidesDir, fileName);
    await fs.copyFile(filePath, destination);
    refs.push({
      id: slideBase,
      relativePath: path.posix.join('slides', fileName),
      sourceFileName: path.basename(filePath),
      mediaKind: mediaKindFromPath(filePath),
    });
    seq += 1;
  }
  return refs;
}

async function pickDirectoryImageFiles() {
  const result = await dialog.showOpenDialog(editorWindow ?? undefined, {
    title: 'Select slide image folder',
    properties: ['openDirectory'],
  });
  if (result.canceled || !result.filePaths[0]) {
    return [];
  }

  const directoryPath = result.filePaths[0];
  const entries = await fs.readdir(directoryPath, { withFileTypes: true });
  const files = entries
    .filter((entry) => entry.isFile())
    .map((entry) => path.join(directoryPath, entry.name))
    .filter((filePath) => DIRECTORY_IMAGE_EXTENSIONS.has(path.extname(filePath).toLowerCase()));

  return sortSlidePaths(files);
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

function buildSlideHtml(title, fileUrl, missingReason, mediaKind = 'image') {
  const safeTitle = title ? String(title) : 'Slide';
  if (missingReason) {
    return `<!doctype html><html><head><meta charset="utf-8"/><title>${safeTitle}</title><style>html,body{margin:0;background:#000;color:#fff;height:100%;font-family:Segoe UI,Arial,sans-serif}.wrap{height:100%;display:flex;align-items:center;justify-content:center;text-align:center;padding:2rem}</style></head><body><div class="wrap">${missingReason}</div></body></html>`;
  }
  const mediaTag =
    mediaKind === 'video'
      ? `<video src="${fileUrl}" autoplay controls style="width:100%;height:100%;object-fit:contain;background:#000;display:block"></video>`
      : `<img src="${fileUrl}" alt="${safeTitle}"/>`;
  return `<!doctype html><html><head><meta charset="utf-8"/><title>${safeTitle}</title><style>html,body{margin:0;background:#000;height:100%;overflow:hidden}img{width:100%;height:100%;object-fit:contain;display:block}</style></head><body>${mediaTag}</body></html>`;
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

function getSlideFileUrl(deckId, relativePath) {
  const absolutePath = path.join(getDeckDir(deckId), relativePath);
  if (!existsSync(absolutePath)) {
    return null;
  }
  return pathToFileURL(absolutePath).href;
}

function isIgnorableLoadError(error) {
  const message = String(error?.message || '');
  return message.includes('ERR_ABORTED') || message.includes('ERR_INVALID_URL');
}

function normalizeWebZoomPercent(value) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return 100;
  }
  return Math.min(MAX_WEB_ZOOM_PERCENT, Math.max(MIN_WEB_ZOOM_PERCENT, Math.round(parsed)));
}

async function loadPresentationUrl(url) {
  if (!presentationWindow || presentationWindow.isDestroyed()) {
    return;
  }
  try {
    await presentationWindow.loadURL(url);
  } catch (error) {
    if (isIgnorableLoadError(error)) {
      return;
    }
    throw error;
  }
}

async function applySlideViewportLayout() {
  if (!presentationWindow || presentationWindow.isDestroyed()) {
    return;
  }
  // Style only the currently loaded slide document so media fills the fullscreen window.
  await presentationWindow.webContents.executeJavaScript(
    `(() => {
      const doc = document;
      const html = doc.documentElement;
      const body = doc.body;
      if (!html) {
        return;
      }
      html.style.margin = '0';
      html.style.width = '100%';
      html.style.height = '100%';
      html.style.overflow = 'hidden';
      html.style.background = '#000';
      if (body) {
        body.style.margin = '0';
        body.style.width = '100%';
        body.style.height = '100%';
        body.style.overflow = 'hidden';
        body.style.background = '#000';
      }

      const rootTag = (html.tagName || '').toLowerCase();
      if (rootTag === 'svg') {
        const svg = html;
        const widthAttr = Number.parseFloat((svg.getAttribute('width') || '').replace(/px$/i, ''));
        const heightAttr = Number.parseFloat((svg.getAttribute('height') || '').replace(/px$/i, ''));
        if (!svg.hasAttribute('viewBox') && Number.isFinite(widthAttr) && Number.isFinite(heightAttr) && widthAttr > 0 && heightAttr > 0) {
          svg.setAttribute('viewBox', '0 0 ' + widthAttr + ' ' + heightAttr);
        }
        svg.removeAttribute('width');
        svg.removeAttribute('height');
        svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
        svg.style.display = 'block';
        svg.style.position = 'fixed';
        svg.style.inset = '0';
        svg.style.width = '100vw';
        svg.style.height = '100vh';
        svg.style.maxWidth = '100vw';
        svg.style.maxHeight = '100vh';
        svg.style.margin = '0';
        svg.style.padding = '0';
        svg.style.overflow = 'hidden';
        svg.style.background = '#000';
        return;
      }

      const media = doc.querySelector('img,video,svg');
      if (!media) {
        return;
      }
      media.style.display = 'block';
      media.style.width = '100vw';
      media.style.height = '100vh';
      media.style.objectFit = 'contain';
      media.style.objectPosition = 'center center';
      media.style.background = '#000';
      media.style.margin = '0';
    })();`,
    true,
  );
}

function queuePresentationStep(index) {
  presentationNavigationQueue = presentationNavigationQueue
    .then(async () => {
      await showPresentationStep(index);
    })
    .catch((error) => {
      if (!isIgnorableLoadError(error)) {
        console.error('Could not load queued presentation step.', error);
      }
    });
  return presentationNavigationQueue;
}

function withPresentationErrorLogging(task, fallbackMessage) {
  void task().catch((error) => {
    if (!isIgnorableLoadError(error)) {
      console.error(fallbackMessage, error);
    }
  });
}

function unregisterPresentationShortcuts() {
  globalShortcut.unregister(PRESENTATION_SHORTCUT_NEXT);
  globalShortcut.unregister(PRESENTATION_SHORTCUT_PREVIOUS);
  globalShortcut.unregister(PRESENTATION_SHORTCUT_EXIT);
}

function registerPresentationShortcuts() {
  unregisterPresentationShortcuts();

  globalShortcut.register(PRESENTATION_SHORTCUT_NEXT, () => {
    if (!presentationWindow || presentationWindow.isDestroyed() || !presentationWindow.isFocused()) {
      return;
    }
    withPresentationErrorLogging(() => goToNextPresentationStep(), 'Could not advance presentation step.');
  });

  globalShortcut.register(PRESENTATION_SHORTCUT_PREVIOUS, () => {
    if (!presentationWindow || presentationWindow.isDestroyed() || !presentationWindow.isFocused()) {
      return;
    }
    withPresentationErrorLogging(() => goToPreviousPresentationStep(), 'Could not go back presentation step.');
  });

  globalShortcut.register(PRESENTATION_SHORTCUT_EXIT, () => {
    if (!presentationWindow || presentationWindow.isDestroyed() || !presentationWindow.isFocused()) {
      return;
    }
    withPresentationErrorLogging(() => stopPresentation(), 'Could not stop presentation window.');
  });
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
    const zoomFactor = normalizeWebZoomPercent(step.webZoom) / 100;
    presentationWindow.webContents.setZoomFactor(zoomFactor);
    await loadPresentationUrl(step.url);
    presentationWindow.webContents.setZoomFactor(zoomFactor);
    return;
  }

  presentationWindow.webContents.setZoomFactor(1);

  if (step.type === 'slide' && step.slideRef?.relativePath) {
    const slideFileUrl = getSlideFileUrl(currentPresentationDeck.id, step.slideRef.relativePath);
    if (slideFileUrl) {
      await loadPresentationUrl(slideFileUrl);
      await applySlideViewportLayout();
      return;
    }
    const html = buildSlideHtml(step.title, '', `Missing slide media: ${step.slideRef.relativePath}`);
    await loadPresentationUrl(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    return;
  }

  const html = buildSlideHtml(step.title, '', 'Invalid step.');
  await loadPresentationUrl(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
}

async function stopPresentation() {
  unregisterPresentationShortcuts();
  if (presentationWindow && !presentationWindow.isDestroyed()) {
    // Some web pages (for example notebook UIs) can block normal close via unload hooks.
    presentationWindow.destroy();
  }
  presentationWindow = null;
  currentPresentationDeck = null;
  currentPresentationIndex = 0;
  requestedPresentationIndex = 0;
  presentationNavigationQueue = Promise.resolve();
  if (editorWindow && !editorWindow.isDestroyed()) {
    editorWindow.focus();
  }
}

async function goToNextPresentationStep() {
  if (!currentPresentationDeck) {
    return;
  }
  const maxIndex = currentPresentationDeck.items.length - 1;
  if (requestedPresentationIndex < maxIndex) {
    requestedPresentationIndex += 1;
    await queuePresentationStep(requestedPresentationIndex);
  }
}

async function goToPreviousPresentationStep() {
  if (requestedPresentationIndex > 0) {
    requestedPresentationIndex -= 1;
    await queuePresentationStep(requestedPresentationIndex);
  }
}

async function startPresentation(deckId, startIndex = 0, displayId) {
  const deck = await loadDeck(deckId);
  if (!deck.items.length) {
    throw new Error('Deck has no steps to present.');
  }

  currentPresentationDeck = deck;
  currentPresentationIndex = Math.min(Math.max(0, startIndex), deck.items.length - 1);
  requestedPresentationIndex = currentPresentationIndex;
  presentationNavigationQueue = Promise.resolve();

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
    unregisterPresentationShortcuts();
    presentationWindow = null;
    currentPresentationDeck = null;
    currentPresentationIndex = 0;
    requestedPresentationIndex = 0;
    presentationNavigationQueue = Promise.resolve();
  });

  presentationWindow.on('leave-full-screen', () => {
    // Keep presentation lifecycle simple: leaving fullscreen ends the presentation window.
    void stopPresentation().catch((error) => {
      console.error('Could not stop presentation after leaving fullscreen.', error);
    });
  });

  presentationWindow.webContents.on('before-input-event', (event, input) => {
    if (input.type !== 'keyDown') {
      return;
    }
    if (input.key === 'ArrowRight' && input.shift) {
      event.preventDefault();
      void goToNextPresentationStep().catch((error) => {
        if (!isIgnorableLoadError(error)) {
          console.error('Could not advance presentation step.', error);
        }
      });
      return;
    }
    if (input.key === 'ArrowLeft' && input.shift) {
      event.preventDefault();
      void goToPreviousPresentationStep().catch((error) => {
        if (!isIgnorableLoadError(error)) {
          console.error('Could not go back presentation step.', error);
        }
      });
      return;
    }
    if (input.key === 'Escape') {
      event.preventDefault();
      void stopPresentation().catch((error) => {
        console.error('Could not stop presentation window.', error);
      });
    }
  });

  presentationWindow.webContents.on('will-prevent-unload', (event) => {
    // Interactive apps like Jupyter can block unload after edits/runs.
    // Presentation navigation should always be allowed to continue.
    event.preventDefault();
  });

  registerPresentationShortcuts();

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
    title: 'Select slide image or video files',
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Media', extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'svg', 'tif', 'tiff', 'mp4', 'webm', 'mov', 'm4v', 'ogv', 'ogg'] },
    ],
  });
  return result.canceled ? [] : result.filePaths;
});
ipcMain.handle('deck:pickSlideDirectory', async () => {
  return pickDirectoryImageFiles();
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

app.on('will-quit', () => {
  unregisterPresentationShortcuts();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

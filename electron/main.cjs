const { app, BrowserWindow, ipcMain, screen, globalShortcut, shell } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const { existsSync } = require('node:fs');
const { pathToFileURL } = require('node:url');

const { getDeckDir, createDeck, loadDeck, saveDeck, getOrCreateInitialDeck, exportDeck, importDeck } = require('./deckManager.cjs');
const { importSlidesToDeck, pickSlideFiles, pickDirectoryImageFiles, pickPptxFile } = require('./slideImporter.cjs');
const { mimeTypeFromPath } = require('./utils.cjs');
const { parsePptx } = require('./pptxParser.cjs');
const { buildPptxPresentationDocument, buildPptxRuntimeUpdateScript, createPptxPresentationDocumentBuilder } = require('./pptxPresentRenderer.cjs');
const { createLatestOnlyQueue } = require('./latestOnlyQueue.cjs');

const isDev = Boolean(process.env.ELECTRON_RENDERER_URL);

const MIN_WEB_ZOOM_PERCENT = 25;
const MAX_WEB_ZOOM_PERCENT = 300;

const PRESENTATION_SHORTCUT_NEXT = 'Shift+Right';
const PRESENTATION_SHORTCUT_PREVIOUS = 'Shift+Left';
const PRESENTATION_SHORTCUT_EXIT = 'Escape';

// ── Presentation state ────────────────────────────────────────────────────────

let editorWindow = null;
let presentationWindow = null;
let currentPresentationDeck = null;
let currentPresentationIndex = 0;
let requestedPresentationIndex = 0;
let enqueuePresentationStep = null;
let buildCachedPptxPresentationDocument = null;
let lastPresentedStep = null;

// ── Helpers ───────────────────────────────────────────────────────────────────

function normalizeWebZoomPercent(value) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return 100;
  return Math.min(MAX_WEB_ZOOM_PERCENT, Math.max(MIN_WEB_ZOOM_PERCENT, Math.round(parsed)));
}

function isIgnorableLoadError(error) {
  const message = String(error?.message || '');
  return message.includes('ERR_ABORTED') || message.includes('ERR_INVALID_URL');
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
  if (!existsSync(absolutePath)) return null;
  const mimeType = mimeTypeFromPath(absolutePath);
  const buffer = await fs.readFile(absolutePath);
  return `data:${mimeType};base64,${buffer.toString('base64')}`;
}

function getSlideFileUrl(deckId, relativePath) {
  const absolutePath = path.join(getDeckDir(deckId), relativePath);
  if (!existsSync(absolutePath)) return null;
  return pathToFileURL(absolutePath).href;
}

async function loadPresentationUrl(url) {
  if (!presentationWindow || presentationWindow.isDestroyed()) return;
  try {
    await presentationWindow.loadURL(url);
  } catch (error) {
    if (!isIgnorableLoadError(error)) throw error;
  }
}

async function applySlideViewportLayout() {
  if (!presentationWindow || presentationWindow.isDestroyed()) return;
  await presentationWindow.webContents.executeJavaScript(
    `(() => {
      const doc = document;
      const html = doc.documentElement;
      const body = doc.body;
      if (!html) return;
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
      if (!media) return;
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

function canUpdatePptxInPlace(previousStep, nextStep) {
  return previousStep?.type === 'pptx-slide' && nextStep?.type === 'pptx-slide';
}

// ── Presentation navigation ───────────────────────────────────────────────────

async function showPresentationStep(index) {
  if (!presentationWindow || !currentPresentationDeck) return;
  if (index < 0 || index >= currentPresentationDeck.items.length) return;

  currentPresentationIndex = index;
  const step = currentPresentationDeck.items[index];
  const previousStep = lastPresentedStep;

  if (step.type === 'web' && step.url) {
    const zoomFactor = normalizeWebZoomPercent(step.webZoom) / 100;
    presentationWindow.webContents.setZoomFactor(zoomFactor);
    await loadPresentationUrl(step.url);
    presentationWindow.webContents.setZoomFactor(zoomFactor);
    lastPresentedStep = step;
    return;
  }

  presentationWindow.webContents.setZoomFactor(1);

  if (step.type === 'slide' && step.slideRef?.relativePath) {
    const slideFileUrl = getSlideFileUrl(currentPresentationDeck.id, step.slideRef.relativePath);
    if (slideFileUrl) {
      await loadPresentationUrl(slideFileUrl);
      await applySlideViewportLayout();
      lastPresentedStep = step;
      return;
    }
    const html = buildSlideHtml(step.title, '', `Missing slide media: ${step.slideRef.relativePath}`);
    await loadPresentationUrl(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    lastPresentedStep = step;
    return;
  }

  if (step.type === 'pptx-slide' && step.pptxSlideData) {
    if (canUpdatePptxInPlace(previousStep, step)) {
      try {
        const updateScript = buildCachedPptxPresentationDocument?.buildUpdateScript
          ? await buildCachedPptxPresentationDocument.buildUpdateScript(step.pptxSlideData, step.pptxAnimationStep || 0)
          : await buildPptxRuntimeUpdateScript(
            step.pptxSlideData,
            step.pptxAnimationStep || 0,
            getDeckDir(currentPresentationDeck.id),
          );
        await presentationWindow.webContents.executeJavaScript(updateScript, true);
        lastPresentedStep = step;
        return;
      } catch (error) {
        if (!isIgnorableLoadError(error)) {
          console.warn('Falling back to full PPTX reload after in-place update failure.', error);
        }
      }
    }

    const html = buildCachedPptxPresentationDocument
      ? await buildCachedPptxPresentationDocument(step.pptxSlideData, step.pptxAnimationStep || 0)
      : await buildPptxPresentationDocument(
        step.pptxSlideData,
        step.pptxAnimationStep || 0,
        getDeckDir(currentPresentationDeck.id),
      );
    await loadPresentationUrl(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    lastPresentedStep = step;
    return;
  }

  const html = buildSlideHtml(step.title, '', 'Invalid step.');
  await loadPresentationUrl(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
  lastPresentedStep = step;
}

function requestPresentationStep(index) {
  if (!currentPresentationDeck || !enqueuePresentationStep) return Promise.resolve();
  const boundedIndex = Math.max(0, Math.min(index, currentPresentationDeck.items.length - 1));
  if (boundedIndex === requestedPresentationIndex && boundedIndex === currentPresentationIndex) {
    return Promise.resolve();
  }
  requestedPresentationIndex = boundedIndex;
  return enqueuePresentationStep(boundedIndex);
}

function withPresentationErrorLogging(task, fallbackMessage) {
  void task().catch((error) => {
    if (!isIgnorableLoadError(error)) console.error(fallbackMessage, error);
  });
}

async function goToNextPresentationStep() {
  if (!currentPresentationDeck) return;
  const maxIndex = currentPresentationDeck.items.length - 1;
  if (requestedPresentationIndex < maxIndex) {
    await requestPresentationStep(requestedPresentationIndex + 1);
  }
}

async function goToPreviousPresentationStep() {
  if (requestedPresentationIndex > 0) {
    await requestPresentationStep(requestedPresentationIndex - 1);
  }
}

function resetPresentationState() {
  currentPresentationDeck = null;
  currentPresentationIndex = 0;
  requestedPresentationIndex = 0;
  enqueuePresentationStep = null;
  buildCachedPptxPresentationDocument = null;
  lastPresentedStep = null;
}

async function stopPresentation() {
  unregisterPresentationShortcuts();
  if (presentationWindow && !presentationWindow.isDestroyed()) {
    presentationWindow.destroy();
  }
  presentationWindow = null;
  resetPresentationState();
  if (editorWindow && !editorWindow.isDestroyed()) {
    editorWindow.focus();
  }
}

async function startPresentation(deckId, startIndex = 0, displayId) {
  const deck = await loadDeck(deckId);
  if (!deck.items.length) throw new Error('Deck has no steps to present.');

  currentPresentationDeck = deck;
  currentPresentationIndex = Math.min(Math.max(0, startIndex), deck.items.length - 1);
  requestedPresentationIndex = currentPresentationIndex;
  enqueuePresentationStep = createLatestOnlyQueue(async (index) => {
    await showPresentationStep(index);
  });
  buildCachedPptxPresentationDocument = createPptxPresentationDocumentBuilder(getDeckDir(deck.id));

  if (presentationWindow && !presentationWindow.isDestroyed()) {
    presentationWindow.close();
  }

  const displays = screen.getAllDisplays();
  const selectedDisplay =
    (typeof displayId === 'number' ? displays.find((d) => d.id === displayId) : undefined) ||
    screen.getPrimaryDisplay();

  presentationWindow = new BrowserWindow({
    x: selectedDisplay.bounds.x,
    y: selectedDisplay.bounds.y,
    width: selectedDisplay.bounds.width,
    height: selectedDisplay.bounds.height,
    backgroundColor: '#000000',
    fullscreen: true,
    autoHideMenuBar: true,
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true },
  });

  presentationWindow.on('closed', () => {
    unregisterPresentationShortcuts();
    presentationWindow = null;
    resetPresentationState();
  });

  presentationWindow.on('leave-full-screen', () => {
    void stopPresentation().catch((error) => {
      console.error('Could not stop presentation after leaving fullscreen.', error);
    });
  });

  presentationWindow.webContents.on('before-input-event', (event, input) => {
    if (input.type !== 'keyDown') return;
    if (input.key === 'ArrowRight' && input.shift) {
      event.preventDefault();
      void goToNextPresentationStep().catch((error) => {
        if (!isIgnorableLoadError(error)) console.error('Could not advance presentation step.', error);
      });
      return;
    }
    if (input.key === 'ArrowLeft' && input.shift) {
      event.preventDefault();
      void goToPreviousPresentationStep().catch((error) => {
        if (!isIgnorableLoadError(error)) console.error('Could not go back presentation step.', error);
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
    event.preventDefault();
  });

  registerPresentationShortcuts();
  await showPresentationStep(currentPresentationIndex);
}

// ── Global shortcuts ──────────────────────────────────────────────────────────

function unregisterPresentationShortcuts() {
  globalShortcut.unregister(PRESENTATION_SHORTCUT_NEXT);
  globalShortcut.unregister(PRESENTATION_SHORTCUT_PREVIOUS);
  globalShortcut.unregister(PRESENTATION_SHORTCUT_EXIT);
}

function registerPresentationShortcuts() {
  unregisterPresentationShortcuts();
  globalShortcut.register(PRESENTATION_SHORTCUT_NEXT, () => {
    if (!presentationWindow || presentationWindow.isDestroyed() || !presentationWindow.isFocused()) return;
    withPresentationErrorLogging(() => goToNextPresentationStep(), 'Could not advance presentation step.');
  });
  globalShortcut.register(PRESENTATION_SHORTCUT_PREVIOUS, () => {
    if (!presentationWindow || presentationWindow.isDestroyed() || !presentationWindow.isFocused()) return;
    withPresentationErrorLogging(() => goToPreviousPresentationStep(), 'Could not go back presentation step.');
  });
  globalShortcut.register(PRESENTATION_SHORTCUT_EXIT, () => {
    if (!presentationWindow || presentationWindow.isDestroyed() || !presentationWindow.isFocused()) return;
    withPresentationErrorLogging(() => stopPresentation(), 'Could not stop presentation window.');
  });
}

// ── Editor window ─────────────────────────────────────────────────────────────

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

// ── IPC handlers ──────────────────────────────────────────────────────────────

ipcMain.handle('deck:getCurrent', () => getOrCreateInitialDeck());
ipcMain.handle('deck:create', () => createDeck());
ipcMain.handle('deck:save', (_event, presentation) => saveDeck(presentation));
ipcMain.handle('deck:import', () => importDeck(editorWindow));
ipcMain.handle('deck:export', (_event, deckId) => exportDeck(deckId, editorWindow));

ipcMain.handle('deck:pickSlideFiles', () => pickSlideFiles(editorWindow));
ipcMain.handle('deck:pickSlideDirectory', () => pickDirectoryImageFiles(editorWindow));
ipcMain.handle('deck:importSlides', (_event, params) => importSlidesToDeck(params.deckId, params.filePaths || []));

ipcMain.handle('deck:pickPptxFile', () => pickPptxFile(editorWindow));
ipcMain.handle('deck:importPptx', (_event, params) => parsePptx(params.filePath, params.deckId, getDeckDir));

ipcMain.handle('deck:resolveSlideUrl', (_event, params) => {
  const fullPath = path.join(getDeckDir(params.deckId), params.relativePath);
  if (!existsSync(fullPath)) return null;
  return pathToFileURL(fullPath).href;
});
ipcMain.handle('deck:resolveSlideDataUrl', (_event, params) =>
  readSlideAsDataUrl(params.deckId, params.relativePath),
);

ipcMain.handle('presentation:start', (_event, params) =>
  startPresentation(params.deckId, params.startIndex, params.displayId),
);
ipcMain.handle('presentation:stop', () => stopPresentation());

ipcMain.handle('system:getDisplays', () =>
  screen.getAllDisplays().map((display, index) => ({
    id: display.id,
    label: display.label || `Display ${index + 1}`,
    width: display.bounds.width,
    height: display.bounds.height,
  })),
);
ipcMain.handle('system:openExternal', (_event, url) => {
  if (!url) return;
  return shell.openExternal(url);
});

// ── App lifecycle ─────────────────────────────────────────────────────────────

app.whenReady().then(() => {
  createEditorWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createEditorWindow();
  });
});

app.on('will-quit', () => {
  unregisterPresentationShortcuts();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

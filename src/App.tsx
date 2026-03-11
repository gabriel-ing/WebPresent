import { DragEvent, useEffect, useMemo, useRef, useState } from 'react';
import './App.css';
import type { DisplayInfo, Presentation, PresentationStep, SlideImportMode, SlideRef } from './types';
import { createId, guessTitleFromUrl, moveItem, nowIso } from './utils';

const AUTOSAVE_DEBOUNCE_MS = 400;
const VIDEO_EXTENSIONS = new Set(['.mp4', '.webm', '.mov', '.m4v', '.ogv', '.ogg']);
const MIN_WEB_ZOOM = 25;
const MAX_WEB_ZOOM = 300;

function normalizeWebZoom(value: number): number {
  if (!Number.isFinite(value)) {
    return 100;
  }
  return Math.min(MAX_WEB_ZOOM, Math.max(MIN_WEB_ZOOM, Math.round(value)));
}

function getSlideMediaKind(slideRef?: SlideRef): 'image' | 'video' {
  if (slideRef?.mediaKind) {
    return slideRef.mediaKind;
  }
  const source = slideRef?.relativePath || slideRef?.sourceFileName || '';
  const extension = source.includes('.') ? source.slice(source.lastIndexOf('.')).toLowerCase() : '';
  return VIDEO_EXTENSIONS.has(extension) ? 'video' : 'image';
}

function getStepTitle(step: PresentationStep): string {
  if (step.title?.trim()) {
    return step.title;
  }
  if (step.type === 'web') {
    return step.url ? guessTitleFromUrl(step.url) : 'Untitled web step';
  }
  return step.slideRef?.sourceFileName || step.slideRef?.id || 'Untitled slide step';
}

function insertAfterSelected(items: PresentationStep[], selectedId: string | null, incoming: PresentationStep[]): PresentationStep[] {
  if (!selectedId) {
    return [...items, ...incoming];
  }
  const selectedIndex = items.findIndex((item) => item.id === selectedId);
  if (selectedIndex === -1) {
    return [...items, ...incoming];
  }
  return [...items.slice(0, selectedIndex + 1), ...incoming, ...items.slice(selectedIndex + 1)];
}

export default function App() {
  const [presentation, setPresentation] = useState<Presentation | null>(null);
  const [selectedStepId, setSelectedStepId] = useState<string | null>(null);
  const [draggingStepId, setDraggingStepId] = useState<string | null>(null);
  const [slideUrls, setSlideUrls] = useState<Record<string, string>>({});
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [displayOptions, setDisplayOptions] = useState<DisplayInfo[]>([]);
  const [selectedDisplayId, setSelectedDisplayId] = useState<number | ''>('');
  const [isAddingWebStep, setIsAddingWebStep] = useState(false);
  const [newWebUrl, setNewWebUrl] = useState('');
  const [newWebTitle, setNewWebTitle] = useState('');
  const autosaveTimerRef = useRef<number | null>(null);
  const hasElectronApi = Boolean(window.presentApi);

  const selectedStep = useMemo(() => {
    if (!presentation || !selectedStepId) {
      return null;
    }
    return presentation.items.find((item) => item.id === selectedStepId) || null;
  }, [presentation, selectedStepId]);

  useEffect(() => {
    let mounted = true;
    async function init() {
      if (!window.presentApi) {
        setToastMessage('Electron bridge not found. Run with npm run dev for desktop mode.');
        return;
      }
      const [loaded, displays] = await Promise.all([window.presentApi.deckGetCurrent(), window.presentApi.getDisplays()]);
      if (!mounted) {
        return;
      }
      setPresentation(loaded);
      setSelectedStepId(loaded.items[0]?.id ?? null);
      setDisplayOptions(displays);
      if (displays[0]) {
        setSelectedDisplayId(displays[0].id);
      }
    }
    void init();
    return () => {
      mounted = false;
    };
  }, []);

  useEffect(() => {
    if (!presentation || !window.presentApi) {
      return;
    }
    if (autosaveTimerRef.current) {
      window.clearTimeout(autosaveTimerRef.current);
    }
    autosaveTimerRef.current = window.setTimeout(async () => {
      try {
        await window.presentApi!.deckSave(presentation);
      } catch {
        setToastMessage('Autosave failed. Try exporting your deck.');
      }
    }, AUTOSAVE_DEBOUNCE_MS);
  }, [presentation]);

  useEffect(() => {
    if (!presentation || !selectedStep?.slideRef || !window.presentApi) {
      return;
    }
    const deckId = presentation.id;
    const relativePath = selectedStep.slideRef.relativePath;
    const key = `${deckId}|${relativePath}`;
    if (slideUrls[key]) {
      return;
    }

    let mounted = true;
    async function resolveSlideUrl() {
      const resolved = await window.presentApi!.resolveSlideDataUrl({
        deckId,
        relativePath,
      });
      if (!mounted || !resolved) {
        return;
      }
      setSlideUrls((prev) => ({ ...prev, [key]: resolved }));
    }

    void resolveSlideUrl();
    return () => {
      mounted = false;
    };
  }, [presentation, selectedStep, slideUrls]);

  useEffect(() => {
    if (!toastMessage) {
      return;
    }
    const timer = window.setTimeout(() => setToastMessage(null), 3500);
    return () => window.clearTimeout(timer);
  }, [toastMessage]);

  const updatePresentation = (updater: (oldValue: Presentation) => Presentation) => {
    setPresentation((current) => {
      if (!current) {
        return current;
      }
      const next = updater(current);
      return { ...next, updatedAt: nowIso() };
    });
  };

  const createNewDeck = async () => {
    if (!window.presentApi) {
      return;
    }
    const next = await window.presentApi.deckCreate();
    setPresentation(next);
    setSelectedStepId(next.items[0]?.id ?? null);
    setSlideUrls({});
  };

  const submitNewWebStep = () => {
    const url = newWebUrl.trim();
    if (!url) {
      setToastMessage('Please enter a URL.');
      return;
    }

    const title = newWebTitle.trim() || undefined;
    const step: PresentationStep = {
      id: createId('step'),
      type: 'web',
      title,
      url,
      webZoom: 100,
    };

    updatePresentation((current) => ({
      ...current,
      items: insertAfterSelected(current.items, selectedStepId, [step]),
    }));
    setSelectedStepId(step.id);
    setIsAddingWebStep(false);
    setNewWebUrl('');
    setNewWebTitle('');
  };

  const addSlideSteps = async () => {
    if (!presentation || !window.presentApi) {
      return;
    }

    const filePaths = await window.presentApi.pickSlideFiles();
    if (!filePaths.length) {
      return;
    }

    const mode: SlideImportMode =
      filePaths.length > 1 && window.confirm('Import as one animation group?\nOK = grouped, Cancel = separate steps')
        ? 'grouped'
        : 'separate';

    let slideRefs: SlideRef[] = [];
    try {
      slideRefs = await window.presentApi.importSlides({
        deckId: presentation.id,
        filePaths,
        mode,
      });
    } catch {
      setToastMessage('Could not import one or more slide media files.');
      return;
    }

    const groupId = mode === 'grouped' && slideRefs.length > 1 ? createId('group') : undefined;
    const steps: PresentationStep[] = slideRefs.map((slideRef) => ({
      id: createId('step'),
      type: 'slide',
      title: slideRef.sourceFileName || slideRef.id,
      groupId,
      slideRef,
    }));

    updatePresentation((current) => ({
      ...current,
      items: insertAfterSelected(current.items, selectedStepId, steps),
    }));
    setSelectedStepId(steps[0]?.id ?? null);
  };

  const updateSelectedStep = (updater: (step: PresentationStep) => PresentationStep) => {
    if (!selectedStepId) {
      return;
    }
    updatePresentation((current) => ({
      ...current,
      items: current.items.map((item) => (item.id === selectedStepId ? updater(item) : item)),
    }));
  };

  const deleteStep = (stepId: string) => {
    updatePresentation((current) => ({
      ...current,
      items: current.items.filter((item) => item.id !== stepId),
    }));
    if (selectedStepId === stepId) {
      setSelectedStepId(null);
    }
  };

  const onStepDrop = (event: DragEvent<HTMLDivElement>, targetId: string) => {
    event.preventDefault();
    if (!draggingStepId || !presentation || draggingStepId === targetId) {
      return;
    }
    const fromIndex = presentation.items.findIndex((item) => item.id === draggingStepId);
    const toIndex = presentation.items.findIndex((item) => item.id === targetId);
    if (fromIndex === -1 || toIndex === -1) {
      return;
    }

    updatePresentation((current) => ({
      ...current,
      items: moveItem(current.items, fromIndex, toIndex),
    }));
    setDraggingStepId(null);
  };

  const startPresentation = async () => {
    if (!presentation || !presentation.items.length || !window.presentApi) {
      return;
    }

    const selectedIndex = selectedStepId ? presentation.items.findIndex((item) => item.id === selectedStepId) : -1;
    const startFromSelected = selectedIndex >= 0 && window.confirm('Start from selected step?\nOK = selected, Cancel = beginning');
    const startIndex = startFromSelected && selectedIndex >= 0 ? selectedIndex : 0;

    try {
      await window.presentApi.startPresentation({
        deckId: presentation.id,
        startIndex,
        displayId: selectedDisplayId === '' ? undefined : selectedDisplayId,
      });
    } catch {
      setToastMessage('Could not start presentation window.');
    }
  };

  const importDeck = async () => {
    if (!window.presentApi) {
      return;
    }
    try {
      const imported = await window.presentApi.deckImport();
      if (!imported) {
        return;
      }
      setPresentation(imported);
      setSelectedStepId(imported.items[0]?.id ?? null);
      setSlideUrls({});
    } catch {
      setToastMessage('Could not import deck file.');
    }
  };

  const exportDeck = async () => {
    if (!presentation || !window.presentApi) {
      return;
    }
    try {
      await window.presentApi.deckExport(presentation.id);
    } catch {
      setToastMessage('Could not export deck file.');
    }
  };

  const openStepExternally = async (step: PresentationStep) => {
    if (step.type !== 'web' || !step.url || !window.presentApi) {
      return;
    }
    try {
      await window.presentApi.openExternal(step.url);
    } catch {
      setToastMessage('Could not open link in external browser.');
    }
  };

  const getSlideUrl = (step: PresentationStep | null): string | null => {
    if (!step?.slideRef || !presentation) {
      return null;
    }
    return slideUrls[`${presentation.id}|${step.slideRef.relativePath}`] || null;
  };

  if (!hasElectronApi) {
    return <div className="loading">Run this project in Electron mode: npm run dev</div>;
  }

  if (!presentation) {
    return <div className="loading">Loading deck...</div>;
  }

  return (
    <div className="app-shell">
      <header className="app-header">
        <input
          className="deck-title-input"
          value={presentation.title}
          onChange={(event) => updatePresentation((current) => ({ ...current, title: event.target.value }))}
          aria-label="Presentation title"
        />
        <div className="header-actions">
          <button onClick={() => void createNewDeck()}>New Deck</button>
          <button onClick={() => void importDeck()}>Import Deck…</button>
          <button onClick={() => void exportDeck()}>Export Deck…</button>
          <select
            value={selectedDisplayId}
            onChange={(event) => setSelectedDisplayId(event.target.value ? Number(event.target.value) : '')}
            title="Presentation display"
          >
            {displayOptions.map((display) => (
              <option key={display.id} value={display.id}>
                {display.label} ({display.width}x{display.height})
              </option>
            ))}
          </select>
          <button onClick={() => void startPresentation()} disabled={presentation.items.length === 0}>
            Play
          </button>
        </div>
      </header>

      <main className="main-layout">
        <aside className="sidebar">
          <div className="sidebar-actions">
            <button
              onClick={() => {
                setIsAddingWebStep((current) => !current);
                if (isAddingWebStep) {
                  setNewWebUrl('');
                  setNewWebTitle('');
                }
              }}
            >
              + Web Step
            </button>
            <button onClick={() => void addSlideSteps()}>+ Slides/Video</button>
          </div>

          {isAddingWebStep ? (
            <div className="add-web-step-form">
              <input
                placeholder="https://example.com"
                value={newWebUrl}
                onChange={(event) => setNewWebUrl(event.target.value)}
              />
              <input
                placeholder="Optional title"
                value={newWebTitle}
                onChange={(event) => setNewWebTitle(event.target.value)}
              />
              <div className="add-web-step-actions">
                <button onClick={submitNewWebStep}>Add</button>
                <button
                  onClick={() => {
                    setIsAddingWebStep(false);
                    setNewWebUrl('');
                    setNewWebTitle('');
                  }}
                >
                  Cancel
                </button>
              </div>
            </div>
          ) : null}

          <div className="step-list">
            {presentation.items.map((item) => {
              const isSelected = item.id === selectedStepId;
              return (
                <div
                  key={item.id}
                  className={`step-item ${isSelected ? 'selected' : ''}`}
                  onClick={() => setSelectedStepId(item.id)}
                  draggable
                  onDragStart={() => setDraggingStepId(item.id)}
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={(event) => onStepDrop(event, item.id)}
                >
                  <span className="step-icon">
                    {item.type === 'web' ? '🌐' : getSlideMediaKind(item.slideRef) === 'video' ? '🎬' : '🖼️'}
                  </span>
                  <span className="step-title">{getStepTitle(item)}</span>
                  <button
                    className="delete-step-button"
                    onClick={(event) => {
                      event.stopPropagation();
                      deleteStep(item.id);
                    }}
                    title="Delete step"
                  >
                    ×
                  </button>
                </div>
              );
            })}
          </div>
        </aside>

        <section className="preview-pane">
          {selectedStep ? (
            <>
              <div className="preview-controls">
                <div className="preview-type">
                  {selectedStep.type === 'web'
                    ? 'Web page'
                    : getSlideMediaKind(selectedStep.slideRef) === 'video'
                      ? 'Slide video clip'
                      : 'Slide image'}
                </div>
                <label>
                  Title
                  <input
                    value={selectedStep.title ?? ''}
                    onChange={(event) => updateSelectedStep((step) => ({ ...step, title: event.target.value }))}
                  />
                </label>
                <label>
                  Notes
                  <textarea
                    value={selectedStep.notes ?? ''}
                    onChange={(event) => updateSelectedStep((step) => ({ ...step, notes: event.target.value }))}
                  />
                </label>
                {selectedStep.type === 'web' ? (
                  <div className="web-controls">
                    <label>
                      URL
                      <input
                        value={selectedStep.url ?? ''}
                        onChange={(event) => updateSelectedStep((step) => ({ ...step, url: event.target.value }))}
                      />
                    </label>
                    <label>
                      Zoom (%)
                      <input
                        type="number"
                        min={MIN_WEB_ZOOM}
                        max={MAX_WEB_ZOOM}
                        step={5}
                        value={normalizeWebZoom(selectedStep.webZoom ?? 100)}
                        onChange={(event) =>
                          updateSelectedStep((step) => ({ ...step, webZoom: normalizeWebZoom(event.target.valueAsNumber) }))
                        }
                      />
                    </label>
                    <div className="inline-actions">
                      <button onClick={() => void openStepExternally(selectedStep)}>Open in external browser</button>
                    </div>
                  </div>
                ) : (
                  <div className="slide-metadata">
                    Source: {selectedStep.slideRef?.sourceFileName || selectedStep.slideRef?.relativePath || 'Unknown'}
                  </div>
                )}
              </div>

              <div className="preview-content">
                {selectedStep.type === 'web' ? (
                  <div className="web-preview-note">
                    Web steps are opened as top-level pages in the presentation window to bypass iframe restrictions.
                  </div>
                ) : getSlideUrl(selectedStep) ? (
                  getSlideMediaKind(selectedStep.slideRef) === 'video' ? (
                    <video className="preview-video" src={getSlideUrl(selectedStep)!} controls preload="metadata" />
                  ) : (
                    <img className="preview-image" src={getSlideUrl(selectedStep)!} alt={getStepTitle(selectedStep)} />
                  )
                ) : (
                  <div className="missing-slide">Missing slide media at {selectedStep.slideRef?.relativePath}</div>
                )}
              </div>
            </>
          ) : (
            <div className="empty-preview">Select a step to preview and edit it.</div>
          )}
        </section>
      </main>

      {toastMessage ? <div className="toast">{toastMessage}</div> : null}
    </div>
  );
}

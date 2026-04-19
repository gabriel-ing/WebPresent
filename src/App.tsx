import { DragEvent, useEffect, useMemo, useRef, useState } from 'react';
import './App.css';
import type { PresentationStep, PptxDeckData, SlideImportMode, SlideRef } from './types';
import { createId, moveItem } from './utils';
import { usePresentation } from './hooks/usePresentation';
import { useSlideUrls } from './hooks/useSlideUrls';
import { StepSidebar } from './components/StepSidebar';
import { StepDetail } from './components/StepDetail';
import { insertAfterSelected } from './stepUtils';

export default function App() {
  const {
    presentation,
    setPresentation,
    displayOptions,
    selectedDisplayId,
    setSelectedDisplayId,
    toastMessage,
    setToastMessage,
    updatePresentation,
  } = usePresentation();

  const [selectedStepId, setSelectedStepId] = useState<string | null>(null);
  const [draggingStepId, setDraggingStepId] = useState<string | null>(null);

  // Select the first step when the deck is first loaded
  const lastDeckIdRef = useRef<string | null>(null);
  useEffect(() => {
    if (presentation && lastDeckIdRef.current !== presentation.id) {
      lastDeckIdRef.current = presentation.id;
      setSelectedStepId(presentation.items[0]?.id ?? null);
    }
  }, [presentation?.id]);
  const [isAddingWebStep, setIsAddingWebStep] = useState(false);
  const [newWebUrl, setNewWebUrl] = useState('');
  const [newWebTitle, setNewWebTitle] = useState('');

  const selectedStep = useMemo(() => {
    if (!presentation || !selectedStepId) return null;
    return presentation.items.find((item) => item.id === selectedStepId) || null;
  }, [presentation, selectedStepId]);

  const slideUrls = useSlideUrls(presentation, selectedStep);

  const hasElectronApi = Boolean(window.presentApi);

  // ── Actions ──────────────────────────────────────────────────────────────────

  const createNewDeck = async () => {
    if (!window.presentApi) return;
    const next = await window.presentApi.deckCreate();
    setPresentation(next);
    // selectedStepId is reset by the useEffect watching presentation.id
  };

  const submitNewWebStep = () => {
    const url = newWebUrl.trim();
    if (!url) { setToastMessage('Please enter a URL.'); return; }
    const step: PresentationStep = {
      id: createId('step'),
      type: 'web',
      title: newWebTitle.trim() || undefined,
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
    if (!presentation || !window.presentApi) return;
    const filePaths = await window.presentApi.pickSlideFiles();
    if (!filePaths.length) return;
    await importSlidesFromPaths(filePaths, 'Import as one animation group?\nOK = grouped, Cancel = separate steps');
  };

  const addSlideDirectorySteps = async () => {
    if (!presentation || !window.presentApi) return;
    const filePaths = await window.presentApi.pickSlideDirectory();
    if (!filePaths.length) { setToastMessage('No supported image files found in that folder.'); return; }
    await importSlidesFromPaths(filePaths, 'Import folder as one animation group?\nOK = grouped, Cancel = separate steps');
  };

  const importSlidesFromPaths = async (filePaths: string[], groupPrompt: string) => {
    if (!presentation || !window.presentApi) return;
    const mode: SlideImportMode =
      filePaths.length > 1 && window.confirm(groupPrompt) ? 'grouped' : 'separate';
    let slideRefs: SlideRef[];
    try {
      slideRefs = await window.presentApi.importSlides({ deckId: presentation.id, filePaths, mode });
    } catch {
      setToastMessage('Could not import slide media files.');
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

  const importPptxSteps = async () => {
    if (!presentation || !window.presentApi) return;
    const filePath = await window.presentApi.pickPptxFile();
    if (!filePath) return;
    let deckData: PptxDeckData;
    try {
      deckData = await window.presentApi.importPptx({ deckId: presentation.id, filePath });
    } catch {
      setToastMessage('Could not parse the PowerPoint file.');
      return;
    }
    const steps: PresentationStep[] = [];
    for (const slide of deckData.slides) {
      const totalSteps = slide.animationStepCount + 1;
      const groupId = totalSteps > 1 ? createId('group') : undefined;
      for (let animStep = 0; animStep < totalSteps; animStep++) {
        const slideNum = slide.slideIndex + 1;
        const defaultTitle = totalSteps > 1 ? `Slide ${slideNum} (build ${animStep})` : `Slide ${slideNum}`;
        steps.push({
          id: createId('step'),
          type: 'pptx-slide',
          title: animStep === 0 ? `Slide ${slideNum}` : defaultTitle,
          notes: slide.notes,
          groupId,
          pptxSlideData: slide,
          pptxAnimationStep: animStep,
        });
      }
    }
    if (!steps.length) { setToastMessage('No slides found in the PowerPoint file.'); return; }
    updatePresentation((current) => ({
      ...current,
      items: insertAfterSelected(current.items, selectedStepId, steps),
    }));
    setSelectedStepId(steps[0]?.id ?? null);
    const animSlides = deckData.slides.filter((s) => s.animationStepCount > 0).length;
    setToastMessage(
      animSlides
        ? `Imported ${deckData.slides.length} slides from ${deckData.sourceFileName} (${animSlides} with animations → ${steps.length} total steps)`
        : `Imported ${deckData.slides.length} slides from ${deckData.sourceFileName}`,
    );
  };

  const updateSelectedStep = (updater: (step: PresentationStep) => PresentationStep) => {
    if (!selectedStepId) return;
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
    if (selectedStepId === stepId) setSelectedStepId(null);
  };

  const onStepDrop = (event: DragEvent<HTMLDivElement>, targetId: string) => {
    event.preventDefault();
    if (!draggingStepId || !presentation || draggingStepId === targetId) return;
    const fromIndex = presentation.items.findIndex((item) => item.id === draggingStepId);
    const toIndex = presentation.items.findIndex((item) => item.id === targetId);
    if (fromIndex === -1 || toIndex === -1) return;
    updatePresentation((current) => ({
      ...current,
      items: moveItem(current.items, fromIndex, toIndex),
    }));
    setDraggingStepId(null);
  };

  const startPresentation = async () => {
    if (!presentation || !presentation.items.length || !window.presentApi) return;
    const selectedIndex = selectedStepId
      ? presentation.items.findIndex((item) => item.id === selectedStepId)
      : -1;
    const startFromSelected =
      selectedIndex >= 0 && window.confirm('Start from selected step?\nOK = selected, Cancel = beginning');
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
    if (!window.presentApi) return;
    try {
      const imported = await window.presentApi.deckImport();
      if (!imported) return;
      setPresentation(imported);
      // selectedStepId is reset by the useEffect watching presentation.id
    } catch {
      setToastMessage('Could not import deck file.');
    }
  };

  const exportDeck = async () => {
    if (!presentation || !window.presentApi) return;
    try {
      await window.presentApi.deckExport(presentation.id);
    } catch {
      setToastMessage('Could not export deck file.');
    }
  };

  const openStepExternally = async (step: PresentationStep) => {
    if (step.type !== 'web' || !step.url || !window.presentApi) return;
    try {
      await window.presentApi.openExternal(step.url);
    } catch {
      setToastMessage('Could not open link in external browser.');
    }
  };

  // ── Guards ───────────────────────────────────────────────────────────────────

  if (!hasElectronApi) {
    return <div className="loading">Run this project in Electron mode: npm run dev</div>;
  }

  if (!presentation) {
    return <div className="loading">Loading deck...</div>;
  }

  // ── Render ───────────────────────────────────────────────────────────────────

  return (
    <div className="app-shell">
      <header className="app-header">
        <input
          className="deck-title-input"
          value={presentation.title}
          onChange={(e) =>
            updatePresentation((current) => ({ ...current, title: e.target.value }))
          }
          aria-label="Presentation title"
        />
        <div className="header-actions">
          <button onClick={() => void createNewDeck()}>New Deck</button>
          <button onClick={() => void importDeck()}>Import Deck…</button>
          <button onClick={() => void exportDeck()}>Export Deck…</button>
          <select
            value={selectedDisplayId}
            onChange={(e) => setSelectedDisplayId(e.target.value ? Number(e.target.value) : '')}
            title="Presentation display"
          >
            {displayOptions.map((display) => (
              <option key={display.id} value={display.id}>
                {display.label} ({display.width}x{display.height})
              </option>
            ))}
          </select>
          <button
            onClick={() => void startPresentation()}
            disabled={presentation.items.length === 0}
          >
            Play
          </button>
        </div>
      </header>

      <main className="main-layout">
        <StepSidebar
          presentation={presentation}
          selectedStepId={selectedStepId}
          isAddingWebStep={isAddingWebStep}
          newWebUrl={newWebUrl}
          newWebTitle={newWebTitle}
          onSelectStep={setSelectedStepId}
          onDeleteStep={deleteStep}
          onDragStart={setDraggingStepId}
          onDrop={onStepDrop}
          onToggleAddWebStep={() => {
            setIsAddingWebStep((v) => !v);
            if (isAddingWebStep) { setNewWebUrl(''); setNewWebTitle(''); }
          }}
          onNewWebUrlChange={setNewWebUrl}
          onNewWebTitleChange={setNewWebTitle}
          onSubmitNewWebStep={submitNewWebStep}
          onCancelAddWebStep={() => {
            setIsAddingWebStep(false);
            setNewWebUrl('');
            setNewWebTitle('');
          }}
          onAddSlideSteps={() => void addSlideSteps()}
          onAddSlideDirectorySteps={() => void addSlideDirectorySteps()}
          onImportPptxSteps={() => void importPptxSteps()}
        />

        <section className="preview-pane">
          {selectedStep ? (
            <StepDetail
              selectedStep={selectedStep}
              presentation={presentation}
              slideUrls={slideUrls}
              onUpdateStep={updateSelectedStep}
              onOpenExternally={(step) => void openStepExternally(step)}
            />
          ) : (
            <div className="empty-preview">Select a step to preview and edit it.</div>
          )}
        </section>
      </main>

      {toastMessage ? <div className="toast">{toastMessage}</div> : null}
    </div>
  );
}

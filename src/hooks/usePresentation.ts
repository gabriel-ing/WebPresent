import { useEffect, useRef, useState } from 'react';
import type { DisplayInfo, Presentation } from '../types';
import { nowIso } from '../utils';

const AUTOSAVE_DEBOUNCE_MS = 400;

export interface UsePresentationResult {
  presentation: Presentation | null;
  setPresentation: React.Dispatch<React.SetStateAction<Presentation | null>>;
  displayOptions: DisplayInfo[];
  selectedDisplayId: number | '';
  setSelectedDisplayId: React.Dispatch<React.SetStateAction<number | ''>>;
  toastMessage: string | null;
  setToastMessage: React.Dispatch<React.SetStateAction<string | null>>;
  updatePresentation: (updater: (prev: Presentation) => Presentation) => void;
}

export function usePresentation(): UsePresentationResult {
  const [presentation, setPresentation] = useState<Presentation | null>(null);
  const [displayOptions, setDisplayOptions] = useState<DisplayInfo[]>([]);
  const [selectedDisplayId, setSelectedDisplayId] = useState<number | ''>('');
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const autosaveTimerRef = useRef<number | null>(null);

  // Initial load
  useEffect(() => {
    let mounted = true;
    async function init() {
      if (!window.presentApi) {
        setToastMessage('Electron bridge not found. Run with npm run dev for desktop mode.');
        return;
      }
      const [loaded, displays] = await Promise.all([
        window.presentApi.deckGetCurrent(),
        window.presentApi.getDisplays(),
      ]);
      if (!mounted) return;
      setPresentation(loaded);
      setDisplayOptions(displays);
      if (displays[0]) setSelectedDisplayId(displays[0].id);
    }
    void init();
    return () => { mounted = false; };
  }, []);

  // Autosave
  useEffect(() => {
    if (!presentation || !window.presentApi) return;
    if (autosaveTimerRef.current) window.clearTimeout(autosaveTimerRef.current);
    autosaveTimerRef.current = window.setTimeout(async () => {
      try {
        await window.presentApi!.deckSave(presentation);
      } catch {
        setToastMessage('Autosave failed. Try exporting your deck.');
      }
    }, AUTOSAVE_DEBOUNCE_MS);
  }, [presentation]);

  // Toast auto-dismiss
  useEffect(() => {
    if (!toastMessage) return;
    const timer = window.setTimeout(() => setToastMessage(null), 3500);
    return () => window.clearTimeout(timer);
  }, [toastMessage]);

  const updatePresentation = (updater: (prev: Presentation) => Presentation) => {
    setPresentation((current) => {
      if (!current) return current;
      return { ...updater(current), updatedAt: nowIso() };
    });
  };

  return {
    presentation,
    setPresentation,
    displayOptions,
    selectedDisplayId,
    setSelectedDisplayId,
    toastMessage,
    setToastMessage,
    updatePresentation,
  };
}

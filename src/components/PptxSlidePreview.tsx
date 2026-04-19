import { useCallback, useMemo, useState } from 'react';
import type { Presentation, PresentationStep } from '../types';
import { renderPptxSlidePreviewHtml } from '../pptxRenderer';

interface Props {
  step: PresentationStep;
  presentation: Presentation;
  slideUrls: Record<string, string>;
  onUpdateStep: (updater: (step: PresentationStep) => PresentationStep) => void;
}

export function PptxSlidePreview({ step, presentation, slideUrls, onUpdateStep }: Props) {
  const [editingShapeId, setEditingShapeId] = useState<string | null>(null);
  const [editText, setEditText] = useState('');

  const slideData = step.pptxSlideData;
  const animStep = step.pptxAnimationStep ?? 0;

  const mediaResolver = useCallback(
    (relativePath: string): string => {
      if (!relativePath) return '';
      return slideUrls[`${presentation.id}|${relativePath}`] || '';
    },
    [slideUrls, presentation.id],
  );

  const previewHtml = useMemo(() => {
    if (!slideData) return '';
    return renderPptxSlidePreviewHtml(slideData, animStep, mediaResolver);
  }, [slideData, animStep, mediaResolver]);

  const startEditingShape = (shapeId: string) => {
    if (!slideData) return;
    const shape = slideData.shapes.find((s) => s.id === shapeId);
    if (!shape?.paragraphs?.length) return;
    const text = shape.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
    setEditingShapeId(shapeId);
    setEditText(text);
  };

  const saveEditedText = () => {
    if (!editingShapeId || !slideData) return;
    const shapeIndex = slideData.shapes.findIndex((s) => s.id === editingShapeId);
    if (shapeIndex === -1) return;

    const lines = editText.split('\n');
    const shape = slideData.shapes[shapeIndex];
    const updatedParagraphs = lines.map((line, i) => {
      const existingPara = shape.paragraphs?.[i];
      const existingRun = existingPara?.runs?.[0];
      return {
        runs: [{ ...existingRun, text: line }],
        alignment: existingPara?.alignment || 'left',
        bulletType: existingPara?.bulletType || ('none' as const),
        level: existingPara?.level || 0,
      };
    });

    onUpdateStep((s) => {
      if (!s.pptxSlideData) return s;
      const newShapes = [...s.pptxSlideData.shapes];
      newShapes[shapeIndex] = { ...newShapes[shapeIndex], paragraphs: updatedParagraphs };
      return { ...s, pptxSlideData: { ...s.pptxSlideData, shapes: newShapes } };
    });
    setEditingShapeId(null);
    setEditText('');
  };

  if (!slideData) {
    return <div className="missing-slide">No slide data available.</div>;
  }

  const textShapes = slideData.shapes.filter(
    (s) => s.paragraphs?.length && s.paragraphs.some((p) => p.runs.some((r) => r.text.trim())),
  );

  return (
    <div className="pptx-preview-container">
      <iframe
        className="pptx-preview-iframe"
        srcDoc={previewHtml}
        sandbox="allow-scripts allow-same-origin"
        title="Slide preview"
      />
      {textShapes.length > 0 && (
        <div className="pptx-text-edit-panel">
          <div className="pptx-text-edit-header">Edit text (click a shape):</div>
          {textShapes.map((shape) => {
            const shapeText = shape.paragraphs!
              .map((p) => p.runs.map((r) => r.text).join(''))
              .join(' ')
              .slice(0, 80);
            const isEditing = editingShapeId === shape.id;

            return (
              <div key={shape.id} className="pptx-text-shape-item">
                {isEditing ? (
                  <div className="pptx-inline-edit">
                    <textarea
                      value={editText}
                      onChange={(e) => setEditText(e.target.value)}
                      rows={Math.min(8, editText.split('\n').length + 1)}
                    />
                    <div className="pptx-inline-edit-actions">
                      <button onClick={saveEditedText}>Save</button>
                      <button onClick={() => { setEditingShapeId(null); setEditText(''); }}>Cancel</button>
                    </div>
                  </div>
                ) : (
                  <button
                    className="pptx-text-shape-button"
                    onClick={() => startEditingShape(shape.id)}
                    title="Click to edit text"
                  >
                    <span className="pptx-shape-label">{shape.name || shape.id}</span>
                    <span className="pptx-shape-text-preview">{shapeText || '(empty)'}</span>
                  </button>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

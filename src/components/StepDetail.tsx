import type { Presentation, PresentationStep } from '../types';
import { getSlideMediaKind, getStepTitle } from '../stepUtils';
import { PptxSlidePreview } from './PptxSlidePreview';

const MIN_WEB_ZOOM = 25;
const MAX_WEB_ZOOM = 300;

function normalizeWebZoom(value: number): number {
  if (!Number.isFinite(value)) return 100;
  return Math.min(MAX_WEB_ZOOM, Math.max(MIN_WEB_ZOOM, Math.round(value)));
}

interface Props {
  selectedStep: PresentationStep;
  presentation: Presentation;
  slideUrls: Record<string, string>;
  onUpdateStep: (updater: (step: PresentationStep) => PresentationStep) => void;
  onOpenExternally: (step: PresentationStep) => void;
}

export function StepDetail({ selectedStep, presentation, slideUrls, onUpdateStep, onOpenExternally }: Props) {
  const slideUrl = selectedStep.slideRef
    ? slideUrls[`${presentation.id}|${selectedStep.slideRef.relativePath}`] || null
    : null;

  const stepTypeLabel =
    selectedStep.type === 'web'
      ? 'Web page'
      : selectedStep.type === 'pptx-slide'
        ? 'PowerPoint slide'
        : getSlideMediaKind(selectedStep.slideRef) === 'video'
          ? 'Slide video clip'
          : 'Slide image';

  return (
    <>
      <div className="preview-controls">
        <div className="preview-type">{stepTypeLabel}</div>
        <label>
          Title
          <input
            value={selectedStep.title ?? ''}
            onChange={(e) => onUpdateStep((s) => ({ ...s, title: e.target.value }))}
          />
        </label>
        <label>
          Notes
          <textarea
            value={selectedStep.notes ?? ''}
            onChange={(e) => onUpdateStep((s) => ({ ...s, notes: e.target.value }))}
          />
        </label>

        {selectedStep.type === 'web' && (
          <div className="web-controls">
            <label>
              URL
              <input
                value={selectedStep.url ?? ''}
                onChange={(e) => onUpdateStep((s) => ({ ...s, url: e.target.value }))}
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
                onChange={(e) =>
                  onUpdateStep((s) => ({ ...s, webZoom: normalizeWebZoom(e.target.valueAsNumber) }))
                }
              />
            </label>
            <div className="inline-actions">
              <button onClick={() => onOpenExternally(selectedStep)}>Open in external browser</button>
            </div>
          </div>
        )}

        {selectedStep.type === 'pptx-slide' && (
          <div className="slide-metadata">
            Source: PowerPoint slide {(selectedStep.pptxSlideData?.slideIndex ?? 0) + 1}
            {selectedStep.pptxAnimationStep ? ` · Build step ${selectedStep.pptxAnimationStep}` : ''}
            {selectedStep.pptxSlideData?.animationStepCount
              ? ` · ${selectedStep.pptxSlideData.animationStepCount} animation(s)`
              : ''}
          </div>
        )}

        {selectedStep.type === 'slide' && (
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
        ) : selectedStep.type === 'pptx-slide' && selectedStep.pptxSlideData ? (
          <PptxSlidePreview
            step={selectedStep}
            presentation={presentation}
            slideUrls={slideUrls}
            onUpdateStep={onUpdateStep}
          />
        ) : slideUrl ? (
          getSlideMediaKind(selectedStep.slideRef) === 'video' ? (
            <video className="preview-video" src={slideUrl} controls preload="metadata" />
          ) : (
            <img className="preview-image" src={slideUrl} alt={getStepTitle(selectedStep)} />
          )
        ) : (
          <div className="missing-slide">
            Missing slide media at {selectedStep.slideRef?.relativePath}
          </div>
        )}
      </div>
    </>
  );
}

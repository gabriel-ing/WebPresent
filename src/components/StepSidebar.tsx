import { DragEvent } from 'react';
import type { Presentation } from '../types';
import { getSlideMediaKind, getStepTitle } from '../stepUtils';

interface Props {
  presentation: Presentation;
  selectedStepId: string | null;
  isAddingWebStep: boolean;
  newWebUrl: string;
  newWebTitle: string;
  onSelectStep: (id: string) => void;
  onDeleteStep: (id: string) => void;
  onDragStart: (id: string) => void;
  onDrop: (event: DragEvent<HTMLDivElement>, targetId: string) => void;
  onToggleAddWebStep: () => void;
  onNewWebUrlChange: (value: string) => void;
  onNewWebTitleChange: (value: string) => void;
  onSubmitNewWebStep: () => void;
  onCancelAddWebStep: () => void;
  onAddSlideSteps: () => void;
  onAddSlideDirectorySteps: () => void;
  onImportPptxSteps: () => void;
}

export function StepSidebar({
  presentation,
  selectedStepId,
  isAddingWebStep,
  newWebUrl,
  newWebTitle,
  onSelectStep,
  onDeleteStep,
  onDragStart,
  onDrop,
  onToggleAddWebStep,
  onNewWebUrlChange,
  onNewWebTitleChange,
  onSubmitNewWebStep,
  onCancelAddWebStep,
  onAddSlideSteps,
  onAddSlideDirectorySteps,
  onImportPptxSteps,
}: Props) {
  return (
    <aside className="sidebar">
      <div className="sidebar-actions">
        <button onClick={onToggleAddWebStep}>+ Web Step</button>
        <button onClick={onAddSlideSteps}>+ Slides/Video</button>
        <button onClick={onAddSlideDirectorySteps}>+ Slide Folder</button>
        <button onClick={onImportPptxSteps}>+ Import PPTX</button>
      </div>

      {isAddingWebStep && (
        <div className="add-web-step-form">
          <input
            placeholder="https://example.com"
            value={newWebUrl}
            onChange={(e) => onNewWebUrlChange(e.target.value)}
          />
          <input
            placeholder="Optional title"
            value={newWebTitle}
            onChange={(e) => onNewWebTitleChange(e.target.value)}
          />
          <div className="add-web-step-actions">
            <button onClick={onSubmitNewWebStep}>Add</button>
            <button onClick={onCancelAddWebStep}>Cancel</button>
          </div>
        </div>
      )}

      <div className="step-list">
        {presentation.items.map((item) => {
          const isSelected = item.id === selectedStepId;
          return (
            <div
              key={item.id}
              className={`step-item ${isSelected ? 'selected' : ''}`}
              onClick={() => onSelectStep(item.id)}
              draggable
              onDragStart={() => onDragStart(item.id)}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => onDrop(e, item.id)}
            >
              <span className="step-icon">
                {item.type === 'web'
                  ? '🌐'
                  : item.type === 'pptx-slide'
                    ? '📊'
                    : getSlideMediaKind(item.slideRef) === 'video'
                      ? '🎬'
                      : '🖼️'}
              </span>
              <span className="step-title">{getStepTitle(item)}</span>
              <button
                className="delete-step-button"
                onClick={(e) => {
                  e.stopPropagation();
                  onDeleteStep(item.id);
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
  );
}

# Plan: PowerPoint (.pptx) Import for WebPresent

## 1. Goal

Allow users to import `.pptx` files directly into WebPresent so that slides are rendered natively in the app — editable (at least text), animated (basic transitions/builds), and fully interleaved with existing web-step and image-slide capabilities.

---

## 2. Current Architecture Summary

| Concept | How it works today |
|---|---|
| **Step types** | `web` (URL loaded as top-level page) and `slide` (image/video file stored in `<deckDir>/slides/`) |
| **Data model** | `Presentation → PresentationStep[]`, each step has optional `slideRef` (file on disk) or `url` |
| **Presentation window** | Electron `BrowserWindow` loads URLs or file:// paths directly; slides are raw image/video files displayed with injected CSS |
| **Editor** | React SPA; sidebar list, preview pane, title/notes/URL editing |
| **Persistence** | Deck JSON + slide media files on disk; export/import as `.presentdeck` zip |
| **Animations** | Simulated by grouping multiple images into a `groupId` (user exports each animation state as a separate image from PowerPoint) |

### Key constraint
Web steps are shown as **top-level pages** (not iframes) to bypass CSP/X-Frame-Options. Slide steps are loaded as file:// URLs. The presentation window is a single `BrowserWindow` that navigates between pages.

---

## 3. High-Level Approach

Introduce a **third step type: `pptx-slide`** that stores parsed slide data (shapes, text, layout) and renders each slide as an **HTML/CSS/SVG document** at presentation time. This keeps the existing `web` and `slide` step types untouched and adds PowerPoint support alongside them.

### Why render slides as HTML rather than converting to images?
- **Editability**: text in HTML `<div>`s / `<span>`s can be edited in the editor preview pane.
- **Animations**: build steps (appear/fade/fly-in) map naturally to CSS animations/transitions driven by step advancement.
- **Scalability**: vector text stays crisp at any resolution; no large raster files.
- **Interoperability**: the rendered HTML can be loaded in the presentation `BrowserWindow` exactly like a web step (via `data:` URL or a temp HTML file).

---

## 4. Parsing the .pptx File

### 4.1 Library choice

| Library | Pros | Cons |
|---|---|---|
| **[pptxgenjs](https://github.com/nicoleahmed/pptxgenjs)** | Popular, mature | Designed for *generation*, not *parsing* |
| **Manual XML parsing** (`adm-zip` + `fast-xml-parser`) | Full control; `adm-zip` is already a dependency | More work; must understand OOXML schema |
| **[pptx2json](https://github.com/nicoleahmed/pptx2json)** | Purpose-built for reading .pptx | Small community; incomplete |
| **Custom parser on `adm-zip` + `fast-xml-parser`** ✅ | Leverages existing dep; can target exactly the features we need | Needs OOXML knowledge; iterative |

**Recommendation:** Build a custom parser using the already-installed `adm-zip` plus `fast-xml-parser` (lightweight, fast XML→JSON). This avoids large dependencies and gives full control over which OOXML features to support.

### 4.2 What lives inside a .pptx

A `.pptx` is a ZIP containing:

```
[Content_Types].xml
_rels/.rels
ppt/
  presentation.xml          ← slide order, slide size
  _rels/presentation.xml.rels
  slideMasters/             ← master layouts
  slideLayouts/             ← layout templates
  slides/
    slide1.xml ... slideN.xml   ← individual slide content
    _rels/slide1.xml.rels       ← media references per slide
  media/                    ← embedded images, videos
  theme/theme1.xml          ← colour palette, fonts
```

### 4.3 Parsing pipeline (runs in Electron main process)

```
.pptx file
  → adm-zip extracts in memory
  → parse presentation.xml  → slide dimensions, slide order
  → parse theme/theme1.xml  → colour scheme, font defaults
  → for each slide:
      parse slideN.xml      → shape tree (sp, pic, graphicFrame, grpSp)
      parse slideN.xml.rels → resolve image/media rIds
      extract referenced media from ppt/media/
      → produce a SlideData object (see §5)
```

### 4.4 Supported OOXML elements (Phase 1 → Phase 2)

| Element | Phase 1 (MVP) | Phase 2 |
|---|---|---|
| `<p:sp>` (shapes: rect, ellipse, rounded rect) | ✅ background fill, border, text | ✅ more shape presets |
| `<p:sp>` text body `<a:p>` / `<a:r>` | ✅ font, size, bold, italic, colour, alignment | ✅ bullet lists, numbered lists |
| `<p:pic>` (images) | ✅ position, size, embedded PNG/JPEG | ✅ cropping, effects |
| `<p:graphicFrame>` (tables) | ❌ | ✅ basic table rendering |
| `<p:grpSp>` (grouped shapes) | ❌ flatten to individual shapes | ✅ proper group transforms |
| Slide background | ✅ solid fill, image fill | ✅ gradient fill |
| Slide master / layout inheritance | ✅ background + placeholders | ✅ full inheritance chain |
| Theme colours | ✅ map scheme colours to hex | ✅ tints/shades |
| Animations (`<p:timing>`) | ✅ basic appear/fade on click | ✅ fly-in, wipe, emphasis |
| Slide transitions | ❌ | ✅ fade/push between slides |
| SmartArt | ❌ render as fallback image | ❌ |
| Charts | ❌ render as fallback image | ❌ |
| Embedded video/audio | ❌ | ✅ |

---

## 5. Internal Data Model for Parsed Slides

### 5.1 New types

```typescript
/* ---- New step type ---- */
export type StepType = 'web' | 'slide' | 'pptx-slide';

/* ---- Parsed slide shape ---- */
export type PptxShapeType = 'rect' | 'ellipse' | 'roundRect' | 'image' | 'line' | 'freeform';

export type PptxTextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;        // in pt
  fontFamily?: string;
  colour?: string;          // #RRGGBB
  highlightColour?: string;
};

export type PptxParagraph = {
  runs: PptxTextRun[];
  alignment?: 'left' | 'center' | 'right' | 'justify';
  bulletType?: 'none' | 'bullet' | 'numbered';
  level?: number;           // indentation level (0-based)
};

export type PptxFill = {
  type: 'solid' | 'image' | 'gradient' | 'none';
  colour?: string;
  imageRelativePath?: string;   // path within deck slides/ dir
  gradientStops?: { position: number; colour: string }[];
};

export type PptxBorder = {
  width: number;            // pt
  colour: string;
  style?: 'solid' | 'dashed' | 'dotted';
};

export type PptxShape = {
  id: string;
  type: PptxShapeType;
  x: number;                // EMUs converted to px (or %)
  y: number;
  width: number;
  height: number;
  rotation?: number;        // degrees
  fill?: PptxFill;
  border?: PptxBorder;
  paragraphs?: PptxParagraph[];
  imageRelativePath?: string;   // for image shapes
  cornerRadius?: number;
  animationGroup?: number;      // 0 = always visible, 1+ = click-to-reveal sequence
  animationEffect?: 'appear' | 'fade' | 'fly-left' | 'fly-right' | 'fly-up' | 'fly-down';
};

export type PptxSlideData = {
  slideIndex: number;
  width: number;            // px (from presentation.xml slide size)
  height: number;
  background?: PptxFill;
  shapes: PptxShape[];
  notes?: string;
  animationStepCount: number;  // how many click-steps this slide has (0 = single state)
};

export type PptxDeckData = {
  sourceFileName: string;
  slides: PptxSlideData[];
  theme?: {
    colours: Record<string, string>;
    defaultFont?: string;
  };
};
```

### 5.2 Updated `PresentationStep`

```typescript
export type PresentationStep = {
  id: string;
  type: StepType;
  title?: string;
  notes?: string;
  groupId?: string;
  url?: string;                     // web steps
  webZoom?: number;                 // web steps
  slideRef?: SlideRef;              // image/video slide steps
  pptxSlideData?: PptxSlideData;    // NEW: parsed pptx slide
  pptxAnimationStep?: number;       // NEW: which build-step within the slide (0 = base)
};
```

### 5.3 How animations create multiple steps

A single PowerPoint slide with 3 click-animations becomes **4 `PresentationStep`s** in the deck:

| Step | `pptxAnimationStep` | Visible shapes |
|---|---|---|
| Base | 0 | Only shapes with `animationGroup === 0` |
| Click 1 | 1 | Groups 0 + 1 |
| Click 2 | 2 | Groups 0 + 1 + 2 |
| Click 3 | 3 | Groups 0 + 1 + 2 + 3 |

These share the same `groupId` so the sidebar can collapse/expand them, matching the existing animation-group UI pattern.

---

## 6. Rendering Slides as HTML

### 6.1 Renderer function

A pure function `renderPptxSlideToHtml(slide: PptxSlideData, animationStep: number, mediaResolver: (path: string) => string): string` that:

1. Creates an HTML document with a root `<div>` sized to `slide.width × slide.height`.
2. Applies the slide background.
3. For each shape where `shape.animationGroup <= animationStep`:
   - Positions it absolutely (`left`, `top`, `width`, `height`, `transform: rotate()`).
   - Renders text paragraphs as styled `<p>` / `<span>` elements.
   - Renders images as `<img>` tags.
   - Applies fill, border, corner radius.
4. For shapes at exactly `animationGroup === animationStep`, applies a CSS animation class (e.g. `fade-in 0.4s ease`).
5. Wraps everything in a responsive viewport (`<meta name="viewport">` + scaling CSS) so the slide scales to fill the presentation window.

### 6.2 Where rendering happens

| Context | Method |
|---|---|
| **Editor preview pane** | Render into a `<div>` using React (or `dangerouslySetInnerHTML` of the HTML string inside a sandboxed container) |
| **Presentation window** | Generate a complete HTML string → load via `data:text/html;charset=utf-8,...` URL or write to a temp `.html` file and load via `file://` (same pattern as current slide steps) |

### 6.3 Media resolution

Embedded images from the .pptx are extracted and stored in `<deckDir>/slides/` alongside existing slide images. The `mediaResolver` callback maps `imageRelativePath` → `file://` URL (presentation) or base64 data URL (editor preview).

---

## 7. Editing Support

### 7.1 Text editing in the editor pane

When a `pptx-slide` step is selected:

1. The preview renders the slide HTML.
2. Clicking a text shape enters **inline edit mode**: the shape's text becomes a `contentEditable` region (or a textarea overlay positioned over the shape).
3. On blur/save, the edited text is written back to `pptxSlideData.shapes[i].paragraphs[j].runs[k].text`.
4. The updated `PresentationStep` is saved via the normal autosave flow.

### 7.2 Scope of edits (Phase 1)

- ✅ Edit text content of any text shape.
- ✅ Change title of the step.
- ✅ Edit speaker notes.
- ❌ Move/resize shapes (Phase 2).
- ❌ Add/delete shapes (Phase 2).
- ❌ Change formatting (bold, colour) via UI controls (Phase 2).

---

## 8. Import Flow (User Experience)

### 8.1 UI changes

1. **New button** in sidebar actions: **`+ Import PPTX`**
2. Clicking it opens a native file picker (`.pptx` filter).
3. A progress indicator shows while parsing ("Importing 24 slides…").
4. After import, steps are inserted after the currently selected step (matching existing insert behaviour).
5. A toast confirms: "Imported 24 slides from presentation.pptx (3 with animations → 31 total steps)".

### 8.2 Main process IPC

New IPC channels:

```
deck:pickPptxFile       → opens file dialog filtered to .pptx
deck:importPptx         → { deckId, filePath } → parses pptx, extracts media,
                           returns PptxDeckData
```

### 8.3 Renderer flow

```
User clicks "+ Import PPTX"
  → renderer calls presentApi.pickPptxFile()
  → main process shows dialog, returns filePath (or null)
  → renderer calls presentApi.importPptx({ deckId, filePath })
  → main process:
      1. Unzips .pptx
      2. Parses XML → PptxDeckData
      3. Extracts media → deckDir/slides/
      4. Returns PptxDeckData
  → renderer converts PptxDeckData into PresentationStep[] (expanding animations)
  → inserts steps into presentation.items
  → autosave triggers
```

---

## 9. Persistence & Export Compatibility

### 9.1 Deck JSON

`pptxSlideData` is serialized directly into the deck JSON. This is self-describing — no need to keep the original .pptx file.

### 9.2 `.presentdeck` export

The existing zip-based export already bundles `deck.json` + `slides/*`. Since parsed pptx media is stored in `slides/`, exports will include everything needed. No changes to the export/import format are required beyond the deck JSON now containing `pptxSlideData` fields.

### 9.3 Backward compatibility

Older versions of WebPresent that don't know about `type: 'pptx-slide'` will skip/ignore those steps. This is acceptable degradation. A `formatVersion` bump in the deck JSON can signal the new capability.

---

## 10. Implementation Phases

### Phase 1 — MVP (Core Import + Display + Basic Edit)

| # | Task | Files affected | Effort |
|---|---|---|---|
| 1 | Add `fast-xml-parser` dependency | `package.json` | S |
| 2 | Build PPTX parser module | `electron/pptxParser.cjs` (new) | L |
| 3 | Parse slide dimensions, backgrounds, basic shapes (rect, text, images) | `electron/pptxParser.cjs` | L |
| 4 | Parse theme colours + default font | `electron/pptxParser.cjs` | M |
| 5 | Extract embedded media to deck slides dir | `electron/pptxParser.cjs` | M |
| 6 | Parse basic animations (`<p:timing>` appear/fade on click) | `electron/pptxParser.cjs` | L |
| 7 | Add new types (`PptxSlideData`, `PptxShape`, etc.) | `src/types.ts` | M |
| 8 | Update `StepType` to include `'pptx-slide'` | `src/types.ts` | S |
| 9 | Add IPC channels (`deck:pickPptxFile`, `deck:importPptx`) | `electron/main.cjs`, `electron/preload.cjs`, `src/electron.d.ts` | M |
| 10 | Build HTML slide renderer function | `src/pptxRenderer.ts` (new) | L |
| 11 | Add "Import PPTX" button + import flow in editor | `src/App.tsx` | M |
| 12 | Render pptx-slide preview in editor pane | `src/App.tsx` | M |
| 13 | Handle pptx-slide steps in presentation window (`showPresentationStep`) | `electron/main.cjs` | M |
| 14 | Inline text editing for pptx-slide shapes in editor | `src/App.tsx` or new `src/PptxSlideEditor.tsx` | L |
| 15 | Update sidebar icons/labels for pptx-slide steps | `src/App.tsx`, `src/App.css` | S |
| 16 | Test with real .pptx files; iterate on parser fidelity | — | L |

**Estimated total Phase 1: ~3–4 weeks of focused work**

### Phase 2 — Enhanced Fidelity

- More shape presets (arrows, stars, callouts)
- Table rendering
- Bullet and numbered lists
- Gradient fills
- Group shape transforms
- Slide master/layout full inheritance
- More animation effects (fly-in, wipe, zoom)
- Slide transitions (fade between slides)

### Phase 3 — Rich Editing

- Move/resize shapes via drag handles
- Formatting toolbar (bold, italic, colour, font size)
- Add/delete shapes
- Undo/redo stack

---

## 11. Risks & Mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| OOXML is extremely complex; many .pptx files use obscure features | Slides render incorrectly | Start with a curated set of test files; provide a "fallback to image" option — user can re-export slides as images for unsupported content |
| Large .pptx files (100+ slides, many high-res images) slow down parsing | Poor UX on import | Parse in a worker/child process; show progress; extract media lazily |
| Text layout differences (PowerPoint uses proprietary text metrics) | Text overflows or wraps differently | Use generous bounding boxes; allow manual text tweaks via editing |
| Animation parsing is incomplete | Missing build steps | Fall back to showing all shapes at once (animationGroup = 0) when parsing fails |
| Deck JSON grows large with inline `pptxSlideData` | Slow save/load | Consider moving `pptxSlideData` to separate per-slide JSON files if decks exceed ~10 MB |

---

## 12. Fallback Strategy

For any slide that can't be parsed adequately, offer an **automatic image fallback**:

1. During import, optionally render each slide to a PNG using a headless Chromium screenshot of the generated HTML.
2. Store the PNG as a regular `slide` step.
3. The user can choose per-slide whether to use the parsed HTML version or the image fallback.

This ensures that even with an imperfect parser, every imported slide is displayable.

---

## 13. File / Module Map (New & Modified)

```
electron/
  main.cjs              ← add IPC handlers for pptx import
  preload.cjs           ← expose new API methods
  pptxParser.cjs        ← NEW: .pptx parsing logic

src/
  types.ts              ← add PptxSlideData, PptxShape, etc.
  electron.d.ts         ← add new API type signatures
  App.tsx               ← add Import PPTX button, handle new step type
  App.css               ← styles for pptx-slide preview & editor
  pptxRenderer.ts       ← NEW: PptxSlideData → HTML string
  PptxSlideEditor.tsx   ← NEW: inline text editing component
  PptxSlidePreview.tsx  ← NEW: preview component for editor pane

package.json            ← add fast-xml-parser dependency
```

---

## 14. Summary

The plan introduces PowerPoint import as a **first-class step type** that converts `.pptx` XML into a structured intermediate representation (`PptxSlideData`), renders slides as responsive HTML/CSS, supports basic click-to-reveal animations by expanding them into multiple grouped steps, and enables inline text editing. The approach integrates cleanly with WebPresent's existing architecture — parsed slides are stored in the deck JSON, media files go into the existing `slides/` directory, and the presentation window loads generated HTML the same way it loads any other content.

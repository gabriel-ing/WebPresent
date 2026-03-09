1. Project Overview
1.1. Product summary
A browser-based presentation tool where each step in a deck is either:

an interactive web page (shown in an iframe), or
a slide image (exported from PowerPoint, etc.),

with:

a simple editor view (sidebar list + preview),
a fullscreen presentation mode for live talks,
autosave using browser storage,
and self-contained deck files that bundle metadata + slide images into a single JSON file for export/import.

1.2. Primary user / usage pattern

Single user (you), technical, comfortable with file pickers, etc.
Typical workflow:

Create a new deck.
Add a sequence of web pages and exported slide images.
Optionally group multiple slide images to simulate animations (e.g. bullet points appearing).
Reorder and tweak titles.
Present fullscreen using keyboard controls.
Export the deck to a single file for backup or moving between machines.




2. Functional Requirements
2.1. Core features


Create / open a deck

Start with an empty deck or open a previously exported deck file.



Add steps

Add Web Step:

User supplies a URL.
App creates a new step of type web.


Add Slide Steps (images):

User can select one or more image files (.png, .jpg, .jpeg).
Each image becomes a slide step.
User can choose to treat multiple imported images as:

separate independent steps, or
a group representing one conceptual slide with multiple “animation” states.







View & edit steps

Sidebar shows an ordered list of all steps with type icon and title.
Selecting a step updates the preview pane.
User can:

Change a step’s title.
For web steps: change the URL.
For slide steps: view which source file it came from (for reference only).





Reorder steps

User can drag-and-drop items in the sidebar to change sequence order.



Delete steps

User can delete one step at a time from a context menu or a delete button.



Presentation mode

Fullscreen mode.
Shows only the current step’s content.
Keyboard navigation:

Right Arrow or Space → next step
Left Arrow → previous step
Esc → exit presentation


For web steps:

Content shown in an iframe.
Scrollable with mouse/trackpad.


For slide steps:

Slide image fills the screen (respecting aspect ratio).





Autosave & recovery

Deck state is automatically saved into browser storage after any change.
On app load, if a previously open deck exists, it should:

restore that deck,
and set the editor to that deck.





Import / export deck file (self-contained)

Export:

Produces a single file (JSON) containing:

Deck metadata.
List of steps.
Embedded slide image data as base64 strings.




Import:

User selects a previously exported file.
App parses JSON and loads deck into storage.






3. Non-Functional Requirements


Platform

Runs in modern browsers (Chrome/Edge/Firefox/Safari latest) on Windows and macOS.
No server-side component required; all logic is client-side.



Performance

Should handle decks with:

~50–100 steps.
Reasonable slide image sizes (e.g. 1920x1080 PNG/JPEG).





Resilience

If the tab/browser is closed or crashes, user should not lose the current deck’s state.
Import/export is the primary way to move decks between machines.



Offline

After initial load, app should work offline (PWA optional but nice).



Security

Web steps use iframes.
No elevated privileges; never execute arbitrary script beyond normal web content rendering.




4. Technology Choices (suggested)
These are suggestions; the builder can adjust, but should preserve behavior.

Framework: React (with hooks), Vue, or Svelte. (Examples assume TypeScript, but not mandatory.)
Bundler/Dev tooling: Vite or similar.
State management:

Single global store (e.g. React Context + useReducer, or Zustand/Pinia).


Persistence:

IndexedDB (via a thin abstraction or library like idb), plus localStorage for last deck ID.


Presentation

Fullscreen API (element.requestFullscreen() / document.exitFullscreen()).


File handling

Browser File API + FileReader for slide image import.
Blob + URL.createObjectURL for export (download).




5. Data Model Specification
5.1. Types (conceptual, TypeScript-style)

```
type StepType = 'web' | 'slide';

type SlideRef = {
  storedId: string;           // internal ID to locate image in IndexedDB/cache
  sourceFileName?: string;    // original filename (for display only)
};

type PresentationStep = {
  id: string;                 // unique ID (e.g. uuid)
  type: StepType;
  title?: string;
  notes?: string;             // optional, for future presenter notes
  groupId?: string;           // all steps in same “animation group” share this

  // For type = 'web':
  url?: string;

  // For type = 'slide':
  slideRef?: SlideRef;
};

type Presentation = {
  id: string;
  title: string;
  createdAt: string;          // ISO8601
  updatedAt: string;          // ISO8601
  items: PresentationStep[];
};
```

5.2. Stored image data (for self-contained package)
For internal persistence (IndexedDB):

```

type StoredSlideImage = {
  id: string;                 // matches SlideRef.storedId
  mimeType: string;           // e.g. 'image/png'
  dataBase64: string;         // base64-encoded image data
};
```

In practice, images may be stored in IndexedDB as blobs, but export format uses base64 so it’s portable.

6. Deck File Format (Import/Export)
6.1. File extension

Suggest: .presentdeck.json (or .presentdeck with JSON inside).
MIME type: application/json.

6.2. JSON structure

{
  "formatVersion": 1,
  "presentation": {
    "id": "deck-123",
    "title": "My Hybrid Talk",
    "createdAt": "2026-03-04T09:00:00.000Z",
    "updatedAt": "2026-03-04T09:15:00.000Z",
    "items": [
      {
        "id": "step-1",
        "type": "web",
        "title": "Intro blog post",
        "url": "https://example.com/intro"
      },
      {
        "id": "step-2",
        "type": "slide",
        "title": "Overview (bullets 1-2)",
        "groupId": "group-1",
        "slideRef": {
          "storedId": "img-1",
          "sourceFileName": "overview-step1.png"
        }
      },
      {
        "id": "step-3",
        "type": "slide",
        "title": "Overview (all bullets)",
        "groupId": "group-1",
        "slideRef": {
          "storedId": "img-2",
          "sourceFileName": "overview-step2.png"
        }
      }
    ]
  },
  "slides": [
    {
      "id": "img-1",
      "mimeType": "image/png",
      "dataBase64": "<base64-data-here>"
    },
    {
      "id": "img-2",
      "mimeType": "image/png",
      "dataBase64": "<base64-data-here>"
    }
  ]
}
```

6.3. Import behavior

On import:

Parse JSON.
Validate formatVersion.
Store presentation in IndexedDB under a new local internal ID (but can reuse presentation.id).
Store each slide image in IndexedDB keyed by slides[i].id.
Mark this as the “current deck” and open it in the editor.



6.4. Export behavior

On export:

Lookup current presentation and all referenced storedIds.
Retrieve associated image data from IndexedDB.
Build the JSON structure above.
Serialize and trigger a browser download.




7. Storage & Autosave
7.1. IndexedDB schema (logical)
Use a database name like present_decks_v1 with object stores:

presentations

key: presentation.id
value: Presentation object


slides

key: slideImage.id
value: StoredSlideImage



7.2. localStorage keys

currentPresentationId: ID of the deck currently open in editor.
Optionally maintain recentPresentationIds as a small list.

7.3. Autosave policy

Trigger save whenever:

Step is added / deleted / reordered.
Step data is updated (title, URL, notes).
Presentation title changes.


Implementation:

Debounce writes (e.g. 300–500 ms) to avoid excessive writes during typing.




8. UI / UX Specification
8.1. High-level layout
Root views:

Editor view (default)
Presentation view (fullscreen)

8.1.1. Editor layout


Header bar (top)

Left: Presentation title (editable text).
Right:

New Deck
Import Deck…
Export Deck…
Present (button to start presentation mode, from either:

current selected step, or
beginning – choice is optional; see below)







Main content

Sidebar (left)

Scrollable list of steps.
“Add Web Step” and “Add Slides” buttons.


Preview pane (right)

Shows preview of selected step (web iframe or image).
Displays basic info (type, URL or filename, etc).
Might include controls for editing step title/URL.






8.2. Sidebar Behavior
8.2.1. Step list items
Each item shows:

Icon:

🌐 for web
🖼 for slide


Title (step.title or fallback):

For web: fallback = hostname or URL if no title.
For slide: fallback = slideRef.sourceFileName.



Optional enhancements:

If groupId present:

Slight indentation for second+ items in the group, or a vertical bar & group label.
Or show “(1/3)”, “(2/3)” style.



Interactions:

Click = select step.
Drag-and-drop = reorder.
Context menu (right-click or ⋮ button):

Rename step.
Delete step.
“Start new group from here” or “Remove from group” (if grouping is implemented in UI).



8.2.2. Add buttons


+ Web Step

Opens a small modal or inline form:

URL input.
Optional custom title.


On “Add”, step is appended after currently selected step (or at end if none selected).



+ Slides

Opens a file picker (accept="image/*" multiple).
After selection:

Ask: “Add as separate steps or one animation group?”

Option A: “Separate steps” → each file becomes its own step.
Option B: “One group (for animations)” → all get the same groupId.


App reads images via FileReader and stores them.






8.3. Preview Pane
8.3.1. Common elements

At top:

Step type label (Web page / Slide image).
Title field (editable).


Optional “Step notes” textarea (future-proofing).

8.3.2. Web step preview

Main area: <iframe> container (e.g. takes most of the pane).
Above or below iframe:

URL input (editable).
“Open in new tab” external link icon.
“Reload” button.



Iframe settings:

sandbox attributes can be minimal (e.g. allow-scripts, allow-same-origin) depending on how cautious you want to be—but at least allow normal browsing.
Height: fill available vertical space, with overflow hidden or auto.
The preview does not need to handle keyboard navigation; that’s only for Presentation mode.

8.3.3. Slide step preview

Show image centred in a light/neutral background.
Fit-to-width or fit-to-height so it’s fully visible without scroll (unless the window is very small).
Below or above:

Show sourceFileName if present.
Show image resolution and size (optional).




9. Presentation Mode Specification
9.1. Entering presentation mode
From editor:

Clicking Present:

Option 1 (simpler): always start from first step.
Option 2 (preferred for flexibility): show small popup:

“Start from beginning”
“Start from selected step”





Implementation detail:

Root React/Vue app has a presentationMode state.
When active:

A dedicated “Presentation” component is rendered, covering the viewport.
That component calls document.documentElement.requestFullscreen().



9.2. Presentation layout

Black or dark background.
Centered content area.
No UI except:

Optional minimal controls overlay (e.g. step counter bottom-right: “3 / 24”), fade out after a few seconds.



9.2.1. Slide step in presentation

Render <img> at max size, preserving aspect ratio.
Ensure no scrollbars (content should scale).

9.2.2. Web step in presentation

Render <iframe> filling most of the screen, with:

Height and width = 100% of presentation content area.


Allow scrolling via mouse/trackpad.

9.3. Keyboard and mouse controls
Global key listener at top document level (not inside iframe):

ArrowRight or Space:

Advance to next step if exists.
If on last step, optional behavior: do nothing, or exit presentation (recommend: do nothing).


ArrowLeft:

Go back to previous step (if exists).


Escape:

Exit presentation mode.
Call document.exitFullscreen() and toggle app state back to editor view.



Technical note:

Attach keydown listener on document in a useEffect (or equivalent) when entering presentation mode.
Always preventDefault() for keys you handle to avoid browser default behavior interfering.

Click-based navigation (optional but nice):

Left half of screen (click) → previous step.
Right half of screen (click) → next step.

9.4. Focus handling
Key requirement:

Key events must be handled at the top-level app so that web iframe cannot “steal” arrow keys / space from navigation.

Implementation:

Do NOT rely on iframe focus for navigation.
Always attach key listener to window/document when presentation mode is active.


10. Error Handling & Edge Cases


Invalid URL in web step

If iframe fails to load:

Show browser default error inside iframe is fine.
No extra handling required for v1.





Missing slide image

If a slide step refers to storedId not found in IndexedDB (e.g., corrupted deck):

In editor: show placeholder “Missing slide image”.
In presentation: show friendly error slide (“Missing slide image: [title]”).


Should be rare if import/export logic is correct.



Large deck file import

If JSON is huge or parsing fails:

Show error message: “Could not import deck – invalid or too large file.”


For personal use, no need for sophisticated handling beyond this.



Autosave failures

If IndexedDB write fails, show non-blocking toast warning:

“Autosave failed. You may lose changes. Try exporting your deck.”





Fullscreen errors

If requestFullscreen() is blocked (e.g., browser doesn’t allow it):

Fallback: present in “maximized” in-page mode (no actual fullscreen) and show warning.






11. Implementation Phasing (suggested for the builder)
Phase 1 – Core editor & basic slides

Implement:

Presentation data model.
Editor view.
Add/delete/reorder steps.
Slide import (images only, no grouping).
Web step preview.
Slide preview.
Autosave to IndexedDB.
Basic import/export (deck + slides).



Phase 2 – Presentation mode

Implement fullscreen presentation view.
Keyboard navigation.
Web iframe + slide rendering.
Exit to editor view.

Phase 3 – Quality-of-life

Grouping of slide steps for animations (just groupId + UI decoration).
“Start from selected step” option.
Better error messages.
Minor UI polish (titles, counters, etc).


12. Optional Extensions (future, out of v1 scope)
Not required now, but you can keep in mind:

PDF slide import (using PDF.js, each page → slide).
“Capture screenshot from URL” to create static backups of web steps.
Presenter notes + dual screen (requires Electron/Tauri-like environment for full control).
Cloud backup or sync (requires backend).
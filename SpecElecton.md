1. Product Overview
1.1. Summary
A cross‑platform desktop app built with Electron that lets you present a sequence of:

Interactive web pages (opened as full pages in a Chromium window), and
Slide images (exported from PowerPoint or similar),

with:

An Editor window for assembling and editing decks.
A Presentation window that goes fullscreen on a chosen display.
Keyboard navigation (Left/Right/Esc) that always works, even on sites like GitHub that block iframes.
Autosave to disk and exportable deck packages for moving decks between machines.

Key design point: web pages are shown as top‑level pages, not iframes, so X‑Frame‑Options / frame-ancestors CSP do not block them.

2. Technology & Constraints
2.1. Tech stack

Runtime: Electron (main process + renderer processes)

A reasonably recent Electron version (e.g. 28+), but builder can choose.


Frontend for Editor:

React + TypeScript (recommended) OR similar SPA framework.


View layer for slides:

Simple HTML/JS (can be React too, but doesn’t have to be; see below).


Persistence:

Node filesystem APIs (via main process).
Decks stored as folders under app data directory, plus explicit export/import as .presentdeck zip/store files.



2.2. Security

Renderer windows must not have unrestricted Node.js access:

contextIsolation: true
nodeIntegration: false
Expose a limited IPC API via preload scripts (contextBridge).


Main process handles:

File system access.
Deck creation, saving, export/import.
Presentation window creation and keyboard interception.




3. High-Level Architecture
3.1. Processes & windows


Main process

Manages:

Application lifecycle.
Deck storage on disk.
Window creation & destruction.
Presentation navigation logic.
Keyboard interception in the Presentation window.


Exposes a set of IPC channels / APIs to renderer processes.



Editor renderer

Runs in the Editor window.
SPA that provides:

Deck list (optional v1), or simply opens the last deck.
Sidebar of steps.
Preview pane.
Buttons for Create/Open/Save/Export/Import and Play.


Communicates with main process via IPC for:

Loading/saving decks.
Selecting files to import slides.
Starting/stopping presentations.





Presentation viewer

Implemented as a Presentation BrowserWindow created by main.
Shows one thing at a time:

For slide steps: a local HTML file (presentation-slide.html) with our own viewer UI.
For web steps: navigates directly to the external URL.


Main process listens to before-input-event on this window to intercept Left/Right/Esc and perform navigation.




4. Data Model
4.1. In-memory types (conceptual)
Use TypeScript-style types to guide implementation:
```
type StepType = 'web' | 'slide';

type SlideRef = {
  id: string;               // internal slide ID (e.g. "slide-0001")
  relativePath: string;     // relative path within deck folder, e.g. "slides/slide-0001.png"
  sourceFileName?: string;  // original filename for user reference
};

type PresentationStep = {
  id: string;               // unique ID (e.g. uuid)
  type: StepType;
  title?: string;
  notes?: string;           // reserved for presenter notes (optional v1)
  groupId?: string;         // used to group multiple slide images into an "animation" sequence

  // Web step
  url?: string;

  // Slide step
  slideRef?: SlideRef;
};

type Presentation = {
  id: string;               // deck ID (e.g. uuid or folder name)
  title: string;
  createdAt: string;        // ISO string
  updatedAt: string;        // ISO string
  items: PresentationStep[];
};
```
4.2. On-disk deck structure
Each deck is stored as a folder:
```
<decks-root>/<deckId>/
  deck.json
  slides/
    slide-0001.png
    slide-0002.png
    ...
```

<decks-root>: e.g. app.getPath('userData') + '/decks' by default.
deck.json: serialized Presentation object (with SlideRef.relativePath values).

4.3. Exportable “all-in-one” deck file
For moving decks between machines:

Export as a single .presentdeck file:

Option A: a zip archive containing the deck folder (deck.json + slides/).
Option B: a JSON bundle with base64-encoded images (slightly more custom; zip is more standard).



Recommended: zip because it is simple and tools exist.
Internal zip structure is the same as the folder:

```
deck.json
slides/slide-0001.png
slides/slide-0002.png
...

```

5. Main Process Responsibilities
5.1. Deck management
Main process will provide IPC handlers for:

deck:create → creates a new deck with:

new id
default title ("Untitled deck" or similar)
empty items
saves to disk (<decks-root>/<deckId>/deck.json).


deck:load → load deck.json from given deckId or folder path.
deck:save → persist updated Presentation object to deck.json.
deck:export → zip the deck folder into .presentdeck and trigger save dialog.
deck:import → open file dialog, unzip .presentdeck into a new deck folder, return new deckId.

5.2. Slide import
IPC: deck:importSlides

Renderer sends:

deckId
array of selected file paths (images).
an option to treat them as:

mode: 'separate' – each file → separate step.
mode: 'grouped' – all files share the same groupId (for animation sequence).





Main process:

For each selected file:

Copy the file into slides/ subfolder under the deck folder.
Generate a canonical name (slide-0001.ext, slide-0002.ext, etc.).


Return to renderer:

An array of SlideRef objects with id and relativePath.


Renderer updates its Presentation.items with new steps and then calls deck:save.

5.3. Presentation window management
Main process handles:


presentation:start (from Editor):

Params:

deckId
startIndex (index of step to start from).
optionally displayId (monitor choice; v1 can default to primary).


Loads deck.json into memory.
Creates a new BrowserWindow:

Fullscreen on chosen display.
Black background.


Stores:

currentDeck
currentIndex
presentationWindow reference.





presentation:stop:

Closes the presentation window if open.
Clears presentation state in main.




6. Presentation Logic (Main Process)
6.1. Step navigation
Playback state:

```
let currentDeck: Presentation | null;
let currentIndex: number | null;
let presentationWindow: BrowserWindow | null;
```

Helper: showStep(index: number)

Bounds check (0 <= index < deck.items.length).
Get step = deck.items[index].
Branch:

step.type === 'slide' → show slide viewer.
step.type === 'web' → navigate to external URL.



6.2. Rendering slide steps
For slide steps:

presentationWindow.loadFile('presentation-slide.html', { query: { deckId, slidePath: step.slideRef.relativePath } });

Or:

presentationWindow.loadURL('file://<app>/presentation-slide.html?deckId=...&slide=slides/slide-0001.png')

The presentation-slide.html is a static HTML/JS file (with optional React) that:

Reads slidePath from query string.
Resolves the actual absolute path:

main process can also inject an IPC event with full path if easier.


Displays the image fullscreen with black background.

You can pass the absolute image path via file:// URL, but be careful with security; better to provide a safe API via preload script to look up the image path.
6.3. Rendering web steps
For web steps:

presentationWindow.loadURL(step.url).

This navigates the BrowserWindow to the external site as a top-level page.
Note:

At this point, our custom HTML is not loaded; the site is hosted by the external origin.
We still maintain control via main process events (see keyboard handling).


7. Keyboard Handling in Presentation Window
Goal: Left/Right/Esc always control your deck, even on arbitrary websites.
Implementation:

In main process, after creating presentationWindow, attach:


```
presentationWindow.webContents.on('before-input-event', (event, input) => {
  if (input.type !== 'keyDown') return;

  // Navigation keys
  if (input.key === 'ArrowRight' || input.key === ' ' /* Space */) {
    event.preventDefault();
    goToNextStep();
  } else if (input.key === 'ArrowLeft') {
    event.preventDefault();
    goToPreviousStep();
  } else if (input.key === 'Escape') {
    event.preventDefault();
    stopPresentation();
  }
});
```

goToNextStep():

if currentIndex < deck.items.length - 1, increment and call showStep(currentIndex).


goToPreviousStep():

if currentIndex > 0, decrement and call showStep(currentIndex).



This ensures navigation is independent of:

whether the page is a slide viewer or a GitHub page, or anything else.

Scrolling in web pages:

Do not intercept scroll wheel events.
If you want to allow PageUp/PageDown or Home/End for navigation, that’s optional – but default scroll wheel / trackpad should “just work”.


8. Editor Window UI / UX
8.1. Layout


Top bar

Editable deck title.
Buttons:

New Deck
Open Deck… (optional if you list recent decks)
Import Deck…
Export Deck…
Play (dropdown: “From beginning” / “From selected step”).





Main area

Left: Sidebar – ordered list of steps.
Right: Preview pane – shows simple preview/editor for selected step.



8.2. Sidebar behavior
Each step item shows:

Icon:

🌐 for web
🖼 for slide


Title:

step.title if present.
For web: fallback = URL hostname.
For slide: fallback = slideRef.sourceFileName or slideRef.id.


Optional grouping cue:

For items with same groupId, display them slightly indented or with a bracket marker.



Interactions:

Click: select step.
Drag & drop: reorder steps. (After reorder, renderer sends updated Presentation to main via deck:save.)
Right-click / kebab menu:

Rename step.
Delete step.
(Future) “Start new group here / remove from group”.



8.3. Step creation controls
Below sidebar or in top bar:


Add Web Step

Opens prompt modal:

URL input.
Optional title input.


On confirm:

Create new step:

type: 'web'
url as entered.
title as provided.


Insert after selected step (or end of list if none).
Trigger save.





Add Slides

Opens native file picker:

Filter *.png;*.jpg;*.jpeg.
multiSelections enabled.


After selection:

Show modal:

“Import as:”:

◉ Separate steps
○ One animation group




Renderer sends deck:importSlides with:

file paths, deckId, mode.


Main copies files into deck folder and returns list of SlideRefs.
Renderer creates new steps and saves.





8.4. Preview pane
Common fields:

Title input (editable).
Notes textarea (optional, can be left out in v1 but good to reserve field).

Web step preview:

Show:

URL input (editable).
Button: Open in external browser (calls shell.openExternal(url) via IPC).


For the preview itself:

Optional:

Use a small <webview> or similar to show a live preview.
However: this may hit the same X‑Frame‑Options/CSP issues as iframes; that’s OK for preview. If it fails, show a friendly message (“This site cannot be embedded; will open fully in presentation mode.”).


V1 can even skip live preview and just display the URL and a note.



Slide step preview:

Load the slide image via a custom file URL or via a safe IPC-driven path resolver.
Show image scaled to fit preview pane.
Display sourceFileName and resolution (optional).


9. Deck Save & Autosave
9.1. Autosave strategy

Renderer maintains the current Presentation state.
On any change (add/remove/reorder/update), it:

Updates local state.
Sends deck:save with the full Presentation object.



Main process:

On deck:save:

Overwrite deck.json with new JSON content.
Update updatedAt timestamp before writing.
Also update a lastOpenedDeckId in a global settings file (optional).



9.2. App startup
On app start:


Main process:

Reads lastOpenedDeckId from a small settings JSON (e.g. <userData>/config.json).
If found and exists:

Loads that deck.


Otherwise:

Creates a new empty deck.





Editor renderer:

Receives initial deck payload via IPC on load.




10. Multi-Monitor & Fullscreen Behavior
10.1. Display selection (nice-to-have for v1, but can be simple)
Use Electron’s screen module:

For v1, simplest approach:

Present on the primary display.



If you want to support choosing:

Expose displays list via IPC:

system:getDisplays → returns array of displays with id, size, etc.


In settings / Play menu, user can choose target display.
When starting presentation:

Create BrowserWindow with x, y, width, height set to the target display’s bounds.
Set fullscreen: true on that display.



10.2. Exit behavior

Esc:

main process closes presentation window.
Focus returns to editor window.




11. Error Handling & Edge Cases


Web pages requiring login

Not a bug: the presentation window is a separate browser.
Suggest adding a “Rehearse” button (optional):

Opens presentation mode and quickly auto-advances through web steps so user can log in ahead of time.





Slide image file missing

If a slide path doesn’t exist (corrupted deck):

Editor preview: show placeholder + “Missing slide image at [relativePath]”.
Presentation: show a black screen with centered error text, but don’t crash navigation.





Export/import failures

If zipping/unzipping fails:

Show error dialog via dialog.showErrorBox.


If imported deck.json is invalid:

Show “Invalid deck file” message; do not crash app.





Presentation window creation failure

If for some reason the window can’t be created (rare):

Show error dialog in editor window, log details.





Keyboard collision

Some web apps may rely heavily on Arrow keys.
We explicitly override Left/Right/Esc for deck navigation; this is intentional and should be documented in code comments.




12. IPC Contract (Suggested Channels)
For clarity, here’s a suggested IPC surface. Builder can adjust names, but keep semantics:
From Renderer (Editor) → Main:


deck:create → Promise<Presentation>


deck:load (deckId: string) → Promise<Presentation>


deck:save (presentation: Presentation) → Promise<void>


deck:export (deckId: string) → Promise<void> // shows save dialog


deck:import () → Promise<Presentation>      // shows open dialog & returns imported deck


deck:importSlides (params: { deckId: string; filePaths: string[]; mode: 'separate' | 'grouped' })
→ Promise<SlideRef[]>


presentation:start (params: { deckId: string; startIndex: number; displayId?: number }) → Promise<void>


presentation:stop () → Promise<void>


From Main → Renderer (Editor):

deck:changed (presentation: Presentation)

Optional: if main ever modifies deck (e.g. after import).



For Presentation window, no rich IPC is strictly needed if main handles all key events and navigation directly; slides are shown by loading a static presentation-slide.html with query parameters.

13. Out-of-Scope / Future Enhancements
Not required for v1, but easy to extend later:

Presenter notes and “presenter view” on laptop screen while audience sees fullscreen.
PDF slide import (convert pages to images or render via HTML).
Direct .pptx parsing (likely overkill).
Automatic screenshot capture from web steps as a fallback if they fail live.
Cloud sync or sharing.
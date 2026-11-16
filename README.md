# 2X Deck Fixer

Drop 2X Interim decks (.pptx) or templates (.potx) into the page and save back "(fixed)" copies where every box can be clicked, moved, resized and deleted in PowerPoint. Slides look exactly the same. Files are fixed inside the browser; nothing is uploaded and the original is never changed.

## Use it
- **Open `index.html`** in any browser, or host the folder on GitHub Pages. Fixed decks save as normal downloads.
- **In claude.ai**, the same page runs as an artifact and saves through Claude's save prompt.

Steps: close the deck in PowerPoint, drop it in, save the "(fixed)" copy, check a slide or two, then use it in place of the original. For SharePoint or OneDrive, upload it over the original; version history keeps the old one.

## What it fixes
| Problem | Fix |
|---|---|
| Short boxes (eyebrow, source, labels, dates) show no selection frame on Mac | Layout boxes grow to a frame-showing height; the extra height goes into the text insets, so text doesn't move. Short boxes on existing slides become plain text boxes. |
| Layout shapes (card backgrounds, quote marks, logo) can't be selected on slides | Each slide gets its own copy of its layout shapes |
| Panels (card, step, stat…) can't be resized or deleted on new slides | Each panel becomes the shape of its body text box |
| Panel numbers can't be deleted | Card, step, list and rung numbers become text boxes; blank numbers are filled so they show when presenting |
| Small boxes hidden under bigger ones; titles covering eyebrows; photos under overlays | Restacked, titles shortened above the eyebrow, overlays ignore clicks |

Agenda numbers and logo tiles stay fixed as part of the design. Decks not built on the 2X Interim template are left alone. Running it twice changes nothing.

## Project
- `index.html` — the app, self-contained (loads JSZip 3.10.1 from cdnjs and Inter from Google Fonts)
- `src/fixer.js` — the fixing engine; works in a browser or Node
- `src/page_head.html`, `src/page_body.html` — the page; `python3 build.py` rebuilds `index.html` after edits
- `test/run.js` — fix a deck from the command line: `npm install`, then `node test/run.js in.pptx out.pptx`

`src/fixer.js` is a JavaScript port of `fix_selection.py` from the 2X Deck Kit (`engine/fix_selection.py`). Both produce identical files; keep them in step when changing either.

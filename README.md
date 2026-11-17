# 2X Deck Fixer

Drop 2X Interim decks (.pptx) or templates (.potx) into the page and save back "(fixed)" copies where every box can be clicked, moved, resized and deleted in PowerPoint, and every header sits where the template says. Files are fixed inside the browser; nothing is uploaded and the original is never changed.

## Use it
- **Open `index.html`** in any browser, or host the folder on GitHub Pages. Fixed decks save as normal downloads.
- **In claude.ai**, the same page runs as an artifact and saves through Claude's save prompt.

Steps: close the deck in PowerPoint, drop it in, save the "(fixed)" copy, check a slide or two, then use it in place of the original. For SharePoint or OneDrive, upload it over the original; version history keeps the old one.

## What it fixes
| Problem | Fix |
|---|---|
| Short boxes (eyebrow, source, labels, dates) show no selection frame on Mac | Layout boxes grow to a frame-showing height; the extra height goes into the text insets, so text doesn't move. Short boxes on existing slides become plain text boxes. |
| Layout shapes (card backgrounds, quote marks) can't be selected on slides | Each slide gets its own copy of its layout shapes |
| Panels (card, step, stat…) can't be resized or deleted on new slides | Each panel becomes the shape of its body text box |
| Panel numbers can't be deleted | Card, step, list and rung numbers become text boxes; blank numbers are filled so they show when presenting |
| Small boxes hidden under bigger ones; titles covering eyebrows; photos under overlays | Restacked, titles shortened above the eyebrow, overlays ignore clicks |
| Brand art moved or deleted by accident | The footer line ("Growth. Orchestrated." \| 2X logo) and the grey baton behind the title are fixed on every slide. Page numbers stay editable. |
| Titles sit at different heights; cover eyebrow too far from the title; agenda numbers off their red stroke | Every title starts 0.608 in from the top, as on the reference slide; cover eyebrows sit the same distance above the cover title; agenda numbers sit centred over their stroke on the item's text line. Applied to layouts too, so New Slide matches. Title boxes stay 0.62 in or taller so Mac still draws a frame. Titles that would run into the content below are left where they are. |

**Clean up as well** (checkbox, off by default) also removes pasted footers, a second footer logo, footer pieces hidden behind a full-screen video and empty text boxes, and lists each removal per slide. Slides with no content of their own are left alone.

Agenda numbers and logo tiles stay fixed as part of the design. Slides built freehand (no 2X layout) are left as they are. Decks not built on the 2X Interim template are left alone. Running it twice changes nothing.

## Not in the browser (use the 2X Deck Kit)
- **Restore the slide master** when a deck's masters have collapsed (for example after a round trip through Google Slides, or after pasting slides with Keep Source Formatting): `engine/restore_template.py deck.pptx [template.potx] [out.pptx]`. Rebuilds the nine master families, scales a 20 × 11.25 in deck to 13.33 × 7.5 in, keeps slides, notes, comments and video, then runs the fix.
- **Embed Inter**: `engine/embed_fonts.py` (needs the font files, so it stays in the kit).

## Project
- `index.html` — the app, self-contained (loads JSZip 3.10.1 from cdnjs and Inter from Google Fonts)
- `src/fixer.js` — the fixing engine; works in a browser or Node
- `src/page_head.html`, `src/page_body.html` — the page; `python3 build.py` rebuilds `index.html` after edits
- `test/run.js` — fix a deck from the command line: `npm install`, then `node test/run.js [--clean] in.pptx out.pptx`

`src/fixer.js` is a JavaScript port of the 2X Deck Kit's `engine/fix_selection.py`, `engine/tidy.py` and `engine/clean_deck.py`. Both produce identical files; keep them in step when changing either (see `FEEDBACK.md`).

# 2X Deck Fixer — feedback and backlog for Claude Code

Read this first, then `README.md`. It records what the fixer does today, what testing found, and the work still open. Start with **Open work** (P1 first). Keep **Rules that must hold** true in every change.

---

## 1. Context

- **Problem the fixer solves:** in PowerPoint (especially for Mac), boxes on 2X Interim template slides could not be selected, showed no selection frame, could not be moved or deleted, and panels and numbers on slides added with **New Slide** were locked.
- **Two implementations that must stay identical:**
  - `src/fixer.js`: the browser app (`index.html`), also runs in Node via `test/run.js`.
  - `engine/fix_selection.py` + `engine/tidy.py` + `engine/clean_deck.py` in the 2X Deck Kit (`~/Desktop/PPT/2X_Deck_Kit_v1/engine/`, also in the `2x-deck-builder` skill's `scripts/`): the kit's **Fix Selection** launchers and the deck builder (`build_deck.py`, which calls `fix_all(out, add_numbers=True)`).
- **Proven parity:** on 11 real decks, `fixer.js` (Node and Chromium) and `fix_selection.py` produce identical slide, layout and master XML and identical relationship IDs. Every fixed slide renders the same as the original (LibreOffice render diff; only sub-pixel text shifts). A second run changes nothing.

## 2. What the fixer does (in this order)

| # | Step | Function (JS / Python) | Notes |
|---|---|---|---|
| 1 | Copy layout and master artwork onto each existing slide; set `showMasterSp="0"` | `detach` | Copies tagged `descr="2X layout artwork"`; image rels remapped. **Must run before 2 and 3 on existing decks.** |
| 2 | Panels (`*_Bg` autoshapes) become the shape of the largest text placeholder inside them (Body/Text/Bio/Note > Value > largest) | `panelize` | Insets grow so text stays put; roundRect text-rect inset (`adj/100000 × min(w,h) × 0.29289`) is subtracted. Slides with copied art get the placeholder pinned (own xfrm, no fill). |
| 3 | Panel numbers (`Card|Step|Item|Rung N_Number`, plus `_Badge`) become body placeholders with `hasCustomPrompt="1"`; prompt text = the number | `numberize`, `fill_numbers` | Agenda layouts skipped. Tag `descr="2X number"`. Blank number placeholders on slides are filled. |
| 4 | Titles anchored bottom that overlap an eyebrow are shortened above it | `untangle_titles` | Sets all four geometry values (python-pptx writes x=0,cx=0 otherwise). |
| 5 | Small text boxes covered ≥50% by a bigger shape are lifted above it | `fix_tree` | Only boxes that are placeholders or hold text (`_wordy`) are lifted. |
| 6 | Photo overlays ignore clicks | `unclick_overlays` | `spLocks noSelect="1"` on shapes named `Photo overlay`. |
| 7 | Short placeholders on slides (< 0.55 in, with text, not title/picture) become plain text boxes with the effective style baked in | `to_textboxes` | Tag `descr="2X text box"`. |
| 8 | Short layout placeholders grow to 0.62 in; extra height goes into tIns/bIns | `grow_short`, `follow_layout` | Filled/outlined placeholders (`_painted`) are not grown. Never grows past the top of the page. Moves to top of z-order inserted **before** `p:extLst`. |
| 9 | Plain artwork that covers a text box on a layout is sent behind it | `text_over_art` | Layout order only; makes reruns stable. |
| 4a | **Header tidy** (before the box fixes, and again at the end): titles, cover eyebrows, agenda | `tidy` (`tidy.py` / `tidy*` in JS) | See "Header tidy" below. |
| 6a | Brand art fixed on slides: `Tagline`, `Footer divider`, `2X logo`, `Baton watermark` copies get `noSelect/noMove/noResize/noRot/noGrp` (+`noTextEdit` / `noChangeAspect`) | `lock_footer` / `lockFooter` | Only copies tagged `2X layout artwork`. Page number stays editable. |
| 10 | Duplicate or non-numeric `cNvPr` ids renumbered | `unique_ids` | Artwork copies renumbered first. |
| 11 | *Opt-in* clean-up: pasted footers, second footer logo, art hidden behind a full-screen video, empty text placeholders | `clean_deck.clean` / `cleanUp` (`opts.clean`) | Skips slides with no content of their own. Returns a per-slide log. |

**Header tidy** (reference: slide 2 of the Town Hall, approved 2026-11-17):
- *Titles:* first line top at **0.608 in** (`TOP = 555955`), one 32 pt line = `LINE = 491491`. The box is written as top `TOP − PAD`, height `n·L + 2·PAD`, `tIns = bIns = PAD = 37718`, anchor `b` — so the box is ≥ 0.62 in (Mac draws a frame) and the text sits exactly at `TOP`. Line count `n` from an Inter Bold advance-width table (UPM 2048, no kerning, identical in both languages). Slides: only slides carrying `Baton watermark` (our content slides). If `TOP + n·L` would come within 0.1 in of the first shape below, the title is pinned where it is. Layouts: titles whose first line is within 0.25 in of `TOP`, always one line.
- *Cover eyebrows:* eyebrow text top = title first line − **0.29 in** (`GAP_EB = 265176`). On layouts the empty space above the title's lines is trimmed and the eyebrow box ends where the title box starts (insets carry the offset); on slides only the eyebrow moves.
- *Agenda:* each number sits centred over its red stroke; the column's centre line is `item.left − 0.20 in − widest number / 2`; number and item text share one centre line; stroke centre 0.30 in under it. "Agenda" heading placed like a title.
- Every check compares **text positions**, not boxes, so later box fixes (untangle, grow) can't make it flip between runs.

**Why the thresholds:** tested in PowerPoint for Mac. No selection frame is drawn on placeholders shorter than about 0.5 in; 0.62 in shows one; plain text boxes always show one.

## 3. Rules that must hold

1. **Parity:** any change to `src/fixer.js` is mirrored in `engine/fix_selection.py` (and the reverse). Run the parity test (section 6) before committing.
2. **No visual change:** existing slides must render the same as the original. Re-run the render diff on the test decks.
3. **Idempotent:** fixing a fixed deck changes nothing (`report` counts all 0 → UI shows "Already fixed"). Known exception: one raw deck (`d11`, slide 18) needs two runs before a page number settles in z-order — pre-existing, not caused by the tidy.
4. **Order:** for existing decks, `detach` runs **before** `panelize`/`numberize`. Reversing it deletes badges and stat panels from existing slides (bug found and fixed on 2026-11-17). Only the deck builder (`add_numbers=True`) runs `detach` after.
5. **Only Interim decks:** the app refuses files without Interim layouts (`Cards_Numbered*`, `IconRow*`, `Cover_Hero*`, `Divider_Section*`).
6. **Original is never modified;** output is `<name> (fixed).<ext>`.
7. **No uploads:** everything runs in the browser. JSZip 3.10.1 from cdnjs is the only script dependency.

## 4. Field test: colleague's Town Hall deck (2026-11-17)

File: `~/Downloads/2026-11-19 2X Town Hall Deck - Q4 v2.pptx` (48 slides, 1 hidden) and its `(fixed)` copy.

**Result: fix succeeded.** 254 layout shapes made clickable, 21 short boxes converted, 118 layout panels and 41 numbers made editable. 47/47 shown slides render the same (slide 40: sub-pixel only). No locked, hidden, tiny or off-slide shapes remain.

**Findings the fixer does not handle yet** (become backlog items below):

| Slides | Finding | Impact |
|---|---|---|
| 7 | Full-bleed video covers the copied footer logo, tagline and page number | Invisible junk shapes under the video; can't be clicked |
| 24–36 | Pasted footer "GROWTH. ORCHESTRATED. N / 36" from another 36-page deck, on top of the template page number and tagline (slide 24 shows "24" and "14 / 36") | Visible duplicate footers |
| 19–21 | Second 2X logo in the footer band, overlapping the template logo | Visible double logo |
| 14 | `PhotoCard_3col` slide with all 18 placeholders empty | Presents as a blank white slide |
| 19–21 | "DRAFT PREVIEW" label still on slide | Easy to forget before presenting |
| 21 slides (e.g. 5, 9, 11) | Empty text placeholders left over | "Click to add text" clutter while editing |
| deck | Playfair Display used; Inter not embedded | Off-brand font; substitution on machines without Inter |
| 42, 43, 45 | Title Case while the rest are sentence case | Inconsistent titles |
| deck | Page size **20 × 11.25 in** (template scaled 1.5×); kit template is 13.33 × 7.5 in | Fixed thresholds (0.55 / 0.62 in) are absolute; see P1-3 |
| 30, 34 | Footer logo under a clear rectangle | Not blocking (clicks pass through) |
| 17 | Two ripple shapes grouped | Not blocking (double-click to edit) |

## 5. Open work

### Done 2026-11-17
- Brand art fixed (footer line + grey baton), both languages.
- Header tidy (titles, cover eyebrows, agenda), both languages; parity on 11 test decks; idempotent.
- Clean-up mode (P2-1…P2-3, plus P1-1 as part of clean-up) in both languages; app checkbox + per-slide log.
- Restore (P1-4): comments now kept (authors part related, slide ids preserved, modern comment pins scaled); placeholders with no inherited fill/outline get explicit `noFill`/`ln noFill` so the new layouts' outlines don't show through; bundled template used when none is given.

### P1 — correctness

**P1-1. Don't copy artwork that a full-bleed picture or video hides.** *(Handled by clean-up when ticked; open for the default fix.)*
In `detach`, skip copying a layout/master shape when the slide has an opaque picture, video or picture placeholder covering ≥95% of that shape's box *and* drawn above where the copy would go. Today the copy lands under the media, invisible and unclickable (slide 7).
*Accept:* slide 7 of the Town Hall gets no footer copies; render diff unchanged; parity holds.

**P1-2. Verify Delete on new-slide panels and numbers in PowerPoint.**
Automated tests couldn't press Delete (remote keystrokes don't reach PowerPoint's canvas). Manually: open a fixed template, New Slide → `IconRow_3col` and `List_Numbered_5`, select an empty panel, press Delete; same for a number badge.
*Accept:* both delete cleanly; record the result here.

**P1-3. Scale thresholds with page size.**
`SHORT` (0.55 in) and `TALL` (0.62 in) assume 13.33 in wide. On a 20 in-wide deck, boxes are 1.5× taller, so some short boxes now pass the threshold and keep no frame, or vice versa. Decide by testing in PowerPoint for Mac at the deck's actual size whether the frame threshold follows on-screen size (zoom-dependent) or slide units. If it follows slide units, keep as is; otherwise scale by `slide_width / 12192000`.
*Accept:* eyebrow and source boxes show frames on the 20 × 11.25 in Town Hall after fixing.

### P1-4. Restore template structure (new, prototype done in Python)

**Problem:** decks that went through another app or had slides pasted in end up with collapsed masters. The colleague's Town Hall: 2 masters instead of 9, 63 layouts crammed into the Covers master, the whole People & Proof family missing, 4 stray layouts (`DEFAULT`, `Title and body`, `1_Stat_Row_3`, `1_Divider_Statement`), page size 20 × 11.25 in (1.5× the template), built on an older template version (layout boxes 0.2–0.25 in off today's).

**Prototype:** `engine/restore_template.py` in the kit (`python restore_template.py deck.pptx 2X_Template_Interim_v1.potx [out.pptx]`). Verified on the Town Hall:
- result has 9 masters, 86 layouts, 13.33 × 7.5 in, 48 slides, slide 36 still hidden, video and all speaker-note text kept
- all 47 shown slides match the original visually (only anti-aliasing differences from the 20 in → 13.33 in render)
- package valid (no missing content types, broken links, duplicate part names); opens in PowerPoint for Mac with no repair prompt; Slide Master view shows the original family order

**How it works (keep this order):**
1. **Freeze** each slide: copy layout artwork onto it (`detach`), give every placeholder its inherited position, shape and effective text style (`_effective_text`), and give free-drawn text the deck's default text style. Without this, slides change because they inherit from different (older or newer) layouts.
2. **Scale** every page-relative value by new/old height (both must be 16:9 within 0.1%): positions, sizes, group interiors, text sizes, letter spacing, line weights, insets, indents, tabs, shadows, table columns/rows, chart and diagram text. Text sizes stay exact (1/100 pt).
3. **Move** each slide onto the clean template's same-named layout (`1_X` → `X`, unknown → `Blank`), renaming foreign parts so names are unique; copy notes text.
4. Run the selection fix on the result.

**Open items:**
- ~~Comments are not carried over~~ — done: `ppt/authors.xml` is related, slide ids are kept, pins scaled. Legacy comments are carried but their pins are not scaled.
- Speaker notes keep text only, not formatting.
- Port to `src/fixer.js` as an opt-in "Restore template structure" checkbox; the app must carry the clean template (embed the .potx, ~1.1 MB, as a base64 asset).

*Accept:* Town Hall restore matches the checks above in both Python and the app, and Python/JS outputs are identical.

### P2 — clean-up mode (opt-in, off by default) — **done** (`clean_deck.py`, `cleanUp`)

Add a "Clean up as well" checkbox in the app (`--clean` in Python). Each item is reported per slide; nothing here runs unless chosen.

**P2-1. Remove pasted footers.** On a slide, delete text shapes in the bottom 12% band that duplicate the template footer: text matching `/\b\d+\s*\/\s*\d+\b/` beside "Growth. Orchestrated.", or a second "Growth. Orchestrated." Keep the template's own (tagged `2X layout artwork`).
*Accept:* Town Hall slides 24–36 keep one footer each.

**P2-2. Remove duplicate footer logos.** Delete a picture in the footer band when the slide already has the tagged template logo there.
*Accept:* slides 19–21 show one logo.

**P2-3. Delete empty text placeholders on slides,** except number placeholders (filled by step 3) and picture placeholders.
*Accept:* no "Click to add text" left on slides 5, 9, 11; no visual change in slideshow render.

### P3 — deck health report (read-only)

Show a "Needs a look" list under each fixed file, without changing anything:
- slides whose placeholders are all empty (slide 14)
- text containing `DRAFT`, `[INSERT`, `TBD`
- fonts other than Inter (and whether Inter is embedded)
- mixed title case (reuse the rule in the kit's `verify_deck.py`)
- page size different from 13.33 × 7.5 in
- hidden slides

*Accept:* the Town Hall shows slides 14, 19–21, 42/43/45, Playfair Display, 20 × 11.25 in and hidden slide 36.

### P4 — nice to have

- Web Worker for very large decks (the 130 MB Town Hall takes ~15 s and ~415 MB heap on the main thread).
- Progress text per step for large files.
- Font embedding is **out of scope** for the browser (needs font files and licensing); keep it in the kit's `embed_fonts.py`.

## 6. How to test

Test decks (not in the repo; keep local): the 11 Interim decks in `~/Desktop/PPT` (samples, the six "(2X rebrand)" decks, both Town Halls) plus the colleague Town Hall in `~/Downloads`. Originals are backed up in `~/Desktop/PPT/Archive/_backup_before_selection_fix_v2_2026-11-16/`.

1. **Parity (must be identical):**
   ```bash
   npm install
   node test/run.js in.pptx js.pptx
   cp in.pptx py.pptx && python3 ../2X_Deck_Kit_v1/engine/fix_selection.py py.pptx
   ```
   Compare every `ppt/**/*.xml` part with lxml C14N (strip whitespace-only text), and `.rels` as sorted `(Id, Type, Target, TargetMode)`. Ignore `[Content_Types].xml` ordering.
2. **No visual change:** render original and fixed with LibreOffice (`soffice --headless --convert-to pdf`, `pdftoppm -r 30`), diff each page with PIL `ImageChops.difference`; flag pages with > 150 pixels above 40. Known sub-pixel-only pages are fine.
3. **Idempotent:** run the fixed file through again; zero changed parts.
4. **Browser:** load `index.html` in Playwright Chromium (route cdnjs to the local `jszip.min.js`), drop files via `#pick`, check rows and the download.
5. **New-slide behaviour (manual, PowerPoint for Mac):** New Slide from the fixed template → eyebrow shows a frame; panels select, resize, delete; numbers show as "01" hints and fill after a rerun.

## 7. Known limits (by design)

- Agenda numbers, `Logos_Grid` tiles and quote marks stay as fixed layout artwork.
- Numbers on slides added with New Slide are hints until typed or until the fixer runs again; hints don't show in slideshow, PDF or print.
- An empty panel on a new slide doesn't show in slideshow (PowerPoint hides empty placeholders).
- The app hands back a fixed `.potx` inside a `.zip` (claude.ai's save allowlist has no `.potx`). Standalone, it could save `.potx` directly: in `saveOne()`, use `browserSave` with the raw template when `window.claude` is absent.

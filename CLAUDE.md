# Индивидуальный план

PWA for editing the text of .docx documents on iPhone; the owner builds it for their wife. UI language is Russian; her documents are in Ukrainian.

## Where things are

- Source: branch **`source`**. Built site: branch **`main`** (GitHub Pages). Publish with `npm run deploy` (builds and pushes `dist/` to `main`). Never hand-edit `main`.
- Standing instruction from the owner: commit and push app changes to `source`, then deploy.
- The owner's documents live only on the **`data`** branch of this (public) repository; the owner accepted that it is public. Never put them on `source` or `main`.

## Checks before pushing

`npm run build` must pass. Then try the change in a phone-sized window (390×844) with a real .docx: open, edit, reload (edits persist), export Word and PDF. Check the exported .docx opens (LibreOffice: `soffice --headless --convert-to pdf`) and its text matches.

## Rules of the code

- Fidelity first: the app only rewrites text inside existing `w:r` runs (`WordDoc.sync`). Never regenerate the document; everything else in the zip is kept byte for byte.
- Runs holding anything other than text are locked (`data-l`) and must survive edits untouched.
- Stored data: IndexedDB `individualnyi-plan`/`docs`, `DocRecord` in `src/store.ts`. Documents already on the phone must keep opening: only add optional fields.
- `src/docx/css.ts` is shared by the editor and the PDF renderer (which draws pages outside the app stylesheet).
- Fill-in lines (`src/fill.ts`): runs of "____" or underlined spaces are marked `data-blank`. Typing over one keeps the paragraph's shape (same lines, same end): the line gives up or gets back characters next to the caret. Text typed onto "____" goes into an underlined run (`data-like`/`data-u`, and `data-font` for the label's typeface on a bare line).
- `toDocx()` puts zero-width spaces between the underscores of long lines (10+) so Word wraps them like the editor does; `open()` takes them out again. Keep both sides in step.
- Undo/redo (editor buttons): per-paragraph XML snapshots (`WordDoc.snapshot` / `restore`), one step per burst of typing.
- The editor page is laid out at full size and scaled with a transform. Do not use CSS `zoom`: on iPhone it enlarged small text and dropped thin table borders.

## Sync and the `data` branch

- Documents sync with the `data` branch of this repository (`src/sync.ts`, GitHub contents API). Layout: `docs/<id>.docx` (the whole document as it is now) and `docs/<id>.json` (`{ "name", "created", "updated" }`, `updated` in ms).
- Without a key (the default, nothing to set up) the app only downloads; edits on the phone stay there. With a key (fine-grained token with Contents read/write, entered in the app, kept in localStorage `ip.sync`) it also uploads edits and deletions.
- The app pulls on start, when it comes to the foreground and every minute (conditional requests, so unchanged answers don't use GitHub's rate limit). A document changed in both places: the repository wins, the phone's version becomes a new document "… (с телефона)".
- To change a document for the owner: edit `docs/<id>.docx` on `data` in place (see below), bump `updated` in its `.json`, commit, push. To add one: pick a new id (`[a-z0-9]+`), add both files. It shows up on the phone within a minute of the app being open.

## Filling a document from the owner's text

The owner sends instructions in chat (and a .docx, or names a document already in the data repo). Fill it with a script that edits `word/document.xml` (unzip → change `w:t` text in place, keeping run formatting → zip, every other zip entry unchanged), check it logically, ask about anything unclear, leave fields without data empty, render it to check (LibreOffice), then put it on the `data` branch as above.

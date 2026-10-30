# Индивидуальный план

PWA for editing the text of .docx documents on iPhone; the owner builds it for their wife. UI language is Russian; her documents are in Ukrainian.

## Where things are

- Source: branch **`source`**. Built site: branch **`main`** (GitHub Pages). Publish with `npm run deploy` (builds and pushes `dist/` to `main`). Never hand-edit `main`.
- Standing instruction from the owner: commit and push app changes to `source`, then deploy.
- Never commit the owner's documents (they contain children's and teachers' names) to this public repo.

## Checks before pushing

`npm run build` must pass. Then try the change in a phone-sized window (390×844) with a real .docx: open, edit, reload (edits persist), export Word and PDF. Check the exported .docx opens (LibreOffice: `soffice --headless --convert-to pdf`) and its text matches.

## Rules of the code

- Fidelity first: the app only rewrites text inside existing `w:r` runs (`WordDoc.sync`). Never regenerate the document; everything else in the zip is kept byte for byte.
- Runs holding anything other than text are locked (`data-l`) and must survive edits untouched.
- Stored data: IndexedDB `individualnyi-plan`/`docs`, `DocRecord` in `src/store.ts`. Documents already on the phone must keep opening: only add optional fields.
- `src/docx/css.ts` is shared by the editor and the PDF renderer (which draws pages outside the app stylesheet).

## Sync and the data repository

- Documents sync through the **private** repository `alex-Drop2201/individualnyi-plan-data` (`src/sync.ts`, GitHub contents API, token kept on the phone in localStorage `ip.sync`). Layout: `docs/<id>.docx` (the whole document as it is now) and `docs/<id>.json` (`{ "name", "created", "updated" }`, `updated` in ms).
- The app pulls on start, when it comes to the foreground and every minute; it pushes edits a few seconds after they are made. A document changed in both places: the repository wins, the phone's version becomes a new document "… (с телефона)".
- To change a document for the owner: pull the data repo, edit `docs/<id>.docx` in place (see below), bump `updated` in its `.json`, commit, push. To add one: pick a new id (`[a-z0-9]+`), add both files. It appears on the phone within a minute of the app being open.

## Filling a document from the owner's text

The owner sends instructions in chat (and a .docx, or names a document already in the data repo). Fill it with a script that edits `word/document.xml` (unzip → change `w:t` text in place, keeping run formatting → zip, every other zip entry unchanged), check it logically, ask about anything unclear, leave fields without data empty, render it to check (LibreOffice), then put it in the data repo as above. Never put the owner's documents in this public repo.

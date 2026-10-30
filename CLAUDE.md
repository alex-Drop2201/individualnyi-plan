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

## Filling a document from the owner's text

The owner sends a .docx and instructions in chat. Fill it with a script that edits `word/document.xml` (unzip → change `w:t` text in place, keeping run formatting → zip), check it logically, ask about anything unclear, render it to check, and send the file back; she opens it in the app with «Открыть файл».

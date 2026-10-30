# Индивидуальный план: read this first

This `main` branch is **only the built site** that GitHub Pages serves. Don't edit files here.

The source code lives in the **`source`** branch of this repository:

```sh
git fetch origin source
git checkout source
npm ci
```

Then follow `CLAUDE.md` in the `source` branch. Publish with `npm run deploy`.

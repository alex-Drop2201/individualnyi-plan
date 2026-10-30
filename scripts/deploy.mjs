// Builds the app and publishes dist/ to the `main` branch, which GitHub Pages serves.
// Usage: npm run deploy   (run from the `source` branch, after committing your changes)
import { execSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const sh = (cmd, opts = {}) => execSync(cmd, { cwd: root, stdio: "inherit", ...opts });
const out = (cmd, cwd = root) => execSync(cmd, { cwd, encoding: "utf8" }).trim();

const dirty = out("git status --porcelain") !== "";
if (dirty && !process.argv.includes("--allow-dirty")) {
  console.error("There are uncommitted changes. Commit and push them to `source` first (or pass --allow-dirty).");
  process.exit(1);
}
const label = out("git rev-parse --short HEAD") + (dirty ? "+local" : "");

sh("npm run build");

const work = fs.mkdtempSync(path.join(os.tmpdir(), "ip-deploy-"));
try {
  sh("git fetch origin main");
  sh(`git worktree add --detach "${work}" origin/main`);
  for (const f of fs.readdirSync(work)) if (f !== ".git") fs.rmSync(path.join(work, f), { recursive: true, force: true });
  fs.cpSync(path.join(root, "dist"), work, { recursive: true });
  // Notes for anyone (or any Claude session) that opens the `main` branch.
  for (const f of fs.readdirSync(path.join(root, "deploy"))) fs.copyFileSync(path.join(root, "deploy", f), path.join(work, f));
  sh("git add -A", { cwd: work });
  if (out("git status --porcelain", work) === "") {
    console.log("The build is the same as what is already deployed. Nothing to do.");
  } else {
    // DEPLOY_TRAILERS: optional extra lines for the commit message (e.g. Co-Authored-By).
    const trailers = process.env.DEPLOY_TRAILERS ? `\n\n${process.env.DEPLOY_TRAILERS}` : "";
    fs.writeFileSync(path.join(work, ".git-message"), `Deploy build of ${label}${trailers}\n`);
    sh(`git commit -q -F .git-message`, { cwd: work });
    fs.rmSync(path.join(work, ".git-message"));
    sh("git push origin HEAD:main", { cwd: work });
    console.log(`Deployed ${label}. GitHub Pages updates in a minute or two.`);
  }
} finally {
  sh(`git worktree remove --force "${work}"`);
}

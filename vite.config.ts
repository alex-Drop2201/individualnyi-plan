import { readFileSync } from "node:fs";
import { fileURLToPath, URL } from "node:url";
import { defineConfig, type Plugin } from "vite";

// Writes the service worker (src/sw.js) into the build with this build's version and the
// list of built files, so the installed app works offline and updates on each deploy.
function serviceWorker(): Plugin {
  return {
    name: "service-worker",
    apply: "build",
    generateBundle(_, bundle) {
      const files = Object.keys(bundle).filter((f) => f.startsWith("assets/") && !f.endsWith(".map")).sort();
      const template = readFileSync(fileURLToPath(new URL("./src/sw.js", import.meta.url)), "utf8");
      this.emitFile({
        type: "asset",
        fileName: "sw.js",
        source: template.replace("__VERSION__", String(Date.now())).replace("__FILES__", JSON.stringify(files)),
      });
    },
  };
}

export default defineConfig({
  base: "./",
  plugins: [serviceWorker()],
  build: { chunkSizeWarningLimit: 1500 },
});

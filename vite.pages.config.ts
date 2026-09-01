import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { copyFileSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

const publicFiles = [
  "favicon.svg",
  "file.svg",
  "globe.svg",
  "local-bond-daily-converter.html",
  "og.png",
  "orient-ficc-logo.png",
  "templates/weekly-bond-report-template.docx",
  "window.svg",
];

export default defineConfig({
  root: "github-pages",
  base: "/bond-report-workbench/",
  publicDir: false,
  plugins: [
    react(),
    {
      name: "github-pages-public-files",
      closeBundle() {
        const output = resolve("docs");
        publicFiles.forEach((file) => {
          const target = resolve(output, file);
          mkdirSync(dirname(target), { recursive: true });
          copyFileSync(resolve("public", file), target);
        });
        writeFileSync(resolve(output, ".nojekyll"), "");
      },
    },
  ],
  define: {
    "import.meta.env.VITE_STORAGE_MODE": JSON.stringify("local"),
  },
  build: {
    outDir: "../docs",
    emptyOutDir: true,
  },
});

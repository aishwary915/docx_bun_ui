import { reactRouter } from "@react-router/dev/vite";
import tailwindcss from "@tailwindcss/vite";
import { defineConfig } from "vite";
import tsconfigPaths from "vite-tsconfig-paths";
import path from "path";

export default defineConfig({
  plugins: [tailwindcss(), reactRouter(), tsconfigPaths()],
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./app"),
    },
  },
  ssr: {
    // Exclude browser-only Univer packages from server-side bundle
    noExternal: [],
    external: [
      "@univerjs/core",
      "@univerjs/docs",
      "@univerjs/docs-ui",
      "@univerjs/engine-render",
      "@univerjs/presets",
      "@univerjs/preset-docs-core",
      "@univerjs/preset-sheets-core",
    ],
  },
});

import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/cli.ts"],
  format: ["esm"],
  outDir: "dist",
  target: "node18",
  // fully self-contained: dist/cli.js runs standalone
  noExternal: [/@stitchdocx\//, "diff", "fflate"],
});

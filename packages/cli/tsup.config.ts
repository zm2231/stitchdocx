import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/cli.ts"],
  format: ["esm"],
  outDir: "dist",
  target: "node18",
  // fully self-contained: the harness invokes dist/cli.js directly
  noExternal: [/@stitchdocx\//, "diff", "fflate"],
});

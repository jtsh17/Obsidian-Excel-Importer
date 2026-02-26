const esbuild = require("esbuild");
const isWatch = process.argv.includes("--watch");

const config = {
  entryPoints: ["src/main.ts"],
  bundle: true,
  format: "cjs",
  target: "es2020",
  platform: "node",
  outfile: "main.js",
  external: ["obsidian"],
  sourcemap: isWatch ? "inline" : false,
};

if (isWatch) {
  esbuild
    .context(config)
    .then((ctx) => {
      ctx.watch();
      console.log("Watching for changes...");
    })
    .catch(() => process.exit(1));
} else {
  esbuild.build(config).catch(() => process.exit(1));
}

import esbuild from "esbuild";
import process from "process";

async function buildContext(watch) {
  const context = await esbuild.context({
    entryPoints: ["src/main.ts"],
    bundle: true,
    external: ["obsidian"],
    format: "cjs",
    target: "es2018",
    platform: "browser",
    outfile: "main.js",
    sourcemap: watch
  });
  if (watch) {
    await context.watch();
  } else {
    await context.rebuild();
    await context.dispose();
  }
}

async function main() {
  const args = process.argv.slice(2);
  const watch = args.includes("dev");
  await buildContext(watch);
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});


/**
 * Builds the publishable package into dist/.
 *
 * Unlike a manifest-rewriting build, this never touches tracked files:
 * package.json already points at dist/ and `files` ships only that directory,
 * so the build is pure output and `git status` stays clean afterwards.
 */

import { $ } from "bun";
import path from "node:path";

const root = import.meta.dir;
const entrypoint = path.join(root, "index.ts");

async function bundle(format: "esm" | "cjs"): Promise<boolean> {
  const extension = format === "esm" ? "mjs" : "cjs";

  const result = await Bun.build({
    entrypoints: [entrypoint],
    outdir: path.join(root, "dist"),
    format,
    // Runtime dependencies and the engine peer stay external: consumers
    // resolve their own copies, and bundling the engine would duplicate it.
    packages: "external",
    target: "browser",
    sourcemap: "external",
    naming: `[name].${extension}`,
  });

  for (const log of result.logs) {
    console.log(`[${log.level}] ${log.message}`);
  }

  if (!result.success) {
    return false;
  }

  console.log(`Bundled dist/index.${extension}`);
  return true;
}

async function emitTypes(): Promise<boolean> {
  const { stdout, stderr, exitCode } = await $`bunx --bun tsc -p tsconfig.build.json`
    .cwd(root)
    .nothrow();

  if (exitCode !== 0) {
    console.error(stderr.toString());
    console.log(stdout.toString());
    return false;
  }

  console.log("Emitted dist/types");
  return true;
}

await $`rm -rf dist`.cwd(root);

const results = await Promise.all([bundle("esm"), bundle("cjs"), emitTypes()]);

if (!results.every(Boolean)) {
  throw new Error("Build failed");
}

console.log("Build complete");

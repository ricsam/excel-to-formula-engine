/**
 * Minimal static server for the demo page. Run with `bun run demo:serve`.
 */

import path from "node:path";

const root = import.meta.dir;
const port = Number(process.env["PORT"] ?? 4321);

Bun.serve({
  port,
  async fetch(request) {
    const url = new URL(request.url);
    const name = url.pathname === "/" ? "/index.html" : url.pathname;
    // Keep the server inside the demo directory.
    const resolved = path.join(root, path.normalize(name));
    if (!resolved.startsWith(root)) {
      return new Response("forbidden", { status: 403 });
    }
    const file = Bun.file(resolved);
    return (await file.exists())
      ? new Response(file)
      : new Response("not found", { status: 404 });
  },
});

console.log(`demo running at http://localhost:${port}`);

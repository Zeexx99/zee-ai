import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";
import { appendFileSync } from "node:fs";
import { getHttpsServerOptions } from "office-addin-dev-certs";

/** Dev-only: the task pane posts its errors here and they are appended to client-log.txt. */
function clientLog(): Plugin {
  const file = fileURLToPath(new URL("./client-log.txt", import.meta.url));
  return {
    name: "zee-client-log",
    configureServer(server) {
      attach(server.middlewares);
    },
    configurePreviewServer(server) {
      attach(server.middlewares);
    },
  };
  function attach(middlewares: { use: (path: string, fn: (req: import("node:http").IncomingMessage, res: import("node:http").ServerResponse) => void) => void }) {
      middlewares.use("/__zee_log", (req, res) => {
        let body = "";
        req.on("data", (c) => (body += c));
        req.on("end", () => {
          try {
            appendFileSync(file, `[${new Date().toISOString()}] ${body}\n`);
          } catch {
            /* ignore */
          }
          res.statusCode = 204;
          res.end();
        });
      });
  }
}

// Office requires the task pane to be served over HTTPS.
// office-addin-dev-certs creates and trusts a localhost certificate (run `npm run certs` once).
// Port 3100 (not 3000) to avoid clashing with other local dev servers.
export default defineConfig(async ({ command }) => {
  // No certificates needed for `vite build` (CI / GitHub Pages).
  const https = command === "serve" ? await getHttpsServerOptions() : undefined;
  return {
    // Relative base so the same build works on https://localhost:3100/ and https://<user>.github.io/zee-ai/
    base: "./",
    plugins: [react(), clientLog()],
    server: { port: 3100, strictPort: true, https },
    preview: { port: 3100, strictPort: true, https },
    build: {
      outDir: "dist",
      sourcemap: true,
      chunkSizeWarningLimit: 900,
      rollupOptions: {
        input: { taskpane: fileURLToPath(new URL("./taskpane.html", import.meta.url)) },
        output: {
          // Split big vendor libraries so the pane can cache them separately.
          manualChunks: {
            fluent: ["@fluentui/react-components"],
            icons: ["@fluentui/react-icons"],
            markdown: ["react-markdown", "remark-gfm", "rehype-highlight", "lowlight"],
          },
        },
      },
    },
  };
});

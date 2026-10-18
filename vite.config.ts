import { defineConfig } from "vitest/config";

// The built site makes no network requests: no analytics, no uploads, no CDN.
// This CSP is added only to the production build because Vite's dev server
// needs a websocket for hot reload.
const CSP = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob:",
  "connect-src 'none'",
  "form-action 'none'",
  "base-uri 'none'",
  "object-src 'none'",
].join("; ");

export default defineConfig({
  base: "./",
  plugins: [
    {
      name: "inject-csp",
      apply: "build",
      transformIndexHtml(html) {
        return html.replace(
          "<head>",
          `<head>\n    <meta http-equiv="Content-Security-Policy" content="${CSP}" />`,
        );
      },
    },
  ],
  test: {
    environment: "jsdom",
    include: ["tests/**/*.test.ts"],
  },
});

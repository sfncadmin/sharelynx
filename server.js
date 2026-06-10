// Minimal HTTPS static server for local add-in development.
// Uses the certs installed by `npm run certs` (office-addin-dev-certs).
const https = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = Number(process.env.PORT) || 3000;
const PLAIN_HTTP = process.argv.includes("--http"); // UI-preview only; Outlook requires HTTPS
const ROOT = __dirname;
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xml": "text/xml; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".map": "application/json",
  ".md": "text/plain; charset=utf-8",
  ".woff2": "font/woff2"
};

let options = {};
if (!PLAIN_HTTP) {
  try {
    options = {
      key: fs.readFileSync(path.join(CERT_DIR, "localhost.key")),
      cert: fs.readFileSync(path.join(CERT_DIR, "localhost.crt"))
    };
  } catch (e) {
    console.error("Dev certificates not found. Run `npm run certs` first.");
    process.exit(1);
  }
}

(PLAIN_HTTP ? http : https)
  .createServer(options, (req, res) => {
    const urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
    let filePath = path.normalize(path.join(ROOT, urlPath));
    if (!filePath.startsWith(ROOT)) {
      res.writeHead(403);
      return res.end("Forbidden");
    }
    if (urlPath === "/") {
      res.writeHead(302, { Location: "/src/taskpane/taskpane.html" });
      return res.end();
    }
    fs.stat(filePath, (err, stat) => {
      if (!err && stat.isDirectory()) filePath = path.join(filePath, "index.html");
      fs.readFile(filePath, (err2, data) => {
        if (err2) {
          res.writeHead(404, { "Content-Type": "text/plain" });
          return res.end("Not found: " + urlPath);
        }
        res.writeHead(200, {
          "Content-Type": MIME[path.extname(filePath).toLowerCase()] || "application/octet-stream",
          "Cache-Control": "no-store"
        });
        res.end(data);
      });
    });
  })
  .listen(PORT, () => {
    const scheme = PLAIN_HTTP ? "http" : "https";
    console.log(`SFNC Share Links dev server running at ${scheme}://localhost:${PORT}`);
    console.log(`Taskpane: ${scheme}://localhost:${PORT}/src/taskpane/taskpane.html`);
  });

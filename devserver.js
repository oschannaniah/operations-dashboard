// Local dev server — plain static file server. The old /api/sheet proxy to Apps Script is
// gone now that the frontend talks to Supabase directly from the browser (no CORS workaround
// needed the way the Sheets backend required).
const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = process.env.PORT || 8734;

const server = http.createServer(async (req, res) => {
  let filePath = req.url === "/" ? "/index.html" : req.url;
  filePath = path.join(__dirname, filePath.split("?")[0]);
  fs.readFile(filePath, (err, data) => {
    if (err) { res.writeHead(404); res.end("not found"); return; }
    const ext = path.extname(filePath);
    const type = { ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".png": "image/png", ".svg": "image/svg+xml" }[ext] || "text/html";
    res.writeHead(200, { "Content-Type": type });
    res.end(data);
  });
});

server.listen(PORT, () => console.log(`operations-dashboard dev server on http://localhost:${PORT}`));

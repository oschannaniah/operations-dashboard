// Vercel serverless function — proxies requests to the Google Apps Script backend.
//
// Why this exists: calling script.google.com directly from the browser hits a CORS wall
// that Apps Script's "who has access" setting can't fix — Google internally redirects the
// request to script.googleusercontent.com to serve the actual response, and that redirect
// doesn't carry the right CORS headers. Browsers block it regardless of deployment settings.
//
// The fix: the browser calls THIS endpoint instead (same origin as the dashboard, so no CORS
// applies at all), and this function calls Apps Script server-to-server, where CORS is a
// browser-only concept and simply doesn't apply. It relays the response straight back.
//
// No configuration needed on Vercel's side — any file under /api is auto-detected as a
// serverless function.

const APPS_SCRIPT_URL = "https://script.google.com/a/macros/oscfamily.com/s/AKfycbw16CyoaHFSXYAj8Zn7SPDNHvw3I4TKyRRvL1P1zxVSnPj1F36sltdj47pBbfCxZP74Og/exec";

module.exports = async (req, res) => {
  try {
    if (req.method === "GET") {
      const params = new URLSearchParams(req.query).toString();
      const upstream = await fetch(`${APPS_SCRIPT_URL}?${params}`);
      const text = await upstream.text();
      res.status(upstream.status).setHeader("Content-Type", "application/json").send(text);
      return;
    }

    if (req.method === "POST") {
      // Sent to Apps Script as text/plain — Apps Script's doPost just JSON.parses the raw
      // body regardless of header, and this avoids any quirks on that side. This is a
      // server-to-server call, so none of the browser CORS-preflight concerns apply here.
      const bodyText = typeof req.body === "string" ? req.body : JSON.stringify(req.body || {});
      const upstream = await fetch(APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: bodyText,
      });
      const text = await upstream.text();
      res.status(upstream.status).setHeader("Content-Type", "application/json").send(text);
      return;
    }

    res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    res.status(500).json({ error: String(err && err.message || err) });
  }
};

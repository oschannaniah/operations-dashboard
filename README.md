# operations-dashboard

OpsCore — campus operations dashboard. Static React app (`bundle.js`/`output.css`, built from
`src/`) plus a Vercel serverless proxy (`api/sheet.js`) in front of a Google Apps Script + Sheets
backend (`Code.gs`, kept outside this repo — ask Han for the current copy).

## Rebuilding after editing src/entry_component.jsx

```
cd src
npm install
npx esbuild entry.jsx --bundle --minify --outfile=../bundle.js --loader:.js=jsx
npx tailwindcss -i input.css -o ../output.css --minify
```

Commit the regenerated `bundle.js`/`output.css` alongside the `src/` change and push — Vercel
redeploys on push to `main`. `index.html` just loads those two files; no build step runs on Vercel.

## Auth

Login/session backend lives in Apps Script (`Code.gs`), not in this repo. It requires a `Users`
sheet tab and an `AUTH_SECRET` Script Property. New accounts self-register (`action: "register"`)
and land as tier `"unassigned"` until a Central account assigns them a campus + role from the
Team Accounts panel — that reassignment takes effect on the user's next sign-in.

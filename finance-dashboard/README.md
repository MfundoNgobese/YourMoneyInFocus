# Finance Dashboard

A single-page React app for South African personal finance: import bank
statements (Absa, FNB, Standard Bank, Capitec) and EasyEquities/TFSA activity,
track spending and savings, and project tax-free retirement growth.

Everything runs in the browser. No data leaves the device; imports and tags are
stored locally in the browser only.

## Deploy (free) — Vercel or Netlify

You do **not** need to install anything on your computer. Both services build
the project in the cloud from your GitHub repo.

### 1. Put this folder on GitHub
- Create a new **public** repository (e.g. `finance-dashboard`).
- Upload the entire contents of this folder (Add file → Upload files → drag
  everything in → Commit).

### 2a. Vercel
1. Go to vercel.com → Sign up → Continue with GitHub.
2. Add New… → Project → Import your `finance-dashboard` repo.
3. Vercel auto-detects Vite (Build: `npm run build`, Output: `dist`).
   Leave the defaults → Deploy.
4. You get a live URL like `finance-dashboard-xxxx.vercel.app`.

### 2b. Netlify (alternative)
1. Go to netlify.com → Sign up → Continue with GitHub.
2. Add new site → Import an existing project → pick the repo.
3. It auto-detects Vite (Build: `npm run build`, Publish: `dist`) → Deploy.

Every push to the repo triggers an automatic rebuild and redeploy.

### Also works on GitHub Pages
The build uses relative asset paths (`base: './'`), so it also runs under a
GitHub Pages sub-path if you prefer that route.

## Run locally (optional)
```bash
npm install
npm run dev      # local dev server
npm run build    # production build into dist/
npm run preview  # preview the production build
```
Requires Node.js 18+.

## Notes
- PDF statement parsing loads pdf.js from a CDN at runtime (needs internet the
  first time a PDF is imported).
- Category tags persist via the browser's local storage on the hosted site.
- Growth projections are illustrative only and not financial advice.

# Finance Dashboard

Single-page React app (Vite) for South African personal finance: import bank
statements (Absa, FNB, Standard Bank, Capitec) and EasyEquities/TFSA activity,
track spending and savings, and project tax-free retirement growth. Runs entirely
in the browser; no data leaves the device.

## Deploy free on Vercel or Netlify
1. Create a PUBLIC GitHub repo and upload ALL files in this folder to the repo ROOT.
   (There is no `src` folder — every file sits at the top level next to index.html.)
2. Vercel: vercel.com -> Continue with GitHub -> Add New… -> Project -> import the
   repo -> it auto-detects Vite -> Deploy.
   Netlify: netlify.com -> Import an existing project -> pick the repo -> Deploy.

Every push rebuilds automatically.

## Files (all at root)
index.html, main.jsx, App.jsx, index.css, package.json, vite.config.js

## Local (optional, Node 18+)
npm install && npm run dev

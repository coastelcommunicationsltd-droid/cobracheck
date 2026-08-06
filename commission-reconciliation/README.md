# Commission Reconciliation (test tool)

Standalone, browser-only tool that reconciles **Cobra**, **NetSuite** and **Sch5**
exports. Everything runs client-side — the spreadsheets are read inside the
browser and never uploaded anywhere. The host only ever serves the app code.

Access code (test gate only): **BTLBDCSDTEST**

---

## 1. Run it on your own machine (to check it works)

You'll need Node.js 18+ installed (same as your other Vite project).

```bash
npm install      # one time — pulls in React + the xlsx library
npm run dev      # starts it locally, prints a http://localhost:5173 URL
```

Open the URL, enter the access code, drop the three .xlsx files in.

## 2. Build the files a host can serve

```bash
npm run build    # produces a /dist folder — that folder IS the website
```

Everything a host needs is inside `dist/`.

---

## 3. Put it online

### Option A — Cloudflare Pages + Access  (recommended: restricts to named people, free)
1. Push this folder to a GitHub repo (or use Cloudflare's direct upload of `dist/`).
2. Cloudflare dashboard → **Workers & Pages** → Create → Pages → connect the repo.
   - Build command: `npm run build`
   - Build output directory: `dist`
3. Once live, go to **Zero Trust → Access → Applications → Add a self-hosted app**,
   point it at the Pages URL, and add an **Allow** policy listing the exact email
   addresses of the few people who should get in. Everyone else is blocked at the door.
   (Free plan covers up to 50 users.)

### Option B — Netlify or Vercel  (fastest way to a live URL)
- Netlify: drag the `dist/` folder onto https://app.netlify.com/drop — live instantly.
- Vercel: `npm i -g vercel` then `vercel` in this folder, or connect the GitHub repo.
- Note: the tool's own passcode is the only gate here unless you add their
  (paid) password protection — use Option A for proper access control.

### Option C — internal hosting
If it should sit behind the company login/VPN alongside the main system,
hand the `dist/` folder (or this repo) to whoever runs that infrastructure.

---

## Changing the access code
Open `src/App.jsx`, find the line near the top:

```js
const PASSCODE = "BTLBDCSDTEST";
```

Change the value, rebuild (`npm run build`), redeploy. This is light gating only,
not real security — real per-user access comes when this folds into the main system.

## Notes on matching
- Join key is the BT order number (Cobra "Job Header" = NetSuite "Order ref" = Sch5 "MAIN ORDER NUM").
- Load exports covering the **same** orders/period, or nothing will cross-match.
- Match tolerance (£ / %) is adjustable on the upload panel.

# Rent Roll → Financial Model

Upload any rent roll (CSV or Excel). AI parses the data. Download a complete 6-sheet Excel financial model.

## What it generates

- **INPUTS** — all assumptions in one place
- **RENT ROLL** — unit-by-unit with actual and pro forma rents
- **CASH FLOW** — full I&E with current vs. pro forma columns
- **PRICING** — cap rate, GRM, $/SF, $/unit, DSCR, comparable analysis
- **10 YEAR DCF** — full projection with IRR, NPV, equity multiple, DSCR by year
- **VALUE SUMMARY** — recommended listing price and key metrics

---

## Option A — Deploy to Vercel (web app with shareable link)

### Prerequisites
- [Node.js 18+](https://nodejs.org)
- [Git](https://git-scm.com)
- A free [Vercel account](https://vercel.com) (sign up with GitHub)

### Steps

1. **Install dependencies**
   ```bash
   npm install
   ```

2. **Push to GitHub**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/rentroll-app.git
   git push -u origin main
   ```

3. **Deploy on Vercel**
   - Go to [vercel.com/new](https://vercel.com/new)
   - Click **"Import Git Repository"**
   - Select your repo
   - Leave all settings as defaults — Vercel auto-detects Next.js
   - Click **Deploy**
   - Your app is live at `https://your-app-name.vercel.app`

4. **Every future update:**
   ```bash
   git add . && git commit -m "update" && git push
   ```
   Vercel auto-redeploys on every push.

---

## Option B — Desktop App (Mac .dmg / Windows .exe installer)

### Prerequisites
- Node.js 18+

### Install dependencies
```bash
npm install
```

### Run in development (app + live reload)
```bash
npm run electron-dev
```

### Build Mac installer (.dmg)
```bash
npm run electron-build-mac
```
Output: `dist/Rent Roll Model-1.0.0.dmg`
- Double-click the .dmg → drag to Applications → done

### Build Windows installer (.exe)
```bash
npm run electron-build-win
```
Output: `dist/Rent Roll Model Setup 1.0.0.exe`
- Run the .exe → installs like any Windows app

### Build both at once
```bash
npm run electron-build-all
```

> **Note for Mac:** If you see "App can't be opened because it's from an unidentified developer", right-click → Open → Open anyway. To avoid this, you'd need an Apple Developer certificate ($99/yr).

---

## Running locally (browser only)

```bash
npm install
npm run dev
```
Open [http://localhost:3000](http://localhost:3000)

---

## Project structure

```
rentroll-app/
├── src/
│   ├── app/
│   │   ├── layout.js        # Root layout + fonts
│   │   ├── page.js          # Entry point
│   │   └── globals.css      # Base styles
│   └── components/
│       ├── App.js           # Main UI (upload → assumptions → preview → download)
│       ├── finance.js       # Financial model calculations (IRR, DCF, NOI etc.)
│       └── excel.js         # Excel workbook generator (6 sheets)
├── electron/
│   └── main.js              # Electron desktop wrapper
├── public/                  # Static assets (add icon.png, icon.ico, icon.icns here)
├── package.json
├── next.config.js
└── vercel.json
```

---

## Customizing default expense assumptions

Edit the `DEFAULT` object at the top of `src/components/App.js`:

```js
const DEFAULT = {
  expenses: {
    insurance: 1200,      // $ per unit per year
    waterSewer: 700,
    superSalary: 500,
    repairsMaint: 750,
    commonElectric: 0.25, // $ per SF
    fuelOil: 900,
    managementPct: 0.04,  // 4% of EGI
    vacancyPct: 0.03,     // 3% of GPR
  },
  rsGrowthRate: 0.03,
  fmGrowthRate: 0.04,
  ...
}
```

## Adding your own app icon

Place these files in the `public/` folder:
- `icon.png` (512×512, for display)
- `icon.icns` (Mac — use [Image2icon](https://img2icnsapp.com/))
- `icon.ico` (Windows — use [convertico.com](https://convertico.com/))

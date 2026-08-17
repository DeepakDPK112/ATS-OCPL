# OCPL ATS — Quick Setup Instructions

## Folder structure to create on your computer

```
ocpl-ats/
├── index.html               ← copy from downloads
├── package.json             ← copy from downloads
├── vite.config.js           ← copy from downloads
├── .gitignore               ← copy from downloads
├── .env.local               ← copy from downloads (contains Supabase keys)
└── src/
    ├── main.jsx             ← copy from downloads
    ├── App.jsx              ← rename recruiting_pipeline.jsx → App.jsx
    └── supabase.js          ← copy from downloads
```

## Steps

**1. Install Node.js** (if not already installed)
Download from https://nodejs.org — choose the LTS version

**2. Create the project folder**
Make a folder called `ocpl-ats` anywhere on your computer (e.g. Desktop)

**3. Copy all the downloaded files** into the folder exactly as shown above
- Create a subfolder called `src` inside `ocpl-ats`
- Put `main.jsx`, `App.jsx`, and `supabase.js` inside `src/`
- Put everything else in the root `ocpl-ats/` folder

**4. Open Terminal / Command Prompt**
Navigate to the project folder:
```
cd Desktop/ocpl-ats
```

**5. Install dependencies**
```
npm install
```

**6. Run locally to test**
```
npm run dev
```
Open http://localhost:5173 — sign in, add a candidate, refresh — data should persist.

**7. Push to GitHub**
```
git init
git add .
git commit -m "OCPL ATS initial commit"
```
- Go to https://github.com/new
- Create a private repo called `ocpl-ats`
- Follow the commands GitHub shows to push

**8. Deploy on Vercel**
- Go to https://vercel.com → Add New Project → Import `ocpl-ats`
- Add Environment Variables:
  - `VITE_SUPABASE_URL` = https://ozjdwqzyizyudsoxaiss.supabase.co
  - `VITE_SUPABASE_ANON_KEY` = sb_publishable_qbxp7oiahZFhOC_Z0-vgTA_9u5aeeB2
- Click Deploy → live in ~60 seconds

## Future updates
Every time you make changes:
```
git add .
git commit -m "describe change"
git push
```
Vercel redeploys automatically.

# MailCraft AI — Deploy to Vercel in 3 steps

## What you're deploying
A tool where anyone can upload an Excel email brief → 5 AI agents → download production HTML email.
No backend. No database. Runs entirely in the browser.

---

## Step 1 — Install Vercel CLI (one time only)
```
npm install -g vercel
```

## Step 2 — Put these 2 files in a folder called `mailcraft`
- index.html  (the app)
- vercel.json (the config)

## Step 3 — Deploy
```
cd mailcraft
vercel --prod
```
- First time: it asks you to log in with GitHub/email → do that
- Project name: mailcraft-ai (or anything you want)
- Hit Enter for all other questions

✅ Done. Vercel gives you a URL like:
   https://mailcraft-ai.vercel.app

Share that URL with your manager.

---

## What your manager needs to do
1. Open the URL
2. Click "API Key" in the top right → enter their Anthropic API key (sk-ant-...)
   - Get one at https://console.anthropic.com (free account, pay per use ~$0.01/email)
   - Key is saved in their browser only, never sent anywhere else
3. Upload any Excel brief
4. Watch 5 agents run → download the HTML email

---

## Files
- index.html  — The entire app (React + all 5 agents + UI, self-contained)
- vercel.json — Tells Vercel to serve index.html as a static site

No npm install, no build step, no server needed.

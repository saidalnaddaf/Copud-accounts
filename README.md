# CUPOD Student Registration

A modern web form for advisors to register CUPOD students directly into a Google Sheet.

## Stack

- HTML / CSS / Vanilla JavaScript (no framework, no build step)
- Google Apps Script as the backend
- Hosted free on Netlify or Vercel

## Quick start

1. Open `SETUP_GUIDE.md` and follow the steps in order.
2. Paste your Apps Script URL into `js/config.js`.
3. Drag the project folder onto [Netlify Drop](https://app.netlify.com/drop).

That's it.

## File map

```
cupod-registration/
├── index.html             form page
├── css/styles.css         styles
├── js/config.js           ← your URLs
├── js/main.js             form logic
├── AppsScript.gs          backend (paste into Apps Script)
└── SETUP_GUIDE.md         full instructions
```

## Live site

After deploying to Netlify, paste your URL here so your team can find it.

# Using Darkwrote on a Samsung Galaxy Z Fold 4

Darkwrote is a Progressive Web App: you install it from Chrome and it then behaves like a normal app —
its own icon, full screen, works offline. Your documents and notes never leave the phone.

## 1. Put the app online once (HTTPS is required for installing)

**Recommended — GitHub Pages (free, automatic):**

1. Merge the pull request into `main`.
2. In the repository: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. Wait for the **Deploy to GitHub Pages** workflow to finish (Actions tab, about a minute).
4. Your app is at `https://<your-github-user>.github.io/darkwrote/`
   (for this repository: `https://grendelpress.github.io/darkwrote/`).

Any static host works too: run `npm install && npm run build` and upload the `dist/` folder
(Netlify Drop, Cloudflare Pages, your own server). It must be served over **https**.

## 2. Install it on the Fold

1. Open the address above in **Chrome** on the Fold.
2. Chrome menu (⋮) → **Add to Home screen** → **Install**. (Or use **Install app** in Darkwrote's ☰ menu
   if Chrome offers it.)
3. Open **Darkwrote** from the home screen / app drawer. Open it once while online; after that it works
   with no connection. Fold and unfold freely — the layout adapts and your place is kept.

## 3. Getting a manuscript in

- **From inside Darkwrote:** ☰ → **Open** → pick the `.docx` (My Files, Google Drive, Downloads…).
- **From another app:** in My Files / Drive / Gmail choose **Share → Darkwrote**. The file opens straight away.
- **Reopening:** the start screen lists documents you've opened, with your reading progress. The app
  also reopens the last document automatically.

## 4. Reading and annotating

- The phone opens documents in **Reading mode**: text reflows to the screen and can't be changed by accident.
  **Aa** sets text size, font and line spacing. **✎ Edit** switches to editing (**📖 Read** switches back).
- **Long-press a word and drag the handles** to select. A bar appears at the bottom with four highlight colours,
  *Remove highlight*, **Comment**, and **Copy**. The comment box sits above the on-screen keyboard.
- **Samsung's own menu:** a long press also pops up Samsung's Copy/Share/Translate menu. A web app can't switch that
  off, so Darkwrote works alongside it: our bar appears on the opposite side of the screen from your selection,
  and the passage stays marked (and the bar stays up) even if you dismiss Samsung's menu by tapping elsewhere.
  Tap **✕** on our bar when you're done.
- Tap highlighted/commented text to open that comment. The outline button (☷) jumps by heading;
  the 💬 button lists all comments.
- Notes, highlights, edits and your reading position are **saved automatically on the phone**, even if Android
  closes the app. To free space, remove a document from the start screen (✕); the original file is untouched.

## 5. Getting your annotations out

☰ → **Export…** (on the Fold this is also what **Save** does):

| Option | What you get |
| --- | --- |
| **Original file + your annotations (.docx)** | A copy of the *original* file with your comments and highlights added as real Word comments/highlights. Nothing else in the file is touched — headers, footnotes, styles, text boxes, tracked changes all stay exactly as they were. |
| Annotations only — Markdown | A readable list of comments and highlights grouped by heading, each with the passage quoted. |
| Annotations only — JSON | Machine-readable, with exact locations in the source document. |
| Edited document, rebuilt | Only needed if you **edited the text**. Rebuilt by Darkwrote, so features it can't display (headers, footnotes, text boxes…) are not carried over. |

**Download** saves to the phone's *Downloads* folder. **Share…** (when offered) sends it to Drive, email, etc.

If you edit the text, the first option is switched off (the original can no longer be patched) — export your
annotations separately, or use the rebuilt copy.

## Troubleshooting

- *No "Install" option:* make sure you're on the https address in Chrome (not an in-app browser) and reload once.
- *Update the app:* open it while online; when "A new version is ready" appears, tap **Reload**.
- *Phone storage:* Darkwrote asks Android to keep its data; clearing Chrome's site data removes saved notes,
  so export annotations you care about.

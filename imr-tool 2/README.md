# IMR Submissions Tool

A small Flask app for collecting contributor submissions and producing a
manuscript pre-styled with the IMR paragraph styles (Article Title, Byline,
Standfirst, Body Text First, Body Text, Subhead, Pull Quote, Caption,
Endnote Text, Footer) using Barlow Condensed and Source Serif 4.

## Why .docx, not .idml

The previous version of this tool tried to hand-build raw IDML and ran into
a long string of InDesign compatibility issues (ID clashes, missing
required Spread/Page attributes, font inheritance quirks) that needed
extensive trial and error to get partially working. This version sidesteps
all of that: it outputs a properly named-and-styled .docx. When you
`File > Place` that into an InDesign document that already has paragraph
styles with these same names (your existing IMR template), InDesign maps
the incoming text onto those styles automatically — no custom XML to
debug, and you can sanity-check the content by just opening the .docx
yourself first.

## Setup (Windows / PowerShell)

```powershell
cd path\to\imr-tool
pip install -r requirements.txt
python app.py
```

Then open **http://127.0.0.1:5000** in your browser.

- Contributor area: `/` — password `imr2026`
- Admin area: `/admin` — password `imradmin2026`

Change both passwords before sharing the contributor link with anyone —
either edit the defaults directly in `app.py`, or set environment
variables before running:

```powershell
$env:SUBMISSION_PASSWORD="your-password"
$env:ADMIN_PASSWORD="your-admin-password"
python app.py
```

## How it works

1. A contributor logs in at `/`, fills in title/author/standfirst, and
   uploads their manuscript as a `.docx`.
2. They see a confirmation message. They never see or download anything
   else.
3. You log in separately at `/admin`, see a list of all submissions, and
   download the styled `.docx` for each one.
4. In InDesign, place that `.docx` into your IMR layout — paragraph styles
   should pick up automatically if the style names match your template.

## Layout

Output follows the print layout of the actual journal: page 1 (title,
byline, standfirst, a fixed blank gap reserved for a photo, and a lead
chunk of running body text, roughly 350 words by default) is a single
column, fully justified, with a bit of space after each paragraph. From
there, subheads, pull quotes, the rest of the body, captions, and
endnotes all flow in two columns starting on page 2, also justified. A
subhead or pull quote always starts the two-column section immediately,
even if the word target hasn't been reached yet, since that's the
natural break point in print layout.

The blank gap after the standfirst is a fixed height (180pt, roughly
2.5") regardless of how long the standfirst or article is - every
submission gets exactly the same gap, since it's meant to reserve a
consistent space for artwork to be dropped in later, not to scale with
content. It's set as `IMAGE_GAP_PT` near the top of `processor.py`.

Margins sit at 0.7" top/bottom and 0.85" left/right - slimmer than
Word's US default (1"/1.25") so more text fits across the page, but not
razor-thin.

The 350-word intro target is a heuristic, not a guarantee of exactly
filling page 1 - actual fit depends on how Word/InDesign renders the
fonts. It's set near the top of `processor.py` as `INTRO_WORD_TARGET` if
you want to tune it; the margin values are right below it as
`MARGIN_TOP_IN` / `MARGIN_BOTTOM_IN` / `MARGIN_LEFT_IN` /
`MARGIN_RIGHT_IN`.

## Subhead auto-detection

If a contributor doesn't tag a section heading with `[SUBHEAD]`, the tool
tries to catch it anyway. Any plain paragraph that's short (8 words or
fewer), has no terminal punctuation, and has every meaningful word
capitalised (small connector words like "the", "of", "in" are allowed to
stay lowercase) gets treated as a subhead automatically - bolded and
styled the same as an explicitly tagged one. This catches things like
"Silicon Economy" or "The MAGA Turn" sitting on their own line without
needing the contributor to remember the marker syntax.

It's a heuristic, so it can occasionally misfire both ways: a genuine
short declarative sentence without a full stop could get mistaken for a
heading, or an unusual heading style might slip through as plain body
text. The `[SUBHEAD]` tag always works as an explicit override regardless
of what the heuristic decides. The detection rules live in
`_looks_like_subhead()` in `processor.py` if you want to tune the word
limit (`MAX_SUBHEAD_WORDS`) or the list of lowercase-allowed connector
words (`SUBHEAD_STOPWORDS`).

## Style markers

In the contributor's uploaded manuscript, any paragraph that starts with
one of these tags (on its own line, at the start of the paragraph) gets
mapped to the matching IMR style and the tag itself is stripped out:

- `[PULLQUOTE]`
- `[SUBHEAD]`
- `[CAPTION]`
- `[ENDNOTE]`

Word "Heading" styles are also mapped to Subhead automatically. Everything
else becomes Body Text (the very first paragraph becomes "Body Text First"
in case you want a drop cap or different opening treatment).

## Deploying to Railway

This repo is ready to deploy as-is — it includes a `Procfile`, `railway.json`,
and `gunicorn` in `requirements.txt`.

1. Push this folder to a GitHub repo (or use the Railway CLI to deploy
   directly without git — `railway up` from inside this folder, after
   `railway login` and `railway init`).
2. In Railway, **New Project → Deploy from GitHub repo**, pick the repo.
   Railway auto-detects it as a Python app via Nixpacks and uses the
   `Procfile`/`railway.json` start command — no extra config needed.
3. In the Railway project's **Variables** tab, set:
   - `SECRET_KEY` — any random string
   - `SUBMISSION_PASSWORD` — your contributor password
   - `ADMIN_PASSWORD` — your admin password
   - `DATA_DIR` — `/data` (see the Volume step below)
4. **Add a Volume**: Railway's filesystem is wiped on every redeploy, which
   would silently delete past submissions. Go to the service →
   **Settings → Volumes → New Volume**, mount it at `/data`. That's what
   `DATA_DIR=/data` points the app at, so uploads, generated `.docx` files,
   and `submissions.json` survive redeploys.
5. Railway will give you a public URL (e.g.
   `yourapp.up.railway.app`) — that's the link to send contributors,
   and `/admin` on the same domain is yours.

Without step 4, the app still works, but every submission gets wiped the
next time you redeploy — worth doing from day one rather than after losing
something.



- `app.py` — Flask routes, login, upload handling, admin dashboard
- `processor.py` — reads the uploaded .docx and builds the styled output
- `templates/` — the four pages (contributor login, submission form, admin
  login, admin dashboard)
- `submissions.json` — created automatically, logs every submission
- `uploads/` / `output/` — created automatically, store raw uploads and
  generated files

## If you want IDML output again

It's possible, but expect another round of InDesign-specific debugging
like last time — IDML's required attributes for Spreads/Pages aren't
fully documented and tend to only reveal themselves when InDesign refuses
to open a file. Worth doing only if the docx-and-Place workflow above
turns out to be too manual for your volume of submissions.

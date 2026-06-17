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

## Files

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

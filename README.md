# KNCCI Certificate Generator — Fixed Bulk Upload (GitHub Pages)

## What’s fixed vs previous build
- **Participants upload input is always visible** (not hidden behind mapping)
- **Auto-mapping + auto-apply** immediately after you upload CSV/XLSX
- Clear “Mapping applied” status and participant-ready count
- Includes `assets/participants_template.csv` for plug-and-play uploads

## Deploy
1. Replace `assets/template.pdf` with your real certificate template.
2. Upload all files to GitHub.
3. Enable GitHub Pages: Settings → Pages → Deploy from branch → `main` / `/root`.

## Strict font accuracy
Upload the exact font files (TTF/OTF) used in the certificate design:
- Regular font
- Bold font
Then set font sizes/positions and RGB.

## Note on wiping
“Wipe” draws a **white rectangle** behind the old text.
If the area behind course/date isn’t white in your design, the wipe may show. In that case, the correct fix is a clean template without that text layer, or a background-matched patch.

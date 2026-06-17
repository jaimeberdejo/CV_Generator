# CV Generator

A small, data-driven CV generator. Keep your CV in one YAML file and render it three ways from a
single source of truth:

- 🌐 **A live one-page website** (with an EN/ES language toggle) — published to GitHub Pages.
- 📄 **A polished Word document** (`.docx`) — and, on macOS, `.pages` + `.pdf`.
- 🖨️ **A print-ready PDF** straight from the website ("Download PDF" → Save as PDF).

Edit `cv_data.yml` once; everything regenerates. No copy-pasting between Word and the web.

> 🔗 **Live demo:** https://jaimeberdejo.github.io/CV_Generator/

## Why

Most people keep their CV trapped in a `.docx`. This project treats the *content* as structured data
(`cv_data.yml`) and the *presentation* as code, so a single edit propagates to a web version, a Word
version, and a PDF — all kept perfectly in sync, in two languages.

## Quick start

```bash
pip install -r requirements.txt

# Build the website -> ./site/index.html
python build_site.py

# Open it
open site/index.html        # macOS
# xdg-open site/index.html  # Linux
```

To regenerate the Word/PDF documents:

```bash
python generate_cv.py                 # EN + ES .docx (+ .pages/.pdf on macOS)
python generate_cv.py --no-pdf        # skip PDF
```

## Editing your CV

Everything lives in [`cv_data.yml`](cv_data.yml). It has:

- `personal` — name, phone, email, photo (shared across languages)
- `en` / `es` — each language's `title`, `location`, and the sections:
  `profile`, `education`, `experience`, `skills`, `languages`, `projects`

Change the data, rerun the build, done.

## Deploying to GitHub Pages

A GitHub Actions workflow ([`.github/workflows/deploy.yml`](.github/workflows/deploy.yml)) builds the
site and publishes it on every push to `main`.

1. Push this repo to GitHub.
2. **Settings → Pages → Build and deployment → Source: GitHub Actions.**
3. Push to `main`; the workflow builds `site/` and deploys it. Your CV goes live at
   `https://<user>.github.io/<repo>/`.

## How it works

```
cv_data.yml ─┬─► build_site.py  ──► Jinja2 (templates/cv.html.j2) + static/style.css ──► site/  ──► GitHub Pages
             └─► generate_cv.py ──► python-docx (direct OpenXML)                       ──► .docx / .pdf / .pages
```

- **`build_site.py`** renders the HTML site. Both languages are emitted into one page; a tiny vanilla
  JS toggle switches between them, and a print stylesheet (`@media print`) turns "Save as PDF" into a
  clean one-page CV.
- **`generate_cv.py`** builds the Word document by manipulating OpenXML directly for pixel-level
  control over spacing, borders, and date alignment.

Both renderers share the same navy/gray theme (`#1F3864` / `#595959`, Calibri) so the web and the
document look like the same CV.

## Tech stack

Python · Jinja2 · python-docx · PyYAML · HTML/CSS · GitHub Actions · GitHub Pages

## License

[MIT](LICENSE)

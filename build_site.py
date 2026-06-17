#!/usr/bin/env python3
"""
Static-site builder for Jaime Berdejo's CV.

Reads the same `cv_data.yml` used by `generate_cv.py` and renders a single,
self-contained one-page CV website (`index.html`) with an EN/ES language
toggle and a "Download PDF" button (browser print-to-PDF). The output folder
is ready to publish on GitHub Pages.

Usage:
    python3 build_site.py                 # -> ./site/
    python3 build_site.py --out docs      # write where Pages serves from
    python3 build_site.py --data other.yml

Dependencies: pyyaml, jinja2
    pip install -r requirements.txt
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import yaml
from jinja2 import Environment, FileSystemLoader, select_autoescape

BASE = Path(__file__).parent
TEMPLATES_DIR = BASE / "templates"
STATIC_DIR = BASE / "static"


def load_data(data_path: Path) -> dict:
    with data_path.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def render_html(data: dict) -> str:
    env = Environment(
        loader=FileSystemLoader(str(TEMPLATES_DIR)),
        autoescape=select_autoescape(["html", "xml"]),
        trim_blocks=True,
        lstrip_blocks=True,
    )
    template = env.get_template("cv.html.j2")
    return template.render(
        personal=data["personal"],
        en=data["en"],
        es=data["es"],
    )


def build(data_path: Path, out_dir: Path) -> None:
    data = load_data(data_path)
    out_dir.mkdir(parents=True, exist_ok=True)

    # index.html
    html = render_html(data)
    (out_dir / "index.html").write_text(html, encoding="utf-8")
    print(f"✅ HTML → {out_dir / 'index.html'}")

    # stylesheet
    css_src = STATIC_DIR / "style.css"
    if css_src.exists():
        shutil.copy2(css_src, out_dir / "style.css")
        print(f"🎨 CSS  → {out_dir / 'style.css'}")
    else:
        print(f"⚠️  Stylesheet not found: {css_src}", file=sys.stderr)

    # photo
    photo_name = data["personal"].get("photo", "photo.jpg")
    photo_src = (BASE / photo_name).resolve()
    if photo_src.exists():
        shutil.copy2(photo_src, out_dir / photo_src.name)
        print(f"🖼️  Photo → {out_dir / photo_src.name}")
    else:
        print(f"⚠️  Photo not found: {photo_src}", file=sys.stderr)

    # .nojekyll so GitHub Pages serves files verbatim
    (out_dir / ".nojekyll").write_text("", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(
        description="Build the CV website (index.html) from cv_data.yml."
    )
    parser.add_argument("--data", default="cv_data.yml",
                        help="YAML data file (default: cv_data.yml)")
    parser.add_argument("--out", default="site",
                        help="Output directory (default: site)")
    args = parser.parse_args()

    data_path = (BASE / args.data).resolve()
    out_dir = (BASE / args.out).resolve()
    build(data_path, out_dir)
    print(f"\n🌐 Open: {out_dir / 'index.html'}")


if __name__ == "__main__":
    main()

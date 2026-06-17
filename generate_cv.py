#!/usr/bin/env python3
"""
Generador del CV de Jaime Berdejo Sánchez.

Lee `cv_data.yml` y produce dos archivos .docx (uno en inglés, otro en
español) con la misma maquetación que el CV original — sin necesidad de
abrir Word ni Pages.

Uso:
    python3 generate_cv.py
    python3 generate_cv.py --pdf          # además exporta a PDF (requiere LibreOffice)
    python3 generate_cv.py --out OTRA/    # cambia carpeta de salida

Dependencias: python-docx, pyyaml
    pip install python-docx pyyaml
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

import yaml
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from docx.shared import Cm, Pt, RGBColor, Twips


# ---------------------------------------------------------------------
# Paleta / fuente del CV (extraída del DOCX original)
# ---------------------------------------------------------------------
FONT = "Calibri"
NAVY = "1F3864"        # color principal (títulos, nombre, links)
GRAY = "595959"        # color secundario (fecha, institución, contacto)

# Tamaños — se expresan en "half-points" (la unidad de OOXML)
SZ_NAME       = 40    # 20 pt
SZ_SUBTITLE   = 24    # 12 pt
SZ_CONTACT    = 20    # 10 pt
SZ_SECTION    = 24    # 12 pt
SZ_BODY       = 22    # 11 pt
SZ_DATE       = 20    # 10 pt
SZ_INST       = 21    # 10.5 pt
SZ_PROJ_STAT  = 21    # 10.5 pt
SZ_STACK      = 20    # 10 pt


# =====================================================================
# Helpers de bajo nivel (manipulación XML)
# =====================================================================
def _set(elem, key, value):
    elem.set(qn(key), str(value))


def _sub(parent, tag, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        _set(el, f"w:{k}", v)
    parent.append(el)
    return el


def _make_run_props(rPr, *, font=FONT, bold=False, italic=False,
                    color=None, size=None, spacing=None, underline=False):
    """Aplica rPr completos a un elemento <w:rPr>."""
    rFonts = _sub(rPr, "w:rFonts")
    for k in ("ascii", "cs", "eastAsia", "hAnsi"):
        _set(rFonts, f"w:{k}", font)
    if bold:
        _sub(rPr, "w:b")
        _sub(rPr, "w:bCs")
    if italic:
        _sub(rPr, "w:i")
        _sub(rPr, "w:iCs")
    if color:
        _sub(rPr, "w:color", val=color)
    if spacing is not None:
        _sub(rPr, "w:spacing", val=spacing)
    if size is not None:
        _sub(rPr, "w:sz", val=size)
        _sub(rPr, "w:szCs", val=size)
    if underline:
        _sub(rPr, "w:u", val="single")


def _add_run(paragraph, text, **kwargs):
    """Añade un <w:r> con su propio rPr al párrafo (XML directo)."""
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    _make_run_props(rPr, **kwargs)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    paragraph._p.append(r)
    return r


def _add_tab(paragraph):
    """Inserta un carácter de tabulación dentro del párrafo."""
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = _sub(rPr, "w:rFonts")
    for k in ("ascii", "cs", "eastAsia", "hAnsi"):
        _set(rFonts, f"w:{k}", FONT)
    r.append(rPr)
    tab = OxmlElement("w:tab")
    r.append(tab)
    paragraph._p.append(r)


def _set_paragraph_spacing(paragraph, *, before=None, after=None, line=None):
    pPr = paragraph._p.get_or_add_pPr()
    # Desactiva widow/orphan control para permitir cortes de página naturales
    for old in pPr.findall(qn("w:widowControl")):
        pPr.remove(old)
    _sub(pPr, "w:widowControl", val="0")
    # Eliminar spacing previo
    for old in pPr.findall(qn("w:spacing")):
        pPr.remove(old)
    spacing = _sub(pPr, "w:spacing")
    if before is not None:
        _set(spacing, "w:before", before)
    if after is not None:
        _set(spacing, "w:after", after)
    if line is not None:
        _set(spacing, "w:line", line)
        _set(spacing, "w:lineRule", "auto")


def _add_right_tab(paragraph, pos=9740):
    """Define una tabulación derecha para el párrafo (para alinear fechas)."""
    pPr = paragraph._p.get_or_add_pPr()
    tabs = _sub(pPr, "w:tabs")
    tab = _sub(tabs, "w:tab", val="right", pos=pos)


def _add_bottom_border(paragraph, color=NAVY, size=8, space=2):
    """Añade subrayado inferior al párrafo (para los encabezados de sección)."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = _sub(pPr, "w:pBdr")
    _sub(pBdr, "w:bottom", val="single", color=color, sz=size, space=space)


def _add_list_bullet(paragraph, *, indent=360):
    """Marca el párrafo como bullet (•) con sangría manual."""
    pPr = paragraph._p.get_or_add_pPr()
    ind = _sub(pPr, "w:ind", left=indent, hanging=indent)


def _keep_together(paragraph):
    """Evita que el párrafo se parta entre dos páginas (mantiene sus líneas juntas)."""
    pPr = paragraph._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:keepLines")):
        pPr.remove(old)
    _sub(pPr, "w:keepLines")


# =====================================================================
# Constructores de bloques del CV
# =====================================================================
def build_header(doc: Document, personal: dict, lang_data: dict):
    """Tabla superior: nombre + datos de contacto a la izquierda, foto a la derecha."""
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    # Anchos de columna (en twips/dxa)
    tblPr = table._tbl.tblPr
    tblW = _sub(tblPr, "w:tblW", type="dxa", w="10178")
    # Quitar bordes
    tblBorders = _sub(tblPr, "w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        _sub(tblBorders, f"w:{side}", val="none", color="FFFFFF", sz=0)
    # Reset grid (python-docx ya creó uno; lo sustituimos)
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        table._tbl.remove(grid)
    new_grid = OxmlElement("w:tblGrid")
    _sub(new_grid, "w:gridCol", w="8478")
    _sub(new_grid, "w:gridCol", w="1700")
    table._tbl.insert(list(table._tbl).index(tblPr) + 1, new_grid)

    cell_text, cell_photo = table.rows[0].cells

    # ---- Celda izquierda ----
    tcPr_left = cell_text._tc.get_or_add_tcPr()
    _sub(tcPr_left, "w:tcW", type="dxa", w="8478")
    tcMar = _sub(tcPr_left, "w:tcMar")
    for side, w in (("top", 0), ("left", 0), ("bottom", 0), ("right", 200)):
        _sub(tcMar, f"w:{side}", type="dxa", w=w)
    _sub(tcPr_left, "w:vAlign", val="center")

    # Limpiar el párrafo inicial creado por add_table
    cell_text.paragraphs[0]._p.getparent().remove(cell_text.paragraphs[0]._p)

    # Nombre
    p_name = cell_text.add_paragraph()
    _set_paragraph_spacing(p_name, after=60)
    _add_run(p_name, personal["name"], bold=True, color=NAVY, size=SZ_NAME)

    # Subtítulo (rol)
    p_role = cell_text.add_paragraph()
    _set_paragraph_spacing(p_role, after=100)
    _add_run(p_role, lang_data["title"], color=GRAY, size=SZ_SUBTITLE, spacing=10)

    # Localización
    p_loc = cell_text.add_paragraph()
    _set_paragraph_spacing(p_loc, after=20)
    _add_run(p_loc, lang_data["location"], color=GRAY, size=SZ_CONTACT)

    # Teléfono
    p_tel = cell_text.add_paragraph()
    _set_paragraph_spacing(p_tel, after=20)
    _add_run(p_tel, personal["phone"], color=GRAY, size=SZ_CONTACT)

    # Email (con hyperlink azul subrayado)
    p_mail = cell_text.add_paragraph()
    _set_paragraph_spacing(p_mail, after=0)
    _add_hyperlink(p_mail, personal["email"], f"mailto:{personal['email']}")

    # ---- Celda derecha (foto) ----
    tcPr_right = cell_photo._tc.get_or_add_tcPr()
    _sub(tcPr_right, "w:tcW", type="dxa", w="1700")
    tcMar_r = _sub(tcPr_right, "w:tcMar")
    for side in ("top", "left", "bottom", "right"):
        _sub(tcMar_r, f"w:{side}", type="dxa", w=0)
    _sub(tcPr_right, "w:vAlign", val="center")

    cell_photo.paragraphs[0]._p.getparent().remove(cell_photo.paragraphs[0]._p)
    p_photo = cell_photo.add_paragraph()
    _set_paragraph_spacing(p_photo, after=0)
    p_photo.alignment = 2  # right
    photo_path = (Path(__file__).parent / personal["photo"]).resolve()
    if photo_path.exists():
        run = p_photo.add_run()
        run.add_picture(str(photo_path), width=Cm(2.35))
    else:
        print(f"⚠️  Foto no encontrada: {photo_path}", file=sys.stderr)



def _add_hyperlink(paragraph, text, url):
    """Añade un hyperlink en estilo "email" (navy + subrayado)."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyper = OxmlElement("w:hyperlink")
    hyper.set(qn("r:id"), r_id)
    hyper.set(qn("w:history"), "1")
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    _make_run_props(rPr, color=NAVY, size=SZ_CONTACT, underline=True)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    hyper.append(r)
    paragraph._p.append(hyper)


def add_section_heading(doc: Document, text: str):
    p = doc.add_paragraph()
    _set_paragraph_spacing(p, before=200, after=10)
    # Mantener encabezado siempre junto a su primer párrafo (evita huérfanos)
    pPr = p._p.get_or_add_pPr()
    _sub(pPr, "w:keepNext")
    _add_run(p, text, bold=True, color=NAVY, size=SZ_SECTION, spacing=20)
    return p


def add_profile(doc, profile):
    add_section_heading(doc, profile["heading"])
    p = doc.add_paragraph()
    _set_paragraph_spacing(p, before=0, after=40, line=260)
    _add_run(p, profile["text"].strip(), size=SZ_BODY)


def _add_titled_entry(doc, title, dates):
    """Línea con título en negrita + tab derecho + fecha en cursiva gris."""
    p = doc.add_paragraph()
    _set_paragraph_spacing(p, before=40, after=10)
    _add_right_tab(p)
    _add_run(p, title, bold=True, size=SZ_BODY)
    _add_tab(p)
    _add_run(p, dates, italic=True, color=GRAY, size=SZ_DATE)
    return p


def _add_subtitle_line(doc, text):
    p = doc.add_paragraph()
    _set_paragraph_spacing(p, before=0, after=10, line=260)
    _add_run(p, text, italic=True, color=GRAY, size=SZ_INST)


def add_education(doc, edu):
    add_section_heading(doc, edu["heading"])
    for item in edu["items"]:
        _add_titled_entry(doc, item["title"], item["dates"])
        _add_subtitle_line(doc, item["institution"])


def add_experience(doc, exp):
    add_section_heading(doc, exp["heading"])
    for job in exp["jobs"]:
        # Línea de rol/empresa
        p = doc.add_paragraph()
        _set_paragraph_spacing(p, before=40, after=10)
        _add_right_tab(p)
        _add_run(p, job["role"], bold=True, size=SZ_BODY)
        _add_run(p, f" — {job['company']}", size=SZ_BODY)
        _add_tab(p)
        _add_run(p, job["dates"], italic=True, color=GRAY, size=SZ_DATE)
        # Bullets
        for bullet in job["bullets"]:
            bp = doc.add_paragraph()
            _set_paragraph_spacing(bp, before=0, after=10, line=240)
            _keep_together(bp)
            _add_list_bullet(bp)
            # bullet glyph
            _add_run(bp, "• ", size=SZ_BODY)
            _add_run(bp, bullet["bold"] + " ", bold=True, size=SZ_BODY)
            _add_run(bp, bullet["text"], size=SZ_BODY)


def add_skills(doc, skills):
    add_section_heading(doc, skills["heading"])
    for group in skills["groups"]:
        p = doc.add_paragraph()
        _set_paragraph_spacing(p, before=0, after=10, line=240)
        _keep_together(p)
        _add_run(p, f"{group['label']}:  ", bold=True, size=SZ_BODY)
        _add_run(p, group["text"], size=SZ_BODY)


def add_languages(doc, langs):
    add_section_heading(doc, langs["heading"])
    p = doc.add_paragraph()
    _set_paragraph_spacing(p, before=0, after=20, line=260)
    _keep_together(p)
    parts = []
    for i, item in enumerate(langs["items"]):
        if i > 0:
            _add_run(p, "    ", size=SZ_BODY)  # separación entre idiomas
        _add_run(p, f"{item['label']}: ", bold=True, size=SZ_BODY)
        _add_run(p, item["value"], size=SZ_BODY)


def add_projects(doc, projects):
    add_section_heading(doc, projects["heading"])
    for proj in projects["items"]:
        # Nombre + estado (en cursiva)
        p_name = doc.add_paragraph()
        _set_paragraph_spacing(p_name, before=80, after=20)
        _keep_together(p_name)
        _sub(p_name._p.get_or_add_pPr(), "w:keepNext")
        _add_run(p_name, proj["name"], bold=True, color=NAVY, size=SZ_BODY)
        if proj.get("status"):
            _add_run(
                p_name,
                f" — {proj['status']}",
                italic=True,
                color=GRAY,
                size=SZ_PROJ_STAT,
            )
        # Stack
        p_stack = doc.add_paragraph()
        _set_paragraph_spacing(p_stack, before=0, after=20, line=260)
        _keep_together(p_stack)
        _sub(p_stack._p.get_or_add_pPr(), "w:keepNext")
        _add_run(p_stack, "Stack: ", bold=True, size=SZ_STACK)
        _add_run(p_stack, proj["stack"], italic=True, color=GRAY, size=SZ_STACK)
        # Descripción
        p_desc = doc.add_paragraph()
        _set_paragraph_spacing(p_desc, before=0, after=40, line=260)
        _add_run(p_desc, proj["description"], size=SZ_BODY)


# =====================================================================
# Configuración global del documento (página, fuente por defecto)
# =====================================================================
def configure_document(doc: Document):
    # Márgenes de la página (864 twips ≈ 0.6 inch) — los del CV original
    section = doc.sections[0]
    section.top_margin    = Twips(864)
    section.bottom_margin = Twips(864)
    section.left_margin   = Twips(864)
    section.right_margin  = Twips(864)

    # Fuente por defecto: Calibri 11
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(11)


# =====================================================================
# Pipeline principal
# =====================================================================
def build_cv(data: dict, lang: str) -> Document:
    doc = Document()
    configure_document(doc)

    lang_data = data[lang]
    personal = data["personal"]

    build_header(doc, personal, lang_data)
    add_profile(doc, lang_data["profile"])
    add_education(doc, lang_data["education"])
    add_experience(doc, lang_data["experience"])
    add_skills(doc, lang_data["skills"])
    add_languages(doc, lang_data["languages"])
    add_projects(doc, lang_data["projects"])

    return doc


def _export_via_pages(docx_path: Path, also_pdf: bool = True) -> bool:
    """En macOS: abre el .docx en Pages, lo guarda como .pages y opcionalmente como .pdf.

    Devuelve True si se ejecutó correctamente, False en caso contrario.
    """
    if sys.platform != "darwin":
        return False
    pages_path = docx_path.with_suffix(".pages")
    pdf_path   = docx_path.with_suffix(".pdf")

    # Borrar destinos previos para evitar diálogos de sobrescritura
    for p in (pages_path, pdf_path if also_pdf else None):
        if p and p.exists():
            p.unlink()

    pdf_block = (
        f'export theDoc to POSIX file "{pdf_path}" as PDF\n'
        if also_pdf else ""
    )
    script = f'''
tell application "Pages"
    activate
    set theDoc to open POSIX file "{docx_path}"
    delay 1
    {pdf_block}save theDoc in POSIX file "{pages_path}"
    close theDoc saving no
end tell
'''
    try:
        subprocess.run(["osascript", "-e", script], check=True,
                       capture_output=True, text=True)
        print(f"📄 Pages → {pages_path.name}" + (f"   📑 PDF → {pdf_path.name}" if also_pdf else ""))
        return True
    except subprocess.CalledProcessError as e:
        print(f"⚠️  No pude exportar con Pages: {e.stderr}", file=sys.stderr)
        return False


def _export_pdf_libreoffice(docx_path: Path) -> bool:
    """Fallback: usar LibreOffice para exportar a PDF (no genera .pages)."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    out_dir = docx_path.parent
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir",
         str(out_dir), str(docx_path)],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    print(f"📑 PDF → {docx_path.with_suffix('.pdf').name} (vía LibreOffice)")
    return True


def export_companions(docx_path: Path, want_pages: bool, want_pdf: bool):
    """Genera los formatos acompañantes (.pages / .pdf) del .docx ya creado."""
    if want_pages and sys.platform == "darwin":
        # En macOS, Pages genera ambos en una sola sesión (más fiable).
        ok = _export_via_pages(docx_path, also_pdf=want_pdf)
        if ok:
            return
    # Sin .pages (otro SO) o falló Pages → usar LibreOffice solo para PDF
    if want_pdf:
        if not _export_pdf_libreoffice(docx_path):
            print(f"ℹ️  Ni Pages ni LibreOffice disponibles — no genero PDF.")


def main():
    parser = argparse.ArgumentParser(
        description="Genera el CV de Jaime en EN y ES (.docx, .pages, .pdf)."
    )
    parser.add_argument("--data", default="cv_data.yml",
                        help="Archivo YAML con los datos (default: cv_data.yml)")
    parser.add_argument("--out",  default=".",
                        help="Carpeta de salida (default: directorio actual)")
    parser.add_argument("--no-pages", action="store_true",
                        help="No generar .pages (por defecto se generan en macOS).")
    parser.add_argument("--no-pdf", action="store_true",
                        help="No generar .pdf (por defecto se generan).")
    args = parser.parse_args()

    base = Path(__file__).parent
    data_path = (base / args.data).resolve()
    out_dir   = (base / args.out).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    with data_path.open(encoding="utf-8") as f:
        data = yaml.safe_load(f)

    outputs = {
        "en": out_dir / "Jaime_Berdejo_CV.docx",
        "es": out_dir / "Jaime_Berdejo_CV_ES.docx",
    }

    want_pages = not args.no_pages
    want_pdf   = not args.no_pdf

    for lang, path in outputs.items():
        doc = build_cv(data, lang)
        doc.save(path)
        print(f"✅ {lang.upper()} → {path.name}")
        if want_pages or want_pdf:
            export_companions(path, want_pages=want_pages, want_pdf=want_pdf)


if __name__ == "__main__":
    main()

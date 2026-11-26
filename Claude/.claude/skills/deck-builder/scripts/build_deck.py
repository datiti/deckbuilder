#!/usr/bin/env python3
"""
build_deck.py : génère un deck PowerPoint (.pptx) à partir d'une spec JSON.

Usage :
  python build_deck.py spec.json -o sortie/deck.pptx
  python build_deck.py spec.json -o sortie/deck.pptx --template modele.potx
  python build_deck.py --list-layouts modele.potx
  python build_deck.py spec.json --lint-only

Dépendance unique : python-pptx (pip install python-pptx).
Le reste est en bibliothèque standard, pour tourner sur un poste verrouillé.
"""
import argparse
import json
import math
import os
import shutil
import sys
import tempfile
import zipfile

# Console Windows (cp1252) : éviter un plantage sur les caractères accentués ou spéciaux
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml.ns import qn
except ImportError:
    sys.stderr.write(
        "ERREUR : python-pptx n'est pas installé.\n"
        "  pip install python-pptx\n"
        "  (sur poste entreprise : pip install --index-url <miroir interne> python-pptx)\n"
    )
    sys.exit(2)

# ---------------------------------------------------------------- constantes
SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.6)
DEFAULT_THEME = {
    "primary": "1F3A5F",    # titres, bandeaux
    "accent": "00915A",     # chiffres clés, puces
    "text": "222222",
    "muted": "6B7280",
    "light": "F3F4F6",
    "font": "Arial",
    "confidentiality": "Interne",
}
WARNINGS = []


def rgb(hexstr):
    return RGBColor.from_string(hexstr.replace("#", "").upper())


# ---------------------------------------------------------------- template
def open_presentation(template):
    """Ouvre un modèle .pptx ou .potx et retire ses slides d'exemple."""
    if not template:
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        return prs, False
    path = template
    if template.lower().endswith((".potx", ".potm")):
        # python-pptx refuse les .potx : on réécrit le content-type en .pptx
        tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
        tmp.close()
        with zipfile.ZipFile(template) as zin, zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = data.replace(
                        b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
                        b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                    ).replace(
                        b"application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
                        b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                    )
                zout.writestr(item, data)
        path = tmp.name
    prs = Presentation(path)
    # suppression des slides existantes du modèle
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)
    return prs, True


def list_layouts(template):
    prs, _ = open_presentation(template)
    print(f"Format : {Emu(prs.slide_width).inches:.2f} x {Emu(prs.slide_height).inches:.2f} pouces")
    for i, layout in enumerate(prs.slide_layouts):
        phs = ", ".join(f"{ph.placeholder_format.idx}:{ph.placeholder_format.type}" for ph in layout.placeholders)
        print(f"[{i}] {layout.name}  ->  {phs}")


def pick_layout(prs, kind, use_template):
    layouts = list(prs.slide_layouts)
    if not use_template:
        return layouts[6]  # Blank du modèle par défaut
    names = {
        "title": ["title slide", "diapositive de titre", "titre", "cover", "couverture"],
        "content": ["title only", "titre seul", "titre uniquement"],
    }[kind]
    for layout in layouts:
        if layout.name.strip().lower() in names:
            return layout
    if kind == "title":
        return layouts[0]
    # à défaut : le layout avec un seul placeholder titre
    for layout in layouts:
        types = [str(ph.placeholder_format.type) for ph in layout.placeholders]
        if any("TITLE" in t for t in types) and len(types) <= 3:
            return layout
    return layouts[min(5, len(layouts) - 1)]


# ---------------------------------------------------------------- texte
class Ctx:
    def __init__(self, prs, theme, use_template):
        self.prs = prs
        self.t = {**DEFAULT_THEME, **(theme or {})}
        self.use_template = use_template
        self.W = prs.slide_width
        self.H = prs.slide_height
        self.page = 0


def style_run(run, ctx, size, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    if not ctx.use_template:
        run.font.name = ctx.t["font"]
    run.font.color.rgb = rgb(color or ctx.t["text"])


def set_bullet(p, ctx, level):
    pPr = p._p.get_or_add_pPr()
    indent = 228600 * (level + 1)  # 0,25 pouce par niveau
    pPr.set("marL", str(indent))
    pPr.set("indent", str(-228600))
    for tag in ("a:buNone", "a:buChar", "a:buClr"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    buClr = pPr.makeelement(qn("a:buClr"), {})
    srgb = buClr.makeelement(qn("a:srgbClr"), {"val": ctx.t["accent"]})
    buClr.append(srgb)
    pPr.append(buClr)
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•" if level == 0 else "–"}))


def estimate_overflow(where, paragraphs, box_w, box_h):
    """Heuristique : Claude Code ne voit pas le rendu, ce contrôle lui sert d'yeux."""
    if not where:
        return
    total_h = 0.0
    for text, size, indent in paragraphs:
        char_w = size * 0.48 / 72  # pouces, Arial moyen
        usable = max(Emu(int(box_w)).inches - 0.2 - indent, 0.5)
        n = max(1, math.ceil(len(text) * char_w / usable))
        total_h += n * size * 1.25 / 72 + size * 0.35 / 72
    if total_h > Emu(int(box_h)).inches + 0.05:
        WARNINGS.append(
            f"{where} : texte probablement trop long ({total_h:.1f} po nécessaires pour "
            f"{Emu(int(box_h)).inches:.1f} po disponibles). Raccourcir ou scinder la slide."
        )


def add_text(slide, ctx, x, y, w, h, items, size=16, color=None, bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, bullets=False, where=""):
    """items : str | list[str | list[str]] (liste imbriquée = sous-puces)."""
    x, y, w, h = (int(v) for v in (x, y, w, h))
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.08)
    if isinstance(items, str):
        items = [items]
    flat = []
    for it in items:
        if isinstance(it, list):
            flat += [(s, 1) for s in it]
        else:
            flat.append((it, 0))
    measure = []
    for i, (text, level) in enumerate(flat):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(size * 0.45)
        run = p.add_run()
        run.text = str(text)
        sz = size if level == 0 else size - 2
        style_run(run, ctx, sz, bold=bold, color=color)
        if bullets:
            set_bullet(p, ctx, level)
        measure.append((str(text), sz, 0.25 * (level + 1) if bullets else 0))
    estimate_overflow(where, measure, w, h)
    return tb


def add_box(slide, ctx, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    x, y, w, h = (int(v) for v in (x, y, w, h))
    shp = slide.shapes.add_shape(shape, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill)
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


# ---------------------------------------------------------------- cadre commun
def new_slide(ctx, spec, kind="content"):
    layout = pick_layout(ctx.prs, kind, ctx.use_template)
    slide = ctx.prs.slides.add_slide(layout)
    ctx.page += 1
    if spec.get("notes"):
        slide.notes_slide.notes_text_frame.text = spec["notes"]
    return slide


def frame(slide, ctx, spec):
    """Titre d'action + bandeau message + pied de page. Retourne la zone de contenu (x, y, w, h)."""
    where = f"Slide {ctx.page} ({spec.get('type')})"
    title = spec.get("title", "")
    top = Inches(0.3)
    if ctx.use_template and slide.shapes.title is not None:
        slide.shapes.title.text = title
        tph = slide.shapes.title
        content_top = tph.top + tph.height + Inches(0.15)
        # taille réelle du modèle inconnue : hypothèse prudente de 36 pt
        estimate_overflow(where + " titre (modèle)", [(title, 36, 0)], tph.width, tph.height)
    else:
        add_text(slide, ctx, MARGIN, top, ctx.W - 2 * MARGIN, Inches(1.0), title,
                 size=24, bold=True, color=ctx.t["primary"], anchor=MSO_ANCHOR.BOTTOM,
                 where=where + " titre")
        add_box(slide, ctx, MARGIN, top + Inches(1.08), Inches(1.2), Inches(0.05), ctx.t["accent"])
        content_top = top + Inches(1.35)
    if spec.get("message"):
        add_box(slide, ctx, MARGIN, content_top, ctx.W - 2 * MARGIN, Inches(0.6), ctx.t["light"])
        add_box(slide, ctx, MARGIN, content_top, Inches(0.08), Inches(0.6), ctx.t["accent"])
        add_text(slide, ctx, MARGIN + Inches(0.2), content_top, ctx.W - 2 * MARGIN - Inches(0.3),
                 Inches(0.6), spec["message"], size=15, bold=True, color=ctx.t["primary"],
                 anchor=MSO_ANCHOR.MIDDLE, where=where + " message")
        content_top += Inches(0.8)
    footer_y = ctx.H - Inches(0.45)
    if not ctx.use_template:
        foot = spec.get("source") and f"Source : {spec['source']}" or ""
        add_text(slide, ctx, MARGIN, footer_y, Inches(9), Inches(0.3), foot, size=9, color=ctx.t["muted"])
        add_text(slide, ctx, ctx.W - MARGIN - Inches(3), footer_y, Inches(3), Inches(0.3),
                 f"{ctx.t['confidentiality']}  |  {ctx.page}", size=9, color=ctx.t["muted"], align=PP_ALIGN.RIGHT)
    elif spec.get("source"):
        add_text(slide, ctx, MARGIN, footer_y - Inches(0.3), Inches(9), Inches(0.3),
                 f"Source : {spec['source']}", size=9, color=ctx.t["muted"])
    bottom = footer_y - Inches(0.15)
    return MARGIN, content_top, ctx.W - 2 * MARGIN, bottom - content_top


# ---------------------------------------------------------------- types de slides
def s_title(ctx, spec):
    slide = new_slide(ctx, spec, "title")
    if ctx.use_template and slide.shapes.title is not None:
        slide.shapes.title.text = spec.get("title", "")
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1:
                ph.text = "\n".join(x for x in [spec.get("subtitle", ""), spec.get("date", "")] if x)
        return
    add_box(slide, ctx, 0, 0, ctx.W, ctx.H, ctx.t["primary"])
    add_box(slide, ctx, MARGIN, Inches(3.55), Inches(1.5), Inches(0.07), ctx.t["accent"])
    add_text(slide, ctx, MARGIN, Inches(1.6), Inches(11), Inches(1.9), spec.get("title", ""),
             size=40, bold=True, color="FFFFFF", anchor=MSO_ANCHOR.BOTTOM, where="Slide titre")
    add_text(slide, ctx, MARGIN, Inches(3.8), Inches(11), Inches(1.2), spec.get("subtitle", ""),
             size=20, color="D1D5DB")
    add_text(slide, ctx, MARGIN, ctx.H - Inches(1.0), Inches(8), Inches(0.4),
             "  |  ".join(x for x in [spec.get("date", ""), ctx.t["confidentiality"]] if x), size=12, color="D1D5DB")


def s_section(ctx, spec):
    slide = new_slide(ctx, spec, "title" if ctx.use_template else "content")
    if ctx.use_template and slide.shapes.title is not None:
        slide.shapes.title.text = spec.get("title", "")
        return
    add_box(slide, ctx, 0, 0, Inches(0.35), ctx.H, ctx.t["accent"])
    add_text(slide, ctx, Inches(1.2), Inches(2.6), Inches(10.5), Inches(1.2), spec.get("title", ""),
             size=36, bold=True, color=ctx.t["primary"], anchor=MSO_ANCHOR.BOTTOM)
    add_text(slide, ctx, Inches(1.2), Inches(3.9), Inches(10.5), Inches(1), spec.get("subtitle", ""),
             size=18, color=ctx.t["muted"])


def s_bullets(ctx, spec):
    slide = new_slide(ctx, spec)
    x, y, w, h = frame(slide, ctx, spec)
    add_text(slide, ctx, x, y, w, h, spec.get("bullets", []), size=18, bullets=True,
             where=f"Slide {ctx.page} puces")


def s_two_columns(ctx, spec):
    slide = new_slide(ctx, spec)
    x, y, w, h = frame(slide, ctx, spec)
    gap = Inches(0.4)
    cw = (w - gap) / 2
    for i, key in enumerate(("left", "right")):
        col = spec.get(key, {})
        cx = x + i * (cw + gap)
        add_box(slide, ctx, cx, y, cw, Inches(0.55), ctx.t["primary"] if i == 0 else ctx.t["accent"])
        add_text(slide, ctx, cx + Inches(0.1), y, cw - Inches(0.2), Inches(0.55), col.get("heading", ""),
                 size=16, bold=True, color="FFFFFF", anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, ctx, cx, y + Inches(0.7), cw, h - Inches(0.7), col.get("bullets", []),
                 size=15, bullets=True, where=f"Slide {ctx.page} colonne {key}")


def s_kpis(ctx, spec):
    slide = new_slide(ctx, spec)
    x, y, w, h = frame(slide, ctx, spec)
    kpis = spec.get("kpis", [])[:4]
    n = max(len(kpis), 1)
    gap = Inches(0.3)
    kw = (w - gap * (n - 1)) / n
    kh = min(h, Inches(2.8))
    for i, k in enumerate(kpis):
        kx = x + i * (kw + gap)
        add_box(slide, ctx, kx, y, kw, kh, ctx.t["light"])
        add_box(slide, ctx, kx, y, kw, Inches(0.08), ctx.t["accent"])
        add_text(slide, ctx, kx, y + Inches(0.35), kw, Inches(1.1), k.get("value", ""), size=40, bold=True,
                 color=ctx.t["accent"], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 where=f"Slide {ctx.page} KPI {i+1}")
        add_text(slide, ctx, kx + Inches(0.15), y + Inches(1.5), kw - Inches(0.3), kh - Inches(1.6),
                 k.get("label", ""), size=14, color=ctx.t["text"], align=PP_ALIGN.CENTER,
                 where=f"Slide {ctx.page} KPI {i+1} libellé")
    if spec.get("comment"):
        add_text(slide, ctx, x, y + kh + Inches(0.3), w, h - kh - Inches(0.3), spec["comment"], size=14,
                 color=ctx.t["muted"], where=f"Slide {ctx.page} commentaire")


def s_table(ctx, spec):
    slide = new_slide(ctx, spec)
    x, y, w, h = frame(slide, ctx, spec)
    cols, rows = spec.get("columns", []), spec.get("rows", [])
    if not cols:
        return
    row_h = Inches(0.5)
    shape = slide.shapes.add_table(len(rows) + 1, len(cols), int(x), int(y), int(w), int(min(h, row_h * (len(rows) + 1))))
    tbl = shape.table
    widths = spec.get("col_widths")
    for j in range(len(cols)):
        tbl.columns[j].width = int(w * (widths[j] if widths else 1 / len(cols)))
    if row_h * (len(rows) + 1) > h:
        WARNINGS.append(f"Slide {ctx.page} tableau : {len(rows)} lignes, probablement trop pour la slide. Scinder.")
    for i in range(len(rows) + 1):
        for j in range(len(cols)):
            cell = tbl.cell(i, j)
            val = cols[j] if i == 0 else (rows[i - 1][j] if j < len(rows[i - 1]) else "")
            cell.text = ""
            run = cell.text_frame.paragraphs[0].add_run()
            run.text = str(val)
            style_run(run, ctx, 13 if i else 14, bold=(i == 0), color="FFFFFF" if i == 0 else ctx.t["text"])
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(ctx.t["primary"] if i == 0 else ("FFFFFF" if i % 2 else ctx.t["light"]))
            cell.margin_left = cell.margin_right = Inches(0.1)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE


def s_roadmap(ctx, spec):
    slide = new_slide(ctx, spec)
    x, y, w, h = frame(slide, ctx, spec)
    phases = spec.get("phases", [])
    n = max(len(phases), 1)
    gap = Inches(0.15)
    pw = (w - gap * (n - 1)) / n
    palette = [ctx.t["primary"], ctx.t["accent"], ctx.t["muted"], ctx.t["primary"], ctx.t["accent"]]
    for i, ph in enumerate(phases):
        px = x + i * (pw + gap)
        add_box(slide, ctx, px, y, pw, Inches(0.75), palette[i % len(palette)], shape=MSO_SHAPE.CHEVRON if n > 1 else MSO_SHAPE.RECTANGLE)
        add_text(slide, ctx, px + Inches(0.35), y, pw - Inches(0.6), Inches(0.75),
                 [ph.get("label", "")], size=14, bold=True, color="FFFFFF",
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, where=f"Slide {ctx.page} phase {i+1}")
        add_text(slide, ctx, px, y + Inches(0.85), pw, Inches(0.4), ph.get("period", ""), size=12,
                 bold=True, color=ctx.t["muted"], align=PP_ALIGN.CENTER)
        add_text(slide, ctx, px, y + Inches(1.3), pw, h - Inches(1.3), ph.get("items", []), size=13,
                 bullets=True, where=f"Slide {ctx.page} phase {i+1} contenu")


def s_matrix(ctx, spec):
    slide = new_slide(ctx, spec)
    x, y, w, h = frame(slide, ctx, spec)
    q = (spec.get("quadrants", []) + [{}] * 4)[:4]
    lab = Inches(0.45)
    gx, gy, gw, gh = x + lab, y, w - lab, h - lab
    cw, ch = gw / 2, gh / 2
    fills = [ctx.t["light"], "E6F4EE", ctx.t["light"], ctx.t["light"]]
    for i, quad in enumerate(q):
        cx = gx + (i % 2) * cw
        cy = gy + (i // 2) * ch
        add_box(slide, ctx, cx + Inches(0.05), cy + Inches(0.05), cw - Inches(0.1), ch - Inches(0.1),
                quad.get("fill", fills[i]).replace("#", ""))
        add_text(slide, ctx, cx + Inches(0.2), cy + Inches(0.15), cw - Inches(0.4), Inches(0.45),
                 quad.get("title", ""), size=15, bold=True, color=ctx.t["primary"])
        add_text(slide, ctx, cx + Inches(0.2), cy + Inches(0.6), cw - Inches(0.4), ch - Inches(0.75),
                 quad.get("items", []), size=14, bullets=True, where=f"Slide {ctx.page} quadrant {i+1}")
    add_text(slide, ctx, gx, y + gh + Inches(0.05), gw, Inches(0.35), spec.get("x_label", "") + "  →",
             size=12, bold=True, color=ctx.t["muted"], align=PP_ALIGN.CENTER)
    vt = slide.shapes.add_textbox(int(x + lab / 2 - gh / 2), int(y + gh / 2 - lab / 2), int(gh), int(lab))
    vt.rotation = 270
    p = vt.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = spec.get("y_label", "") + "  →"
    style_run(r, ctx, 12, bold=True, color=ctx.t["muted"])


BUILDERS = {
    "title": s_title, "section": s_section, "bullets": s_bullets,
    "two_columns": s_two_columns, "kpis": s_kpis, "table": s_table,
    "roadmap": s_roadmap, "matrix": s_matrix,
}


# ---------------------------------------------------------------- contrôles
def lint_spec(spec):
    errs = []
    slides = spec.get("slides", [])
    if not slides:
        errs.append("La spec ne contient aucune slide.")
    for i, s in enumerate(slides, 1):
        t = s.get("type")
        if t not in BUILDERS:
            errs.append(f"Slide {i} : type inconnu '{t}'. Types valides : {', '.join(BUILDERS)}")
        if t not in ("title", "section") and len(s.get("title", "")) > 90:
            WARNINGS.append(f"Slide {i} : titre de {len(s['title'])} caractères (viser < 90).")
        if t == "bullets" and len(s.get("bullets", [])) > 6:
            WARNINGS.append(f"Slide {i} : {len(s['bullets'])} puces de 1er niveau (viser 6 au maximum).")
        if t == "kpis" and len(s.get("kpis", [])) > 4:
            WARNINGS.append(f"Slide {i} : plus de 4 KPI, seuls les 4 premiers sont affichés.")
    return errs


def main():
    ap = argparse.ArgumentParser(description="Génère un .pptx depuis une spec JSON")
    ap.add_argument("spec", nargs="?")
    ap.add_argument("-o", "--output", default="sortie/deck.pptx")
    ap.add_argument("--template", help="modèle .pptx ou .potx de l'entreprise")
    ap.add_argument("--list-layouts", metavar="TEMPLATE")
    ap.add_argument("--lint-only", action="store_true")
    a = ap.parse_args()

    if a.list_layouts:
        list_layouts(a.list_layouts)
        return
    if not a.spec:
        ap.error("spec JSON manquante")
    with open(a.spec, encoding="utf-8") as f:
        spec = json.load(f)

    errs = lint_spec(spec)
    if errs:
        print("ERREURS :\n  " + "\n  ".join(errs))
        sys.exit(1)

    template = a.template or spec.get("template")
    prs, use_template = open_presentation(template)
    ctx = Ctx(prs, spec.get("theme"), use_template)
    for s in spec["slides"]:
        BUILDERS[s["type"]](ctx, s)

    if WARNINGS:
        print("AVERTISSEMENTS (à corriger avant livraison) :")
        for w in WARNINGS:
            print("  - " + w)
    if a.lint_only:
        print(f"Lint terminé : {len(spec['slides'])} slides, {len(WARNINGS)} avertissement(s).")
        sys.exit(1 if WARNINGS else 0)

    os.makedirs(os.path.dirname(os.path.abspath(a.output)), exist_ok=True)
    prs.core_properties.title = spec.get("title", "")
    prs.save(a.output)
    print(f"OK : {a.output} ({len(spec['slides'])} slides, {len(WARNINGS)} avertissement(s))")


if __name__ == "__main__":
    main()

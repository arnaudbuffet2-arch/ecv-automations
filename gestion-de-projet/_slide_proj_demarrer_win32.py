#!/usr/bin/env python3
"""
Slide PowerPoint via win32com : Projets à démarrer.
100% editable — icones SVG, textes et formes natifs PowerPoint.
"""
import sys, shutil, tempfile, urllib.request
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    print("pip install pywin32"); sys.exit(1)

def rgb(r, g, b): return r + g*256 + b*65536

NAVY  = rgb(0x1B, 0x2A, 0x4A)
WHITE = rgb(0xFF, 0xFF, 0xFF)
LGREY = rgb(0xCC, 0xD2, 0xDA)
STRIP = rgb(0xF0, 0xF2, 0xF5)
GOLD  = rgb(0xC4, 0x9C, 0x1A)
NGREY = rgb(0xD0, 0xD8, 0xE4)   # numéros filigrane

def pt(i): return i * 72

TDIR = Path(tempfile.mkdtemp(prefix="proj_"))

def _fetch(prefix, name, color):
    url = f"https://api.iconify.design/{prefix}/{name}.svg?color=%23{color}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            svg = r.read().decode()
            return svg if "viewBox" in svg else None
    except Exception as e:
        print(f"  [fetch] {prefix}:{name} -- {e}"); return None

LOGO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">
<g stroke="#1B2A4A" stroke-linecap="round" fill="none">
<line x1="24" y1="24" x2="24" y2="3"  stroke-width="2.2"/>
<line x1="24" y1="24" x2="24" y2="45" stroke-width="2.2"/>
<line x1="24" y1="24" x2="3"  y2="24" stroke-width="2.2"/>
<line x1="24" y1="24" x2="45" y2="24" stroke-width="2.2"/>
<line x1="24" y1="24" x2="10" y2="10" stroke-width="1.8"/>
<line x1="24" y1="24" x2="38" y2="38" stroke-width="1.8"/>
<line x1="24" y1="24" x2="38" y2="10" stroke-width="1.8"/>
<line x1="24" y1="24" x2="10" y2="38" stroke-width="1.8"/>
<line x1="24" y1="4"  x2="21" y2="8"  stroke-width="1.4"/>
<line x1="24" y1="4"  x2="27" y2="8"  stroke-width="1.4"/>
<line x1="24" y1="44" x2="21" y2="40" stroke-width="1.4"/>
<line x1="24" y1="44" x2="27" y2="40" stroke-width="1.4"/>
<line x1="4"  y1="24" x2="8"  y2="21" stroke-width="1.4"/>
<line x1="4"  y1="24" x2="8"  y2="27" stroke-width="1.4"/>
<line x1="44" y1="24" x2="40" y2="21" stroke-width="1.4"/>
<line x1="44" y1="24" x2="40" y2="27" stroke-width="1.4"/>
</g></svg>"""

SIGMA_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
<path fill="white" d="M19 3H5v2.5l6.5 6.5L5 18.5V21h14v-3H9.5l5-5.5-5-5.5H19V3z"/>
</svg>"""

print("Chargement des icones...")
SVGS = {
    "logo":     LOGO_SVG,
    "chart":    _fetch("fluent", "data-bar-vertical-ascending-24-regular", "ffffff"),
    "link":     _fetch("fluent", "link-24-regular",              "ffffff"),
    "wallet":   _fetch("fluent", "wallet-24-regular",            "ffffff"),
    "calendar": _fetch("fluent", "calendar-24-regular",          "ffffff"),
    "sigma":    SIGMA_SVG,
    "send":     _fetch("fluent", "send-24-regular",              "ffffff"),
}

PATHS = {}
for name, svg in SVGS.items():
    if svg:
        p = TDIR / f"{name}.svg"
        p.write_text(svg, encoding="utf-8")
        PATHS[name] = str(p)
print(f"  {len(PATHS)}/{len(SVGS)} icones chargees")

# Fermer uniquement le fichier de sortie si déjà ouvert (évite le verrou)
try:
    _ex = win32.GetActiveObject("PowerPoint.Application")
    for _i in range(_ex.Presentations.Count, 0, -1):
        _p = _ex.Presentations.Item(_i)
        if "slide_proj_demarrer" in _p.Name:
            _p.Saved = True; _p.Close()
except Exception: pass

import time
ppt = win32.Dispatch("PowerPoint.Application")
ppt.Visible = True
time.sleep(0.5)   # laisse PowerPoint se stabiliser
prs = ppt.Presentations.Add(WithWindow=True)
time.sleep(0.5)
prs.PageSetup.SlideWidth  = pt(13.33)
prs.PageSetup.SlideHeight = pt(7.50)
sl = prs.Slides.Add(1, 12)
print(f"Slide créée, {sl.Shapes.Count} formes initiales")

# ── HELPERS ───────────────────────────────────────────────────────────────────

def Rect(x, y, w, h, fill, lc=None, lw=0.5, alpha=0.0):
    s = sl.Shapes.AddShape(1, pt(x), pt(y), pt(w), pt(h))
    s.Fill.Solid(); s.Fill.ForeColor.RGB = fill; s.Fill.Transparency = alpha
    if lc: s.Line.ForeColor.RGB = lc; s.Line.Weight = lw
    else:  s.Line.Visible = False
    return s

def RRect(x, y, w, h, fill, lc=None, lw=0.5, alpha=0.0):
    s = sl.Shapes.AddShape(5, pt(x), pt(y), pt(w), pt(h))
    s.Fill.Solid(); s.Fill.ForeColor.RGB = fill; s.Fill.Transparency = alpha
    if lc: s.Line.ForeColor.RGB = lc; s.Line.Weight = lw
    else:  s.Line.Visible = False
    return s

def Oval(x, y, d, fill=None, lc=None, lw=1.5):
    s = sl.Shapes.AddShape(9, pt(x), pt(y), pt(d), pt(d))
    if fill is not None:
        s.Fill.Solid(); s.Fill.ForeColor.RGB = fill
    else:
        s.Fill.Visible = False
    if lc: s.Line.ForeColor.RGB = lc; s.Line.Weight = lw
    else:  s.Line.Visible = False
    return s

def Line(x1, y, x2, color=None, lw=0.5):
    if color is None: color = LGREY
    s = sl.Shapes.AddLine(pt(x1), pt(y), pt(x2), pt(y))
    s.Line.ForeColor.RGB = color; s.Line.Weight = lw
    return s

def T(x, y, w, h, text, size, bold=False, italic=False, color=None,
      align=1, font="Calibri", valign=1, spacing=0, lspace=1.0):
    if color is None: color = NAVY
    tb = sl.Shapes.AddTextbox(1, pt(x), pt(y), pt(w), pt(h))
    tf = tb.TextFrame
    tf.WordWrap = True; tf.AutoSize = 0
    tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
    try: tf.VerticalAnchor = valign
    except Exception: pass
    tr = tf.TextRange
    tr.Text = text.replace('\n', '\r')
    f = tr.Font
    f.Size = size; f.Bold = bold; f.Italic = italic
    f.Color.RGB = color; f.Name = font
    if spacing != 0:
        try: f.Spacing = spacing
        except Exception: pass
    pf = tr.ParagraphFormat
    pf.SpaceBefore = 0; pf.SpaceAfter = 0; pf.SpaceWithin = lspace
    pf.Alignment = align
    tb.Fill.Visible = False; tb.Line.Visible = False
    return tb

def Shadow(shape, offset=1.5, blur=4, transparency=0.80):
    sh = shape.Shadow
    sh.Visible = -1; sh.OffsetX = offset; sh.OffsetY = offset
    sh.Blur = blur; sh.Transparency = transparency

def Svg(key, x, y, w, h):
    if key not in PATHS: return None
    try:
        s = sl.Shapes.AddPicture(PATHS[key], False, True, pt(x), pt(y), pt(w), pt(h))
        s.Width = pt(w); s.Height = pt(h); s.Left = pt(x); s.Top = pt(y)
        return s
    except Exception as e:
        print(f"  SVG {key}: {e}"); return None

def SvgC(key, cx, cy, size):
    return Svg(key, cx - size/2, cy - size/2, size, size)

# ── FOND + LIGNE DE TÊTE ──────────────────────────────────────────────────────
Rect(0, 0, 13.33, 7.50, WHITE)
Line(0, 0.015, 13.33, NAVY, 2.5)

# ── LOGO ──────────────────────────────────────────────────────────────────────
Svg("logo", 12.50, 0.08, 0.72, 0.72)

# ── TITRE + ACCENT DORÉ + SOUS-TITRE ─────────────────────────────────────────
T(0.33, 0.17, 9.50, 0.82, "PROJETS À DÉMARRER",
  40, bold=True, font="Georgia")
Line(0.33, 1.06, 0.70, GOLD, 2.5)
T(0.33, 1.12, 7.50, 0.55,
  "Les initiatives prioritaires à lancer pour créer de la valeur\n"
  "et soutenir nos objectifs stratégiques.",
  10, lspace=1.35)

# ── GRILLE CARTES 3×2 ─────────────────────────────────────────────────────────
CW, CH = 4.10, 2.59
CX = [0.33, 0.33 + CW + 0.18, 0.33 + 2*(CW + 0.18)]   # [0.33, 4.61, 8.89]
CY = [1.80, 1.80 + CH + 0.20]                            # [1.80, 4.59]

CD, IS = 0.62, 0.38   # diamètre cercle, taille icone

PROJECTS = [
    ("01", "chart",    "Amélioration\ndu reporting",
     "Optimisation et automatisation\ndes reportings et consolidations."),
    ("02", "link",     "Interconnectivité\nentre outils",
     "Fluidification des échanges\nde données entre applications."),
    ("03", "wallet",   "Projet Frais\nBelgique &\nLuxembourg",
     "Harmonisation et simplification\nde la gestion des frais."),
    ("04", "calendar", "Outil de\nplanification",
     "Meilleure visibilité des charges\net des capacités."),
    ("05", "sigma",    "Étude\nmathématique\ndigitalisée",
     "Digitalisation et automatisation\ndes calculs."),
    ("06", "send",     "Invitations",
     "Modernisation du processus\nde gestion des invitations."),
]

for i, (num, icon, title, body) in enumerate(PROJECTS):
    cx, cy = CX[i % 3], CY[i // 3]

    # Fond carte (numéro d'abord → passe derrière le cercle automatiquement)
    card = RRect(cx, cy, CW, CH, STRIP, lc=LGREY, lw=0.4)
    Shadow(card, offset=1.5, blur=5, transparency=0.84)

    # Numéro filigrane
    T(cx + 0.10, cy + 0.06, 0.75, 0.72, num,
      50, bold=False, color=NGREY, font="Georgia")

    # Cercle navy + icône blanche centrée
    cl, ct = cx + 0.70, cy + 0.12
    Oval(cl, ct, CD, fill=NAVY)
    SvgC(icon, cl + CD/2, ct + CD/2, IS)

    # Titre
    T(cx + 1.44, cy + 0.12, CW - 1.52, 0.88, title,
      12, bold=True, lspace=1.2)

    # Tiret doré séparateur
    Line(cx + 0.18, cy + 1.10, cx + 0.46, GOLD, 1.5)

    # Corps de texte
    T(cx + 0.18, cy + 1.20, CW - 0.28, CH - 1.30, body,
      9, lspace=1.3)

# ── FOOTER ────────────────────────────────────────────────────────────────────
FY = 7.20
T(0.33, FY, 12.67, 0.26,
  "EXIGENCE   |   INDÉPENDANCE   |   ESPRIT DE CONQUÊTE",
  8, color=GOLD, align=2, spacing=2)
Line(0.40, FY + 0.09, 3.78, GOLD, 0.8)
Line(9.55, FY + 0.09, 12.93, GOLD, 0.8)

# ── SAUVEGARDE + PREVIEW PNG ──────────────────────────────────────────────────
OUT = r"C:\Users\arnau\OneDrive\Bureau\slide_proj_demarrer.pptx"
prs.SaveAs(OUT)
print(f"Sauvegarde : {OUT}")
try:
    sl.Export(OUT.replace(".pptx", "_preview.png"), "PNG")
    print(f"Preview    : {OUT.replace('.pptx', '_preview.png')}")
except Exception as e:
    print(f"  Preview ignoré : {e}")

shutil.rmtree(TDIR, ignore_errors=True)

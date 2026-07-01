#!/usr/bin/env python3
"""
Slide PowerPoint via win32com : Statut du portefeuille projets.
100% editable — icones SVG, textes et formes natifs PowerPoint.
"""
import sys, shutil, tempfile, urllib.request
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    print("pip install pywin32"); sys.exit(1)

# ── COULEURS ──────────────────────────────────────────────────────────────────
def rgb(r, g, b): return r + g*256 + b*65536

NAVY   = rgb(0x1B, 0x2A, 0x4A)
WHITE  = rgb(0xFF, 0xFF, 0xFF)
LGREY  = rgb(0xCC, 0xD2, 0xDA)
MGREY  = rgb(0x8A, 0x9B, 0xB5)
STRIP  = rgb(0xF0, 0xF2, 0xF5)
GREEN  = rgb(0x1A, 0x4A, 0x2C)
AMBER  = rgb(0xBF, 0x8A, 0x10)
BURGND = rgb(0x8B, 0x1A, 0x26)

# ── UNITÉS + CALIBRAGE PIXEL ──────────────────────────────────────────────────
def pt(i): return i * 72

IMG_W, IMG_H = 1265, 712          # dimensions pixel de l'image source
def from_px(px): return px * 13.33 / IMG_W
def from_py(py): return py *  7.50 / IMG_H

# ── ICONES SVG via Iconify (FluentUI, sans clé API) ──────────────────────────
TDIR = Path(tempfile.mkdtemp(prefix="ath_"))

def _fetch(prefix, name, color):
    # User-Agent obligatoire, Iconify refuse sans (403)
    url = f"https://api.iconify.design/{prefix}/{name}.svg?color=%23{color}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            svg = r.read().decode()
            return svg if "viewBox" in svg else None
    except Exception as e:
        print(f"  [fetch] {prefix}:{name} -- {e}"); return None

print("Chargement des icones...")
SVGS = {
    "logo": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">
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
</g>
</svg>""",
    "info":   _fetch("fluent", "info-24-regular",             "1B2A4A"),
    "check":  _fetch("fluent", "checkmark-circle-24-regular", "1A4A2C"),
    "flask":  _fetch("fluent", "beaker-24-regular",           "1B2A4A"),
    "gear":   _fetch("fluent", "settings-24-regular",         "BF8A10"),
    "rocket": _fetch("fluent", "rocket-24-regular",           "8B1A26"),
}

PATHS = {}
for name, content in SVGS.items():
    if content:
        p = TDIR / f"{name}.svg"
        p.write_text(content, encoding="utf-8")
        PATHS[name] = str(p)
print(f"  {len(PATHS)}/{len(SVGS)} icones chargees")

# ── POWERPOINT — fermer l'ancien fichier (évite le verrou) ───────────────────
try:
    _ex = win32.GetActiveObject("PowerPoint.Application")
    for _i in range(_ex.Presentations.Count, 0, -1):
        _p = _ex.Presentations.Item(_i)
        if "slide_ath_win32" in _p.Name:
            _p.Saved = True; _p.Close()
except Exception:
    pass

ppt = win32.Dispatch("PowerPoint.Application")
ppt.Visible = True
prs = ppt.Presentations.Add(WithWindow=True)
prs.PageSetup.SlideWidth  = pt(13.33)
prs.PageSetup.SlideHeight = pt(7.50)
sl = prs.Slides.Add(1, 12)

# ── HELPERS FORMES ────────────────────────────────────────────────────────────

def Rect(x, y, w, h, fill, lc=None, lw=0, alpha=0.0):
    s = sl.Shapes.AddShape(1, pt(x), pt(y), pt(w), pt(h))
    s.Fill.Solid()
    s.Fill.ForeColor.RGB = fill
    s.Fill.Transparency  = alpha
    if lc: s.Line.ForeColor.RGB = lc; s.Line.Weight = lw
    else:  s.Line.Visible = False
    return s

def RRect(x, y, w, h, fill, lc=None, lw=0, alpha=0.0):
    s = sl.Shapes.AddShape(5, pt(x), pt(y), pt(w), pt(h))
    s.Fill.Solid()
    s.Fill.ForeColor.RGB = fill
    s.Fill.Transparency  = alpha
    if lc: s.Line.ForeColor.RGB = lc; s.Line.Weight = lw
    else:  s.Line.Visible = False
    return s

def Oval(x, y, d, fill=None, lc=None, lw=1.5, alpha=0.0):
    s = sl.Shapes.AddShape(9, pt(x), pt(y), pt(d), pt(d))
    if fill is not None:
        s.Fill.Solid()
        s.Fill.ForeColor.RGB = fill
        s.Fill.Transparency  = alpha
    else:
        s.Fill.Visible = False
    if lc: s.Line.ForeColor.RGB = lc; s.Line.Weight = lw
    else:  s.Line.Visible = False
    return s

def Line(x1, y, x2, color=None, lw=0.5):
    if color is None: color = LGREY
    s = sl.Shapes.AddLine(pt(x1), pt(y), pt(x2), pt(y))
    s.Line.ForeColor.RGB = color
    s.Line.Weight        = lw
    return s

def T(x, y, w, h, text, size,
      bold=False, italic=False, color=None,
      align=1, font="Calibri",
      valign=1,    # 1=haut  3=milieu  4=bas
      spacing=0,   # tracking caractères (try/except — non exposé dans ce dispatch)
      lspace=1.0): # interligne en lignes : 1.0=simple, 1.5=1,5 ligne, -14=exactement 14pt
    if color is None: color = NAVY
    tb = sl.Shapes.AddTextbox(1, pt(x), pt(y), pt(w), pt(h))
    tf = tb.TextFrame
    tf.WordWrap       = True
    tf.AutoSize       = 0
    tf.MarginLeft   = 0     # supprimer padding interne par défaut (~7pt)
    tf.MarginRight  = 0
    tf.MarginTop    = 0
    tf.MarginBottom = 0
    try: tf.VerticalAnchor = valign
    except Exception: pass   # read-only dans certaines versions COM
    tr = tf.TextRange
    tr.Text = text.replace('\n', '\r')
    f = tr.Font        # cacher l'objet Font — ne pas chaîner les accès
    f.Size      = size
    f.Bold      = bold
    f.Italic    = italic
    f.Color.RGB = color
    f.Name      = font
    if spacing != 0:
        try: f.Spacing = spacing
        except Exception: pass  # non exposé dans certaines versions COM
    pf = tr.ParagraphFormat
    pf.SpaceBefore = 0
    pf.SpaceAfter  = 0
    pf.SpaceWithin = lspace
    pf.Alignment   = align
    tb.Fill.Visible = False
    tb.Line.Visible = False
    return tb

def Shadow(shape, offset=1.5, blur=4, transparency=0.80):
    sh = shape.Shadow
    sh.Visible      = -1    # msoTrue
    sh.OffsetX      = offset
    sh.OffsetY      = offset
    sh.Blur         = blur
    sh.Transparency = transparency

def ZOrder(shape, z):
    """0=SendToBack 1=BringToFront 2=SendBackward 3=BringForward
    ATTENTION: SendToBack(0) envoie derrière TOUS les éléments, y compris le fond blanc.
    Utiliser BringForward(3) de préférence."""
    shape.ZOrder(z)

def Svg(key, x, y, w, h):
    """Insère un SVG et force les dimensions exactes via COM (évite les décalages d'AddPicture)."""
    if key not in PATHS: return None
    try:
        s = sl.Shapes.AddPicture(PATHS[key], False, True, pt(x), pt(y), pt(w), pt(h))
        # Forcer les dimensions exactes — AddPicture peut les ajuster selon le viewBox
        s.Width  = pt(w)
        s.Height = pt(h)
        s.Left   = pt(x)
        s.Top    = pt(y)
        return s
    except Exception as e:
        print(f"  Avertissement SVG {key}: {e}"); return None

def SvgCentered(key, cx, cy, size):
    """Insère un SVG centré sur (cx, cy) avec une taille donnée."""
    return Svg(key, cx - size/2, cy - size/2, size, size)

def MeasureH(x, y, w, text, size, **kw):
    """Mesure la hauteur réelle occupée par un texte (AutoSize), puis supprime la forme."""
    tb = T(x, y, w, 9.0, text, size, **kw)
    tb.TextFrame.AutoSize = 1   # ppAutoSizeShapeToFitText
    h = tb.Height / 72          # points → pouces
    tb.Delete()
    return h

# ── FOND + LIGNE HAUT ─────────────────────────────────────────────────────────
Rect(0, 0, 13.33, 7.50, WHITE)
Line(0, 0.02, 13.33, NAVY, 3)

# ── TITRE + LOGO ──────────────────────────────────────────────────────────────
T(0.33, 0.18, 9.80, 0.70, "Statut du portefeuille projets",
  30, font="Calibri Light", valign=3)
Svg("logo", 12.28, 0.07, 0.77, 0.77)

# ── BOÎTE INFO ────────────────────────────────────────────────────────────────
INFO_Y = 1.06
INFO_H = 0.68
RRect(0.33, INFO_Y, 12.67, INFO_H, STRIP)
# Icone "i" centrée verticalement dans la boîte
SvgCentered("info", 0.33 + 0.33, INFO_Y + INFO_H / 2, 0.48)
T(1.11, INFO_Y, 11.60, INFO_H,
  "Chaque projet du portefeuille est associé à l'un des quatre statuts ci-dessous, "
  "permettant un suivi clair et homogène de l'avancement global.",
  9.5, valign=3)

# ── 4 CARTES STATUTS ─────────────────────────────────────────────────────────
CW  = 6.05   # largeur carte
CH  = 2.42   # hauteur carte
BW  = 0.07   # barre de bord gauche
CD  = 0.72   # diamètre cercle icone
IS  = CD * 0.62   # taille icone = 62% du cercle (≈ 0.446")
CX1 = 0.33
CX2 = 6.95
CY1 = 1.94
CY2 = CY1 + CH + 0.20

TX_OFF = BW + 0.15 + CD + 0.18   # offset colonne texte depuis bord gauche
TW     = CW - TX_OFF - 0.08      # largeur colonne texte

CARDS = [
    (CX1, CY1, GREEN,  "check",
     rgb(0xE6, 0xF2, 0xE8), "TERMINÉ",    GREEN,
     "Bénéfices sécurisés",
     "Aucune action complémentaire\nrequise."),

    (CX2, CY1, NAVY,   "flask",
     rgb(0xE8, 0xEC, 0xF5), "LIVRÉ",      NAVY,
     "Validation pilote",
     "Solution livrée, phase de test\nou de confirmation en conditions\n"
     "réelles avant clôture ou\ndéploiement."),

    (CX1, CY2, AMBER,  "gear",
     rgb(0xF5, 0xEE, 0xDE), "EN COURS",   AMBER,
     "Exécution active",
     "Projet actif avec actions,\njalons et livrables en cours\nde réalisation."),

    (CX2, CY2, BURGND, "rocket",
     rgb(0xF5, 0xE8, 0xEA), "À DÉMARRER", BURGND,
     "Cadrage requis",
     "Projet non lancé, nécessitant\ncadrage et initialisation avant\n"
     "toute mise en œuvre."),
]

for cx, cy, col, icon, cir_bg, title, sub_col, subtitle, body in CARDS:
    # Ordre de création = z-order automatique (fond → éléments visuels → textes)
    card_bg = Rect(cx + BW, cy, CW - BW, CH, STRIP)
    Shadow(card_bg, offset=1.5, blur=5, transparency=0.82)
    Rect(cx, cy, BW, CH, col)   # barre colorée gauche

    # Cercle + icône centrés précisément
    ix = cx + BW + 0.15          # left du cercle
    iy = cy + (CH - CD) / 2 + 0.10  # top du cercle (légèrement sous le centre)
    cx_circle = ix + CD / 2      # centre X du cercle
    cy_circle = iy + CD / 2      # centre Y du cercle
    Oval(ix, iy, CD, fill=cir_bg)
    SvgCentered(icon, cx_circle, cy_circle, IS)  # icône centrée sur le cercle

    # Colonne texte
    tx = cx + TX_OFF
    Line(tx, cy + 0.22, tx + 0.22, col, 1.5)    # tiret couleur statut
    T(tx, cy + 0.29, TW, 0.27, title,    10.5, bold=True)
    T(tx, cy + 0.57, TW, 0.24, subtitle,  9.0, color=sub_col)
    Line(tx, cy + 0.83, tx + 0.22, LGREY, 1.0)  # séparateur léger
    T(tx, cy + 0.92, TW, 1.36, body, 8.5, lspace=1.2)

# ── FOOTER ────────────────────────────────────────────────────────────────────
Line(0.33, 7.22, 6.44, LGREY, 0.5)
Line(6.89, 7.22, 13.00, LGREY, 0.5)
T(0.33, 7.26, 12.67, 0.20, "Interne", 8.0, color=MGREY, align=2)
T(12.87, 7.26, 0.35, 0.20, "3", 8.0, color=MGREY, align=3)

# ── SAUVEGARDE + PREVIEW PNG (export natif PowerPoint) ───────────────────────
OUT = r"C:\Users\arnau\OneDrive\Bureau\slide_ath_win32.pptx"
prs.SaveAs(OUT)
print(f"Sauvegarde : {OUT}")

try:
    PNG = OUT.replace(".pptx", "_preview.png")
    sl.Export(PNG, "PNG")
    print(f"Preview    : {PNG}")
except Exception as e:
    print(f"  Preview PNG ignoré : {e}")

shutil.rmtree(TDIR, ignore_errors=True)

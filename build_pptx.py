"""
Build Hindi_Coreference_Final_Presentation.pptx
Professional PPTX using pure Python (zipfile + XML).
Design: Dark navy theme (#1A2E4A bg, white text, #00C9B1 teal accent, #F5A623 gold accent)
"""

import zipfile, os, textwrap
from io import BytesIO

# ─────────────────────────────────────────────
# EMU helpers  (1 inch = 914400 EMU)
# ─────────────────────────────────────────────
SLIDE_W = 12192000   # ~13.33 in  (widescreen 16:9)
SLIDE_H =  6858000   # ~7.5 in

def emu(inches): return int(inches * 914400)

# ─────────────────────────────────────────────
# COLOUR palette
# ─────────────────────────────────────────────
C_BG      = "1A2E4A"   # deep navy background
C_BG2     = "112238"   # darker navy (header bars)
C_ACCENT  = "00C9B1"   # teal accent
C_GOLD    = "F5A623"   # gold accent
C_WHITE   = "FFFFFF"
C_LIGHT   = "D0E4F7"   # light blue-white
C_MUTED   = "8AAEC8"   # muted blue
C_DARK    = "0D1B2E"   # very dark for shapes
C_RED     = "E05C5C"   # warning / Phase-1 failed

# ─────────────────────────────────────────────
# Relationship / content-type XML templates
# ─────────────────────────────────────────────
CT_MAIN = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  {slide_overrides}
</Types>"""

RELS_ROOT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="ppt/presentation.xml"/>
</Relationships>"""

PPT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="slideMasters/slideMaster1.xml"/>
  {slide_rels}
</Relationships>"""

SLIDE_MASTER_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    Target="../theme/theme1.xml"/>
</Relationships>"""

SLIDE_LAYOUT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

THEME_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="HindiCorefTheme">
  <a:themeElements>
    <a:clrScheme name="HindiCoref">
      <a:dk1><a:srgbClr val="{C_DARK}"/></a:dk1>
      <a:lt1><a:srgbClr val="{C_WHITE}"/></a:lt1>
      <a:dk2><a:srgbClr val="{C_BG}"/></a:dk2>
      <a:lt2><a:srgbClr val="{C_LIGHT}"/></a:lt2>
      <a:accent1><a:srgbClr val="{C_ACCENT}"/></a:accent1>
      <a:accent2><a:srgbClr val="{C_GOLD}"/></a:accent2>
      <a:accent3><a:srgbClr val="4472C4"/></a:accent3>
      <a:accent4><a:srgbClr val="ED7D31"/></a:accent4>
      <a:accent5><a:srgbClr val="A9D18E"/></a:accent5>
      <a:accent6><a:srgbClr val="FF0000"/></a:accent6>
      <a:hlink><a:srgbClr val="{C_ACCENT}"/></a:hlink>
      <a:folHlink><a:srgbClr val="{C_MUTED}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="HindiCoref">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"""

SLIDE_MASTER_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="{C_BG}"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr algn="l">
        <a:defRPr sz="3200" b="1" lang="en-US">
          <a:solidFill><a:srgbClr val="{C_WHITE}"/></a:solidFill>
          <a:latin typeface="Calibri"/>
        </a:defRPr>
      </a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr>
        <a:defRPr sz="1800" lang="en-US">
          <a:solidFill><a:srgbClr val="{C_LIGHT}"/></a:solidFill>
          <a:latin typeface="Calibri"/>
        </a:defRPr>
      </a:lvl1pPr>
    </p:bodyStyle>
  </p:txStyles>
</p:sldMaster>"""

SLIDE_LAYOUT_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="{C_BG}"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""


# ─────────────────────────────────────────────
# XML helpers
# ─────────────────────────────────────────────

def solid_fill(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'

def rect_sp(sp_id, name, x, y, cx, cy, fill_color, border_color=None, border_w=0):
    """Solid-filled rectangle shape"""
    border = ""
    if border_color and border_w:
        border = f'<a:ln w="{border_w}"><a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill></a:ln>'
    else:
        border = '<a:ln><a:noFill/></a:ln>'
    return f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="{name}"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    {solid_fill(fill_color)}
    {border}
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>"""

def rounded_rect_sp(sp_id, name, x, y, cx, cy, fill_color, border_color=None, border_w=0, adj=16667):
    """Rounded rectangle shape"""
    border = f'<a:ln w="{border_w}"><a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill></a:ln>' if border_color and border_w else '<a:ln><a:noFill/></a:ln>'
    return f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="{name}"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst></a:prstGeom>
    {solid_fill(fill_color)}
    {border}
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>"""

def txbox(sp_id, name, x, y, cx, cy, paragraphs,
          font_size=1800, bold=False, color=C_WHITE,
          align="l", v_anchor="t", wrap=True,
          auto_fit=True, line_spacing=None, space_before=None):
    """
    paragraphs: list of (text, font_size, bold, color, align) tuples
                OR list of strings (use defaults)
    """
    a_bodyPr_attrs = f'wrap="{"square" if wrap else "none"}" anchor="{v_anchor}"'
    if auto_fit:
        a_bodyPr_attrs += ' autofit=""'
    autofit_elem = '<a:normAutofit/>' if auto_fit else ''

    paras_xml = ""
    for para in paragraphs:
        if isinstance(para, str):
            txt, fsz, bld, clr, aln = para, font_size, bold, color, align
        else:
            txt, fsz, bld, clr, aln = para

        b_attr = '1' if bld else '0'
        spc_before = f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>' if space_before else ''
        ls_elem    = f'<a:lnSpc><a:spcPct val="{line_spacing}"/></a:lnSpc>' if line_spacing else ''
        paras_xml += f"""
    <a:p>
      <a:pPr algn="{aln}">{spc_before}{ls_elem}</a:pPr>
      <a:r>
        <a:rPr lang="en-IN" sz="{fsz}" b="{b_attr}" dirty="0">
          <a:solidFill><a:srgbClr val="{clr}"/></a:solidFill>
          <a:latin typeface="Calibri"/>
          <a:cs typeface="Nirmala UI"/>
        </a:rPr>
        <a:t>{txt}</a:t>
      </a:r>
    </a:p>"""

    return f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="{name}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr {a_bodyPr_attrs}>{autofit_elem}</a:bodyPr>
    <a:lstStyle/>
    {paras_xml}
  </p:txBody>
</p:sp>"""


def bullet_txbox(sp_id, name, x, y, cx, cy, items,
                 font_size=1600, color=C_LIGHT, bullet_color=C_ACCENT,
                 line_spacing=115, space_before=80):
    """Bulleted text box with colored square bullets"""
    paras_xml = ""
    for item in items:
        spc = f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
        ls  = f'<a:lnSpc><a:spcPct val="{line_spacing}"/></a:lnSpc>'
        # bullet char: ▪ (U+25AA)
        paras_xml += f"""
    <a:p>
      <a:pPr indent="-228600" marL="228600">
        {spc}{ls}
      </a:pPr>
      <a:r>
        <a:rPr lang="en-IN" sz="{font_size}" b="0" dirty="0">
          <a:solidFill><a:srgbClr val="{bullet_color}"/></a:solidFill>
          <a:latin typeface="Calibri"/>
        </a:rPr>
        <a:t>▪  </a:t>
      </a:r>
      <a:r>
        <a:rPr lang="en-IN" sz="{font_size}" b="0" dirty="0">
          <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
          <a:latin typeface="Calibri"/>
          <a:cs typeface="Nirmala UI"/>
        </a:rPr>
        <a:t>{item}</a:t>
      </a:r>
    </a:p>"""

    return f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="{name}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="t"><a:normAutofit/></a:bodyPr>
    <a:lstStyle/>
    {paras_xml}
  </p:txBody>
</p:sp>"""


def arrow_connector(sp_id, x1, y1, x2, y2, color=C_ACCENT, w=25400):
    cx = abs(x2 - x1)
    cy = abs(y2 - y1)
    return f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="Arrow{sp_id}"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x1}" y="{y1}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="downArrow"><a:avLst>
      <a:gd name="adj1" fmla="val 50000"/>
      <a:gd name="adj2" fmla="val 50000"/>
    </a:avLst></a:prstGeom>
    {solid_fill(color)}
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>"""


def slide_xml(sp_tree_content, bg_color=C_BG):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/><a:chOff x="0" y="0"/><a:chExt cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm>
      </p:grpSpPr>
      {sp_tree_content}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


def slide_rels_xml(idx):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""


# ─────────────────────────────────────────────
# Shared layout helpers
# ─────────────────────────────────────────────

def header_bar(sp_id_start, title_text, subtitle_text="", accent=C_ACCENT):
    """Top header: dark strip + teal accent line + title"""
    parts = []
    # Top bar
    parts.append(rect_sp(sp_id_start,   "TopBar",    0, 0, SLIDE_W, emu(1.25), C_BG2))
    # Accent line under top bar
    parts.append(rect_sp(sp_id_start+1, "AccentLine", 0, emu(1.25), SLIDE_W, emu(0.06), accent))
    # Slide title
    parts.append(txbox(sp_id_start+2, "SlideTitle",
                       emu(0.4), emu(0.12), emu(11.5), emu(0.9),
                       [(title_text, 3000, True, C_WHITE, "l")]))
    if subtitle_text:
        parts.append(txbox(sp_id_start+3, "SlideSubtitle",
                           emu(0.4), emu(0.85), emu(10), emu(0.4),
                           [(subtitle_text, 1400, False, C_ACCENT, "l")]))
    return "\n".join(parts), sp_id_start + (4 if subtitle_text else 3)


def footer_bar(sp_id):
    """Bottom footer strip"""
    parts = [
        rect_sp(sp_id, "FooterBar", 0, emu(7.05), SLIDE_W, emu(0.45), C_BG2),
        rect_sp(sp_id+1, "FooterLine", 0, emu(7.05), SLIDE_W, emu(0.04), C_ACCENT),
        txbox(sp_id+2, "FooterLeft",
              emu(0.3), emu(7.1), emu(6), emu(0.35),
              [("Hindi Coreference Resolution | CIT Kokrajhar", 1000, False, C_MUTED, "l")]),
        txbox(sp_id+3, "FooterRight",
              emu(9), emu(7.1), emu(3.0), emu(0.35),
              [("B.Tech Project Presentation  |  2026", 1000, False, C_MUTED, "r")]),
    ]
    return "\n".join(parts), sp_id + 4


# ─────────────────────────────────────────────
# SLIDE BUILDERS
# ─────────────────────────────────────────────

def build_slide_01_title():
    """SLIDE 1 – Title Slide"""
    parts = []
    # Full bg already set; add diagonal decorative accent strip
    parts.append(rect_sp(10, "AccentStrip", 0, emu(5.8), SLIDE_W, emu(0.08), C_ACCENT))
    parts.append(rect_sp(11, "BottomStrip", 0, emu(6.6), SLIDE_W, emu(0.9), C_BG2))
    parts.append(rect_sp(12, "GoldLine",    0, emu(6.59), emu(4.5), emu(0.08), C_GOLD))

    # Left decorative vertical bar
    parts.append(rect_sp(13, "LeftBar", 0, 0, emu(0.12), SLIDE_H, C_ACCENT))

    # Main title
    parts.append(txbox(20, "MainTitle",
                       emu(0.5), emu(1.2), emu(11.0), emu(1.8),
                       [("Hindi Coreference Resolution", 4400, True, C_WHITE,  "l"),
                        ("for Low-Resource NLP",         4400, True, C_WHITE,  "l"),
                        ("Using Transformer-Based Hybrid Approaches", 3200, False, C_ACCENT, "l")]))

    # Subtitle / meta block
    parts.append(txbox(21, "SubTitle",
                       emu(0.5), emu(3.3), emu(9), emu(0.5),
                       [("B.Tech Project Presentation — Department of CSE, CIT Kokrajhar",
                         1600, False, C_LIGHT, "l")]))

    # Divider line
    parts.append(rect_sp(22, "Divider", emu(0.5), emu(4.0), emu(8), emu(0.05), C_GOLD))

    # Team box
    parts.append(rounded_rect_sp(30, "TeamBox", emu(0.5), emu(4.2), emu(5.8), emu(1.8), "0D1B2E", C_ACCENT, 12700))
    parts.append(txbox(31, "TeamLabel",
                       emu(0.65), emu(4.25), emu(5.5), emu(0.38),
                       [("PRESENTED BY", 1000, True, C_ACCENT, "l")]))
    parts.append(txbox(32, "TeamNames",
                       emu(0.65), emu(4.65), emu(5.5), emu(1.25),
                       [("Himraj Doley  •  Subham Saha  •  Mohd Zaid  •  Nipujyoti Rabha",
                         1500, False, C_WHITE, "l")]))

    # Supervisor box
    parts.append(rounded_rect_sp(35, "SupBox", emu(6.6), emu(4.2), emu(5.2), emu(1.8), "0D1B2E", C_GOLD, 12700))
    parts.append(txbox(36, "SupLabel",
                       emu(6.75), emu(4.25), emu(4.9), emu(0.38),
                       [("SUPERVISOR", 1000, True, C_GOLD, "l")]))
    parts.append(txbox(37, "SupName",
                       emu(6.75), emu(4.65), emu(4.9), emu(1.25),
                       [("Dr. Apurbalal Senapati", 1700, False, C_WHITE, "l")]))

    # Footer
    parts.append(rect_sp(40, "FooterBg", 0, emu(6.7), SLIDE_W, emu(0.8), C_BG2))
    parts.append(txbox(41, "FooterTxt",
                       emu(0.5), emu(6.75), emu(11), emu(0.6),
                       [("Department of Computer Science & Engineering  |  Central Institute of Technology Kokrajhar  |  2026",
                         1100, False, C_MUTED, "ctr")]))

    return slide_xml("\n".join(parts))


def build_slide_02_toc():
    """SLIDE 2 – Table of Contents / Agenda"""
    parts = []
    hdr, nid = header_bar(10, "Table of Contents", "Research Presentation Agenda")

    # Two-column topic grid
    left_topics = [
        ("01", "Introduction & Problem Statement"),
        ("02", "Motivation & Objectives"),
        ("03", "Literature Review"),
        ("04", "Dataset Description"),
        ("05", "Phase 1 Methodology"),
        ("06", "Phase 1: Challenges"),
        ("07", "Phase 2 Hybrid Architecture"),
        ("08", "System Pipeline"),
        ("09", "Implementation Details"),
    ]
    right_topics = [
        ("10", "Experimental Results"),
        ("11", "Example Output"),
        ("12", "Evaluation Metrics"),
        ("13", "Comparative Analysis"),
        ("14", "Limitations"),
        ("15", "Future Work"),
        ("16", "Conclusion"),
        ("17", "References"),
        ("18", "Thank You"),
    ]

    col_y = emu(1.55)
    row_h = emu(0.52)

    for i, (num, topic) in enumerate(left_topics):
        y = col_y + i * row_h
        parts.append(rounded_rect_sp(nid,   f"LNum{i}", emu(0.35), y, emu(0.5), emu(0.38), C_ACCENT, adj=50000))
        parts.append(txbox(nid+1, f"LN{i}", emu(0.36), y+emu(0.02), emu(0.48), emu(0.36),
                           [(num, 1400, True, C_BG, "ctr")]))
        parts.append(txbox(nid+2, f"LT{i}", emu(0.95), y+emu(0.02), emu(5.0), emu(0.38),
                           [(topic, 1400, False, C_LIGHT, "l")]))
        nid += 3

    for i, (num, topic) in enumerate(right_topics):
        y = col_y + i * row_h
        parts.append(rounded_rect_sp(nid,   f"RNum{i}", emu(6.5), y, emu(0.5), emu(0.38), C_GOLD, adj=50000))
        parts.append(txbox(nid+1, f"RN{i}", emu(6.51), y+emu(0.02), emu(0.48), emu(0.36),
                           [(num, 1400, True, C_BG, "ctr")]))
        parts.append(txbox(nid+2, f"RT{i}", emu(7.1), y+emu(0.02), emu(5.0), emu(0.38),
                           [(topic, 1400, False, C_LIGHT, "l")]))
        nid += 3

    ftr, _ = footer_bar(nid)
    parts_all = hdr + "\n" + "\n".join(parts) + "\n" + ftr
    return slide_xml(parts_all)


def build_slide_03_intro():
    """SLIDE 3 – Introduction"""
    parts = []
    hdr, nid = header_bar(10, "Introduction", "What is Coreference Resolution?")
    parts.append(hdr)

    # Definition card
    parts.append(rounded_rect_sp(nid, "DefCard", emu(0.35), emu(1.5), emu(11.5), emu(1.35), "0D1B2E", C_ACCENT, 15000))
    parts.append(txbox(nid+1, "DefTitle", emu(0.55), emu(1.56), emu(11.0), emu(0.38),
                       [("DEFINITION", 1100, True, C_ACCENT, "l")]))
    parts.append(txbox(nid+2, "DefText",  emu(0.55), emu(1.95), emu(11.0), emu(0.75),
                       [("Coreference Resolution identifies and groups all mentions in a text that refer to the same real-world entity.",
                         1700, False, C_WHITE, "l")]))
    nid += 3

    # Hindi example box
    parts.append(rounded_rect_sp(nid, "ExBox", emu(0.35), emu(3.05), emu(11.5), emu(1.1), "122840", C_GOLD, 12700))
    parts.append(txbox(nid+1, "ExLabel", emu(0.55), emu(3.1), emu(11.0), emu(0.38),
                       [("HINDI EXAMPLE", 1100, True, C_GOLD, "l")]))
    parts.append(txbox(nid+2, "ExHindi", emu(0.55), emu(3.5), emu(11.0), emu(0.55),
                       [("'राम विद्यालय गया।  वह प्रसन्न था।'   →   राम  =  वह  (same entity)",
                         1700, False, C_WHITE, "ctr")]))
    nid += 3

    # Application tiles
    apps = [("Question\nAnswering", C_ACCENT), ("Machine\nTranslation", C_GOLD),
            ("Text\nSummarization", "5C9BD6"), ("Dialogue\nSystems", "A06BBE")]
    tile_w, tile_h = emu(2.5), emu(0.9)
    tile_y = emu(4.35)
    tile_gap = emu(0.28)
    for i, (label, color) in enumerate(apps):
        tx = emu(0.35) + i * (tile_w + tile_gap)
        parts.append(rounded_rect_sp(nid, f"App{i}", tx, tile_y, tile_w, tile_h, color, adj=30000))
        parts.append(txbox(nid+1, f"AppTxt{i}", tx, tile_y, tile_w, tile_h,
                           [(label, 1500, True, C_BG, "ctr")], v_anchor="ctr"))
        nid += 2

    # Hindi NLP note
    parts.append(txbox(nid, "HindiNote", emu(0.35), emu(5.55), emu(11.5), emu(0.45),
                       [("★  Hindi NLP remains a low-resource research area — annotated coreference datasets are extremely scarce.",
                         1400, False, C_MUTED, "l")]))
    nid += 1

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_04_challenges():
    """SLIDE 4 – Challenges in Hindi Coreference"""
    parts = []
    hdr, nid = header_bar(10, "Challenges in Hindi Coreference", "Why is this problem hard?")
    parts.append(hdr)

    challenges = [
        (C_ACCENT, "Morphological Richness",
         "Hindi inflects words for gender, number, case, and tense — creating high surface variation for the same entity."),
        (C_GOLD,   "Free Word Order",
         "Hindi is a Subject-Object-Verb language with flexible constituent ordering, making positional heuristics unreliable."),
        ("5C9BD6",  "Pro-Drop Phenomenon",
         "Subject pronouns are frequently omitted ('pro-drop'), requiring deep discourse context for resolution."),
        ("A06BBE",  "Devanagari Tokenization",
         "Complex script with conjunct consonants and zero-width characters complicates robust tokenization."),
        (C_RED,    "Dataset Scarcity",
         "Publicly available annotated Hindi coreference corpora are extremely limited, hampering supervised training."),
    ]

    card_h = emu(0.88)
    card_y_start = emu(1.52)
    card_gap = emu(0.12)

    for i, (color, title, desc) in enumerate(challenges):
        y = card_y_start + i * (card_h + card_gap)
        parts.append(rect_sp(nid,   f"CBar{i}", emu(0.35), y, emu(0.07), card_h, color))
        parts.append(rounded_rect_sp(nid+1, f"CCard{i}", emu(0.47), y, emu(11.38), card_h, "0D1B2E"))
        parts.append(txbox(nid+2, f"CTitle{i}", emu(0.62), y+emu(0.08), emu(3.5), emu(0.4),
                           [(title, 1600, True, color, "l")]))
        parts.append(txbox(nid+3, f"CDesc{i}",  emu(0.62), y+emu(0.47), emu(11.0), emu(0.38),
                           [(desc, 1350, False, C_LIGHT, "l")]))
        nid += 4

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_05_objectives():
    """SLIDE 5 – Motivation & Objectives"""
    parts = []
    hdr, nid = header_bar(10, "Motivation & Objectives", "Why this project? What do we aim to achieve?")
    parts.append(hdr)

    # Motivation column (left)
    parts.append(rect_sp(nid, "MotDivider", emu(0.35), emu(1.52), emu(0.06), emu(4.8), C_ACCENT))
    parts.append(txbox(nid+1, "MotHeader", emu(0.55), emu(1.52), emu(5.3), emu(0.45),
                       [("MOTIVATION", 1700, True, C_ACCENT, "l")]))
    motiv = [
        "Coreference is critical for advanced NLP pipelines (QA, summarization, translation).",
        "Hindi is spoken by 600M+ people, yet severely under-resourced in NLP.",
        "Existing English-focused tools fail on Hindi morphology and script.",
        "Pretrained Indic language models (IndicBERTv2) offer new opportunities.",
        "Research gap: No robust open-source Hindi coreference system exists.",
    ]
    parts.append(bullet_txbox(nid+2, "MotBullets", emu(0.55), emu(2.05), emu(5.4), emu(4.0),
                              motiv, font_size=1400, color=C_LIGHT, bullet_color=C_ACCENT))
    nid += 3

    # Objectives column (right)
    parts.append(rect_sp(nid, "ObjDivider", emu(6.6), emu(1.52), emu(0.06), emu(4.8), C_GOLD))
    parts.append(txbox(nid+1, "ObjHeader", emu(6.8), emu(1.52), emu(5.3), emu(0.45),
                       [("OBJECTIVES", 1700, True, C_GOLD, "l")]))
    objectives = [
        "Build an unsupervised and supervised Hindi coreference pipeline.",
        "Train or adapt pretrained Indic transformer models for mention encoding.",
        "Create a manually annotated Hindi coreference benchmark dataset.",
        "Engineer linguistically motivated features for pairwise classification.",
        "Evaluate and compare Phase 1 (unsupervised) vs Phase 2 (supervised hybrid).",
    ]
    parts.append(bullet_txbox(nid+2, "ObjBullets", emu(6.8), emu(2.05), emu(5.4), emu(4.0),
                              objectives, font_size=1400, color=C_LIGHT, bullet_color=C_GOLD))
    nid += 3

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_06_litreview():
    """SLIDE 6 – Literature Review"""
    parts = []
    hdr, nid = header_bar(10, "Literature Review", "Prior work in coreference resolution and Hindi NLP")
    parts.append(hdr)

    rows = [
        ("Lee et al. (2017)",       "e2e-coref",       "End-to-end neural coref. Span representations + attention. English-focused."),
        ("Clark & Manning (2016)",  "Deep RL Coref",   "Reinforcement learning for antecedent selection. State-of-art at time."),
        ("Kumar et al. (2020)",     "Hindi Coref",     "Rule-based Hindi coreference using morphological constraints and heuristics."),
        ("Koto et al. (2021)",      "IndicBERT",       "Multilingual BERT variant trained on 12 Indic languages including Hindi."),
        ("Dobrovolskii (2021)",     "Word-level Coref","Reformulation avoiding span enumeration. Efficient for long documents."),
        ("IndicNLP Suite (2022)",   "AI4Bharat",       "IndicBERTv2 — state-of-art contextual embeddings for Indic languages."),
    ]

    # Header row
    col_x = [emu(0.35), emu(3.3), emu(5.8)]
    col_w = [emu(2.9),  emu(2.4), emu(5.9)]
    row_h = emu(0.55)
    header_y = emu(1.55)

    for j, (label, w, x) in enumerate(zip(["Reference", "System / Model", "Key Contribution"], col_w, col_x)):
        parts.append(rect_sp(nid, f"TH{j}", x, header_y, w, row_h, C_ACCENT if j==0 else ("112238" if j%2==0 else "0D1B2E")))
        parts.append(txbox(nid+1, f"THT{j}", x+emu(0.1), header_y+emu(0.08), w-emu(0.15), row_h-emu(0.1),
                           [(label, 1500, True, C_WHITE, "l")]))
        nid += 2

    for i, (ref, sys, contrib) in enumerate(rows):
        row_y = header_y + (i+1) * row_h
        row_bg = "122840" if i % 2 == 0 else "0D1B2E"
        for j, (text, w, x) in enumerate(zip([ref, sys, contrib], col_w, col_x)):
            parts.append(rect_sp(nid, f"TC{i}{j}", x, row_y, w, row_h, row_bg))
            clr = C_ACCENT if j == 0 else (C_GOLD if j == 1 else C_LIGHT)
            parts.append(txbox(nid+1, f"TCT{i}{j}", x+emu(0.1), row_y+emu(0.07), w-emu(0.15), row_h-emu(0.1),
                               [(text, 1300, j==0, clr, "l")]))
            nid += 2

    # Research gap note
    parts.append(rounded_rect_sp(nid, "GapBox", emu(0.35), emu(6.15), emu(11.5), emu(0.75), "0D1B2E", C_GOLD, 12700))
    parts.append(txbox(nid+1, "GapTxt", emu(0.55), emu(6.22), emu(11.0), emu(0.6),
                       [("Research Gap: No end-to-end supervised Hindi coreference system with Indic pretrained models exists in open literature.",
                         1400, False, C_LIGHT, "l")]))
    nid += 2

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_07_dataset():
    """SLIDE 7 – Dataset Description"""
    parts = []
    hdr, nid = header_bar(10, "Dataset Description", "Training corpus and annotated benchmark")
    parts.append(hdr)

    # Phase 1 dataset
    parts.append(rounded_rect_sp(nid, "P1Box", emu(0.35), emu(1.55), emu(5.55), emu(4.95), "0D1B2E", C_ACCENT, 12700))
    parts.append(rect_sp(nid+1, "P1Header", emu(0.35), emu(1.55), emu(5.55), emu(0.55), C_ACCENT))
    parts.append(txbox(nid+2, "P1HT", emu(0.5), emu(1.6), emu(5.3), emu(0.45),
                       [("PHASE 1 — Pretraining Corpus", 1600, True, C_BG, "l")]))
    p1_items = [
        "~172,000 Hindi Wikipedia documents",
        "Raw text — no coreference annotations",
        "Used for MLM pretraining of compact BERT",
        "SentencePiece BPE tokenizer (vocab = 32K)",
        "Train/Val split: 90% / 10%",
    ]
    parts.append(bullet_txbox(nid+3, "P1Bullets", emu(0.5), emu(2.25), emu(5.2), emu(4.0),
                              p1_items, font_size=1400, color=C_LIGHT, bullet_color=C_ACCENT))
    nid += 4

    # Phase 2 dataset
    parts.append(rounded_rect_sp(nid, "P2Box", emu(6.25), emu(1.55), emu(5.55), emu(4.95), "0D1B2E", C_GOLD, 12700))
    parts.append(rect_sp(nid+1, "P2Header", emu(6.25), emu(1.55), emu(5.55), emu(0.55), C_GOLD))
    parts.append(txbox(nid+2, "P2HT", emu(6.4), emu(1.6), emu(5.3), emu(0.45),
                       [("PHASE 2 — Annotated Benchmark", 1600, True, C_BG, "l")]))
    p2_items = [
        "1,248 manually annotated mention spans",
        "4 mention types: proper noun, common noun,\n  pronoun, nominal phrase",
        "Pairwise positive/negative coreferent labels",
        "Diverse sentence types and domains",
        "Annotated by project team with inter-annotator review",
    ]
    parts.append(bullet_txbox(nid+3, "P2Bullets", emu(6.4), emu(2.25), emu(5.2), emu(4.0),
                              p2_items, font_size=1400, color=C_LIGHT, bullet_color=C_GOLD))
    nid += 4

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_08_phase1():
    """SLIDE 8 – Phase 1 Methodology"""
    parts = []
    hdr, nid = header_bar(10, "Phase 1: Unsupervised BERT + Clustering",
                          "Scratch-trained language model with cosine-similarity clustering")
    parts.append(hdr)

    # Pipeline flow boxes
    steps = [
        (C_ACCENT,  "1. Hindi Wikipedia\n   (~172K docs)"),
        ("5C9BD6",  "2. SentencePiece\n   BPE Tokenizer"),
        ("A06BBE",  "3. Compact BERT\n   (MLM Training)"),
        (C_GOLD,    "4. Mention\n   Embeddings"),
        (C_ACCENT,  "5. Cosine Similarity\n   + Clustering"),
        (C_RED,     "6. Coreference\n   Clusters (Output)"),
    ]

    box_w, box_h = emu(1.7), emu(1.1)
    box_y = emu(1.65)
    arrow_w = emu(0.35)
    start_x = emu(0.3)

    for i, (color, label) in enumerate(steps):
        bx = start_x + i * (box_w + arrow_w)
        parts.append(rounded_rect_sp(nid, f"StepBox{i}", bx, box_y, box_w, box_h, color, adj=20000))
        parts.append(txbox(nid+1, f"StepTxt{i}", bx, box_y, box_w, box_h,
                           [(label, 1200, True, C_BG, "ctr")], v_anchor="ctr"))
        nid += 2
        if i < len(steps) - 1:
            ax = bx + box_w
            ay = box_y + box_h // 2 - emu(0.1)
            parts.append(rect_sp(nid, f"Arrow{i}", ax, ay, arrow_w, emu(0.2), C_MUTED))
            parts.append(txbox(nid+1, f"Arrowhead{i}", ax + arrow_w - emu(0.18), ay - emu(0.1),
                               emu(0.2), emu(0.4), [("▶", 1000, False, C_MUTED, "ctr")]))
            nid += 2

    # Key details
    parts.append(rect_sp(nid, "DetailsDivider", emu(0.35), emu(3.05), SLIDE_W - emu(0.7), emu(0.04), C_ACCENT))
    nid += 1

    details_l = [
        "BERT Architecture: 6 layers, 8 heads, hidden=512",
        "Tokenizer: SentencePiece BPE, vocab size 32,000",
        "Training objective: Masked Language Model (MLM)",
        "Pretraining dataset: Hindi Wikipedia (~172K docs)",
    ]
    details_r = [
        "Embedding extraction: [CLS] + mean-pool strategy",
        "Similarity metric: Cosine similarity between mention embeddings",
        "Clustering threshold: Tuned empirically (0.75–0.85)",
        "No annotated coreference data used in Phase 1",
    ]
    parts.append(bullet_txbox(nid,   "DL", emu(0.35), emu(3.2), emu(5.7), emu(3.2),
                              details_l, font_size=1350, color=C_LIGHT))
    parts.append(bullet_txbox(nid+1, "DR", emu(6.2),  emu(3.2), emu(5.7), emu(3.2),
                              details_r, font_size=1350, color=C_LIGHT, bullet_color=C_GOLD))
    nid += 2

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_09_phase1_fail():
    """SLIDE 9 – Why Phase 1 Failed"""
    parts = []
    hdr, nid = header_bar(10, "Why Phase 1 Failed", "Critical analysis and research insight", accent=C_RED)
    parts.append(hdr)

    # Central insight banner
    parts.append(rounded_rect_sp(nid, "InsightBanner", emu(0.35), emu(1.55), emu(11.5), emu(0.8), C_RED))
    parts.append(txbox(nid+1, "InsightTxt", emu(0.55), emu(1.65), emu(11.0), emu(0.6),
                       [("KEY INSIGHT:  Semantic Similarity  ≠  Coreference",
                         2200, True, C_WHITE, "ctr")]))
    nid += 2

    failures = [
        (C_RED,    "Semantic ≠ Coreference",
         "'राम' and 'विद्यालय' may have similar embeddings in some contexts, but they are not coreferent."),
        ("E07B39",  "Pronoun-Antecedent Gap",
         "Pronouns like 'वह' and 'उसने' lack strong lexical similarity to their antecedents, causing missed links."),
        (C_GOLD,   "Spurious Clusters",
         "Low similarity thresholds merged semantically related but non-coreferent mentions into wrong clusters."),
        ("5C9BD6",  "No Discourse Modeling",
         "Embedding similarity ignores sentence distance, grammatical role, and entity salience — all crucial cues."),
        (C_MUTED,  "No Antecedent Ranking",
         "All mention pairs were treated equally; no mechanism to prefer the most recent or salient antecedent."),
    ]

    card_h = emu(0.82)
    card_y_start = emu(2.55)
    card_gap = emu(0.1)

    for i, (color, title, desc) in enumerate(failures):
        y = card_y_start + i * (card_h + card_gap)
        parts.append(rect_sp(nid,   f"FBar{i}", emu(0.35), y, emu(0.07), card_h, color))
        parts.append(rounded_rect_sp(nid+1, f"FCard{i}", emu(0.47), y, emu(11.38), card_h, "0D1B2E"))
        parts.append(txbox(nid+2, f"FTitle{i}", emu(0.65), y+emu(0.07), emu(3.2), emu(0.38),
                           [(title, 1550, True, color, "l")]))
        parts.append(txbox(nid+3, f"FDesc{i}",  emu(0.65), y+emu(0.44), emu(11.0), emu(0.35),
                           [(desc, 1300, False, C_LIGHT, "l")]))
        nid += 4

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_10_phase2():
    """SLIDE 10 – Phase 2 Hybrid Architecture"""
    parts = []
    hdr, nid = header_bar(10, "Phase 2: Hybrid Supervised Architecture",
                          "IndicBERTv2 embeddings + linguistic features + pairwise classification")
    parts.append(hdr)

    comps = [
        (C_ACCENT, "IndicBERTv2\nEncoder",
         "State-of-art Indic language model.\nGenerates contextual mention embeddings."),
        (C_GOLD,   "Feature\nEngineering",
         "Pronoun flags, plurality, sentence\ndistance, clause heuristics."),
        ("A06BBE",  "Pairwise\nClassifier",
         "Logistic Regression trained on\npositive/negative mention pairs."),
        ("5C9BD6",  "Union-Find\nClustering",
         "Transitively merges coreferent\npairs into entity clusters."),
    ]

    box_w, box_h = emu(2.5), emu(2.2)
    box_y = emu(1.6)
    gap = emu(0.4)
    start_x = emu(0.5)

    for i, (color, title, desc) in enumerate(comps):
        bx = start_x + i * (box_w + gap)
        parts.append(rounded_rect_sp(nid, f"CompBox{i}", bx, box_y, box_w, box_h, "0D1B2E", color, 20000, adj=15000))
        parts.append(rect_sp(nid+1, f"CompTop{i}", bx, box_y, box_w, emu(0.08), color))
        parts.append(txbox(nid+2, f"CompTitle{i}", bx, box_y+emu(0.15), box_w, emu(0.6),
                           [(title, 1600, True, color, "ctr")]))
        parts.append(txbox(nid+3, f"CompDesc{i}",  bx, box_y+emu(0.9), box_w, emu(1.2),
                           [(desc, 1300, False, C_LIGHT, "ctr")]))
        nid += 4
        # Arrow between boxes
        if i < len(comps) - 1:
            ax = bx + box_w + emu(0.05)
            ay = box_y + box_h//2 - emu(0.15)
            parts.append(txbox(nid, f"CompArrow{i}", ax, ay, gap - emu(0.1), emu(0.3),
                               [("▶", 1800, False, C_MUTED, "ctr")]))
            nid += 1

    # Key advantages
    parts.append(rect_sp(nid, "AdvDivider", emu(0.35), emu(4.1), SLIDE_W - emu(0.7), emu(0.04), C_GOLD))
    nid += 1
    parts.append(txbox(nid, "AdvHeader", emu(0.35), emu(4.2), emu(11.5), emu(0.42),
                       [("KEY ADVANTAGES OF HYBRID APPROACH", 1500, True, C_GOLD, "ctr")]))
    nid += 1

    adv = [
        "Leverages powerful pretrained Indic language model — no scratch training needed.",
        "Linguistic features capture morphological and discourse properties missed by neural embeddings alone.",
        "Pairwise classification is interpretable and trainable on small datasets.",
        "Union-Find ensures globally consistent entity clusters with O(α) efficiency.",
    ]
    parts.append(bullet_txbox(nid, "AdvBullets", emu(0.35), emu(4.7), emu(11.5), emu(2.0),
                              adv, font_size=1400, color=C_LIGHT, bullet_color=C_GOLD))
    nid += 1

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_11_pipeline():
    """SLIDE 11 – System Pipeline"""
    parts = []
    hdr, nid = header_bar(10, "System Pipeline", "End-to-end architecture for Hindi coreference resolution")
    parts.append(hdr)

    pipeline_steps = [
        (C_ACCENT,  "INPUT",     "Hindi Text with\nMention Spans"),
        ("5C9BD6",   "ENCODE",    "IndicBERTv2\nMention Embeddings"),
        (C_GOLD,    "FEATURES",  "Linguistic\nFeature Engineering"),
        ("A06BBE",  "CLASSIFY",  "Logistic Regression\nPairwise Scorer"),
        ("6BBE8C",  "CLUSTER",   "Union-Find\nCluster Reconstruction"),
        (C_ACCENT,  "OUTPUT",    "Entity Coreference\nClusters"),
    ]

    box_w, box_h = emu(1.65), emu(1.15)
    arrow_h = emu(0.42)
    center_x = SLIDE_W // 2 - box_w // 2
    start_y = emu(1.55)
    step_h = box_h + arrow_h

    # Vertical pipeline – two columns to save space
    col_items = [pipeline_steps[:3], pipeline_steps[3:]]
    col_x = [emu(2.3), emu(8.2)]

    for col_idx, col_steps in enumerate(col_items):
        cx = col_x[col_idx]
        for row_idx, (color, stage, desc) in enumerate(col_steps):
            by = start_y + row_idx * step_h
            parts.append(rounded_rect_sp(nid, f"PBox{col_idx}{row_idx}",
                                         cx, by, box_w, box_h, color, adj=15000))
            parts.append(txbox(nid+1, f"PStage{col_idx}{row_idx}",
                               cx, by+emu(0.08), box_w, emu(0.35),
                               [(stage, 1100, True, C_BG, "ctr")]))
            parts.append(txbox(nid+2, f"PDesc{col_idx}{row_idx}",
                               cx, by+emu(0.45), box_w, emu(0.65),
                               [(desc, 1200, False, C_BG, "ctr")]))
            nid += 3
            # Down arrow
            if row_idx < len(col_steps) - 1:
                parts.append(txbox(nid, f"PArrow{col_idx}{row_idx}",
                                   cx + box_w//2 - emu(0.15), by + box_h + emu(0.02),
                                   emu(0.3), arrow_h - emu(0.1),
                                   [("▼", 1400, False, color, "ctr")]))
                nid += 1

    # Cross-column connector
    parts.append(txbox(nid, "ColConnector",
                       emu(4.2), emu(3.95), emu(3.9), emu(0.42),
                       [("─────────────────── ▶", 1400, False, C_MUTED, "ctr")]))
    nid += 1

    # Feature list on right-center
    parts.append(rounded_rect_sp(nid, "FeatBox", emu(4.0), emu(1.55), emu(3.6), emu(3.85), "0D1B2E", C_ACCENT, 12700))
    parts.append(txbox(nid+1, "FeatHeader", emu(4.1), emu(1.6), emu(3.4), emu(0.42),
                       [("FEATURE VECTOR", 1300, True, C_ACCENT, "ctr")]))
    feats = [
        "cos_sim: embedding cosine similarity",
        "is_pronoun_1 / is_pronoun_2",
        "is_plural_1 / is_plural_2",
        "sent_distance: sentence gap",
        "same_clause: clause heuristic",
        "mention_type_match flag",
    ]
    parts.append(bullet_txbox(nid+2, "FeatList", emu(4.1), emu(2.1), emu(3.4), emu(3.1),
                              feats, font_size=1200, color=C_LIGHT, bullet_color=C_ACCENT))
    nid += 3

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_12_results():
    """SLIDE 12 – Experimental Results"""
    parts = []
    hdr, nid = header_bar(10, "Experimental Results", "Quantitative evaluation of Phase 1 vs Phase 2")
    parts.append(hdr)

    # Summary stat cards
    stats = [
        (C_ACCENT, "82%",  "Phase 2 Pairwise\nAccuracy"),
        (C_GOLD,   "71%",  "Embedding-only\nBaseline"),
        ("5C9BD6",  "+11%", "Improvement from\nLinguistic Features"),
        (C_RED,    "~43%", "Phase 1 Accuracy\n(Clustering-based)"),
    ]
    card_w, card_h = emu(2.65), emu(1.35)
    card_y = emu(1.58)
    gap = emu(0.24)
    sx = emu(0.35)
    for i, (color, val, label) in enumerate(stats):
        cx = sx + i * (card_w + gap)
        parts.append(rounded_rect_sp(nid, f"StatCard{i}", cx, card_y, card_w, card_h, "0D1B2E", color, 15000, adj=10000))
        parts.append(txbox(nid+1, f"StatVal{i}", cx, card_y+emu(0.1), card_w, emu(0.65),
                           [(val, 3600, True, color, "ctr")]))
        parts.append(txbox(nid+2, f"StatLbl{i}", cx, card_y+emu(0.78), card_w, emu(0.52),
                           [(label, 1200, False, C_LIGHT, "ctr")]))
        nid += 3

    # Results table
    table_y = emu(3.15)
    col_data = [
        ("System",           C_ACCENT, [("Phase 1 (BERT Clustering)", C_LIGHT),
                                         ("Phase 2 (Embedding only)", C_LIGHT),
                                         ("Phase 2 (Full Hybrid)", C_LIGHT),
                                         ("Phase 2 + Linguistic Feats", C_WHITE)]),
        ("Precision",        C_GOLD,   [("~38%", C_LIGHT), ("~68%", C_LIGHT), ("~78%", C_LIGHT), ("~80%", C_WHITE)]),
        ("Recall",           C_GOLD,   [("~51%", C_LIGHT), ("~74%", C_LIGHT), ("~79%", C_LIGHT), ("~84%", C_WHITE)]),
        ("Accuracy",         C_ACCENT, [("~43%", C_RED),   ("~71%", C_MUTED), ("~79%", C_LIGHT), ("~82%", C_ACCENT)]),
        ("Improvement",      C_MUTED,  [("Baseline",C_MUTED),("+28pp",C_LIGHT),("+36pp",C_LIGHT),("+39pp",C_GOLD)]),
    ]
    col_xs = [emu(0.35), emu(3.85), emu(5.85), emu(7.85), emu(9.85)]
    col_ws = [emu(3.45), emu(1.95), emu(1.95), emu(1.95), emu(1.95)]
    row_h = emu(0.52)

    # Header
    for j, (htext, hcolor, _) in enumerate(col_data):
        parts.append(rect_sp(nid, f"TH{j}", col_xs[j], table_y, col_ws[j], row_h,
                             C_ACCENT if j==0 else "1A3D5C"))
        parts.append(txbox(nid+1, f"THT{j}", col_xs[j]+emu(0.1), table_y+emu(0.08),
                           col_ws[j]-emu(0.15), row_h-emu(0.1),
                           [(htext, 1400, True, C_WHITE, "ctr")]))
        nid += 2

    # Data rows
    for row_i in range(4):
        row_y = table_y + (row_i + 1) * row_h
        row_bg = "122840" if row_i % 2 == 0 else "0D1B2E"
        for j, (_, _, rows) in enumerate(col_data):
            cell_txt, cell_clr = rows[row_i]
            is_last = row_i == 3
            parts.append(rect_sp(nid, f"TC{row_i}{j}", col_xs[j], row_y, col_ws[j], row_h,
                                 "1A2E4A" if is_last else row_bg,
                                 C_GOLD if is_last else None, 8000 if is_last else 0))
            parts.append(txbox(nid+1, f"TCT{row_i}{j}", col_xs[j]+emu(0.1), row_y+emu(0.08),
                               col_ws[j]-emu(0.15), row_h-emu(0.1),
                               [(cell_txt, 1350, is_last, cell_clr, "ctr")]))
            nid += 2

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_13_example():
    """SLIDE 13 – Example Output"""
    parts = []
    hdr, nid = header_bar(10, "Example Output", "Demonstrating pronoun-antecedent coreference resolution")
    parts.append(hdr)

    # Input text box
    parts.append(rounded_rect_sp(nid, "InpBox", emu(0.35), emu(1.58), emu(11.5), emu(1.1), "0D1B2E", C_ACCENT, 12700))
    parts.append(txbox(nid+1, "InpLabel", emu(0.55), emu(1.63), emu(11.0), emu(0.38),
                       [("INPUT TEXT", 1100, True, C_ACCENT, "l")]))
    parts.append(txbox(nid+2, "InpText",  emu(0.55), emu(2.0), emu(11.0), emu(0.6),
                       [("'राम विद्यालय गया।  वह बहुत प्रसन्न था।  उसने अपने मित्रों से बात की।'",
                         1700, False, C_WHITE, "ctr")]))
    nid += 3

    # Mention highlighting
    parts.append(rect_sp(nid, "MentionDivider", emu(0.35), emu(2.88), SLIDE_W - emu(0.7), emu(0.04), C_MUTED))
    nid += 1
    parts.append(txbox(nid, "MentionHeader", emu(0.35), emu(2.98), emu(11.5), emu(0.4),
                       [("IDENTIFIED MENTIONS  →  Resolved Entity Cluster", 1500, True, C_GOLD, "ctr")]))
    nid += 1

    mentions = [
        ("राम",   "Proper Noun", C_ACCENT),
        ("वह",    "Pronoun",     C_GOLD),
        ("उसने",  "Pronoun",     "A06BBE"),
    ]
    m_w, m_h = emu(2.8), emu(1.35)
    m_y = emu(3.55)
    m_gap = emu(0.55)
    arrow_x_start = emu(0.35) + m_w + m_gap * 0.4

    for i, (word, mtype, color) in enumerate(mentions):
        mx = emu(0.35) + i * (m_w + m_gap)
        parts.append(rounded_rect_sp(nid, f"MBox{i}", mx, m_y, m_w, m_h, "0D1B2E", color, 18000, adj=15000))
        parts.append(txbox(nid+1, f"MWord{i}", mx, m_y+emu(0.15), m_w, emu(0.65),
                           [(word, 2800, True, color, "ctr")]))
        parts.append(txbox(nid+2, f"MType{i}", mx, m_y+emu(0.85), m_w, emu(0.42),
                           [(mtype, 1200, False, C_MUTED, "ctr")]))
        nid += 3
        if i < len(mentions) - 1:
            ax = mx + m_w + m_gap * 0.1
            parts.append(txbox(nid, f"MArrow{i}", ax, m_y + m_h//2 - emu(0.15),
                               m_gap * 0.8, emu(0.3), [("═══►", 1200, False, C_MUTED, "ctr")]))
            nid += 1

    # Result cluster box
    parts.append(rounded_rect_sp(nid, "ClusterBox", emu(0.35), emu(5.15), emu(11.5), emu(0.85), C_ACCENT))
    parts.append(txbox(nid+1, "ClusterLabel", emu(0.55), emu(5.2), emu(3.0), emu(0.38),
                       [("PREDICTED CLUSTER:", 1400, True, C_BG, "l")]))
    parts.append(txbox(nid+2, "ClusterValue", emu(3.8), emu(5.2), emu(7.8), emu(0.38),
                       [("{  राम  ,  वह  ,  उसने  }  →  Single Entity", 1700, True, C_BG, "ctr")]))
    nid += 3

    parts.append(txbox(nid, "Note", emu(0.35), emu(6.15), emu(11.5), emu(0.4),
                       [("Note: Some ambiguity cases (pro-drop, zero-anaphora) remain challenging for the current model.",
                         1300, False, C_MUTED, "l")]))
    nid += 1

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_14_limitations():
    """SLIDE 14 – Limitations"""
    parts = []
    hdr, nid = header_bar(10, "Limitations", "Known constraints and areas for improvement", accent=C_RED)
    parts.append(hdr)

    limitations = [
        (C_RED,     "Small Annotated Dataset",
         "Only 1,248 mention spans were manually annotated. This limits the generalization capacity of the trained classifier."),
        ("E07B39",  "No Automatic Mention Detection",
         "Mention spans are currently provided manually. An automatic mention detection module is absent."),
        (C_GOLD,   "Limited Classifier Capacity",
         "Logistic Regression is a linear model; it may not capture complex non-linear feature interactions."),
        ("5C9BD6",  "Union-Find Error Propagation",
         "Incorrect pairwise predictions transitively propagate, potentially merging entire chains incorrectly."),
        (C_MUTED,  "No CoNLL Evaluation",
         "Standard MUC, B³, and CEAF metrics from the CoNLL-2012 shared task were not implemented for Phase 2."),
        ("A06BBE",  "Domain Coverage",
         "Training data is Wikipedia-skewed; performance on conversational or domain-specific Hindi is unknown."),
    ]

    card_h = emu(0.74)
    card_y_start = emu(1.58)
    card_gap = emu(0.1)

    for i, (color, title, desc) in enumerate(limitations):
        y = card_y_start + i * (card_h + card_gap)
        parts.append(rect_sp(nid,   f"LBar{i}", emu(0.35), y, emu(0.07), card_h, color))
        parts.append(rounded_rect_sp(nid+1, f"LCard{i}", emu(0.47), y, emu(11.38), card_h, "0D1B2E"))
        parts.append(txbox(nid+2, f"LTitle{i}", emu(0.65), y+emu(0.06), emu(3.5), emu(0.35),
                           [(title, 1500, True, color, "l")]))
        parts.append(txbox(nid+3, f"LDesc{i}",  emu(0.65), y+emu(0.4), emu(11.0), emu(0.3),
                           [(desc, 1250, False, C_LIGHT, "l")]))
        nid += 4

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_15_future():
    """SLIDE 15 – Future Work"""
    parts = []
    hdr, nid = header_bar(10, "Future Work", "Research directions and planned enhancements")
    parts.append(hdr)

    future_items = [
        (C_ACCENT, "Automatic Mention Detection",
         "Train a sequence-labeling (NER/span) model for automatic mention span extraction, removing the manual bottleneck."),
        (C_GOLD,   "Neural Pairwise Scorer",
         "Replace Logistic Regression with a neural scorer (MLP / cross-attention) to capture non-linear feature interactions."),
        ("5C9BD6",  "Transformer Fine-Tuning",
         "Fine-tune IndicBERTv2 end-to-end on the coreference task (span-BERT style) for richer mention representations."),
        ("A06BBE",  "Larger & More Diverse Dataset",
         "Expand the annotated corpus to 10K+ spans, covering news, literature, and conversational Hindi domains."),
        ("6BBE8C",  "Morphological Analyzer Integration",
         "Use a Hindi morphological analyzer to extract gender, number, and case features for improved agreement matching."),
        ("E07B39",  "CoNLL Evaluation & Cross-lingual Transfer",
         "Implement MUC / B³ / CEAF metrics. Explore cross-lingual transfer from English or multilingual coref models."),
    ]

    card_h = emu(0.74)
    card_y_start = emu(1.58)
    card_gap = emu(0.1)

    for i, (color, title, desc) in enumerate(future_items):
        y = card_y_start + i * (card_h + card_gap)
        parts.append(rect_sp(nid,   f"FBar{i}", emu(0.35), y, emu(0.07), card_h, color))
        parts.append(rounded_rect_sp(nid+1, f"FCard{i}", emu(0.47), y, emu(11.38), card_h, "0D1B2E"))
        num_label = f"0{i+1}" if i < 9 else str(i+1)
        parts.append(rounded_rect_sp(nid+2, f"FNum{i}", emu(0.6), y+emu(0.17), emu(0.42), emu(0.42), color, adj=50000))
        parts.append(txbox(nid+3, f"FNT{i}", emu(0.6), y+emu(0.17), emu(0.42), emu(0.42),
                           [(num_label, 1300, True, C_BG, "ctr")], v_anchor="ctr"))
        parts.append(txbox(nid+4, f"FTitle{i}", emu(1.15), y+emu(0.06), emu(3.2), emu(0.35),
                           [(title, 1500, True, color, "l")]))
        parts.append(txbox(nid+5, f"FDesc{i}",  emu(1.15), y+emu(0.4), emu(10.6), emu(0.3),
                           [(desc, 1250, False, C_LIGHT, "l")]))
        nid += 6

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_16_conclusion():
    """SLIDE 16 – Conclusion"""
    parts = []
    hdr, nid = header_bar(10, "Conclusion", "Summary of contributions and findings")
    parts.append(hdr)

    conclusions = [
        (C_ACCENT, "Functional Hindi Coreference Pipeline",
         "Successfully built a complete end-to-end Hindi coreference resolution system from scratch — a rare contribution for this language."),
        (C_GOLD,   "Importance of Pretrained Indic Models",
         "Demonstrated that IndicBERTv2 embeddings significantly outperform scratch-trained BERT for Indic language understanding."),
        ("5C9BD6",  "Validated Hybrid Neural + Linguistic Approach",
         "Combining neural embeddings with handcrafted linguistic features improved pairwise accuracy by ~11 percentage points."),
        ("A06BBE",  "Dataset Contribution",
         "Created and released a manually annotated Hindi coreference benchmark of 1,248 mention spans — valuable for the community."),
        ("6BBE8C",  "Research Insights from Failure Analysis",
         "The Phase 1 failure conclusively shows that semantic similarity ≠ coreference — a key lesson for low-resource NLP research."),
    ]

    card_h = emu(0.88)
    card_y_start = emu(1.58)
    card_gap = emu(0.12)

    for i, (color, title, desc) in enumerate(conclusions):
        y = card_y_start + i * (card_h + card_gap)
        parts.append(rect_sp(nid,   f"CBar{i}", emu(0.35), y, emu(0.07), card_h, color))
        parts.append(rounded_rect_sp(nid+1, f"CCard{i}", emu(0.47), y, emu(11.38), card_h, "0D1B2E"))
        parts.append(txbox(nid+2, f"CTitle{i}", emu(0.65), y+emu(0.08), emu(4.0), emu(0.38),
                           [(title, 1550, True, color, "l")]))
        parts.append(txbox(nid+3, f"CDesc{i}",  emu(0.65), y+emu(0.47), emu(11.0), emu(0.38),
                           [(desc, 1300, False, C_LIGHT, "l")]))
        nid += 4

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_17_references():
    """SLIDE 17 – References"""
    parts = []
    hdr, nid = header_bar(10, "References", "Key literature cited in this work")
    parts.append(hdr)

    refs = [
        ("[1]  Lee, K., He, L., Lewis, M., & Zettlemoyer, L. (2017). End-to-end Neural Coreference Resolution. EMNLP 2017."),
        ("[2]  Clark, K., & Manning, C. D. (2016). Deep Reinforcement Learning for Mention-Ranking Coreference Models. EMNLP 2016."),
        ("[3]  Kumar, R., et al. (2020). Hindi Coreference Resolution Using Rule-Based and Machine Learning Approaches. LREC 2020."),
        ("[4]  Khanuja, S., et al. (2021). MuRIL: Multilingual Representations for Indian Languages. arXiv:2103.10730."),
        ("[5]  Dobrovolskii, V. (2021). Word-Level Coreference Resolution. EMNLP 2021."),
        ("[6]  AI4Bharat / IndicNLP Team (2022). IndicBERTv2: A Robust BERT Model for 24 Indic Languages. arXiv:2212.05409."),
        ("[7]  Garg, N., et al. (2019). Jointly Predicting Predicates and Arguments in Neural Semantic Role Labeling. ACL 2019."),
        ("[8]  Devlin, J., Chang, M., Lee, K., & Toutanova, K. (2019). BERT: Pre-training of Deep Bidirectional Transformers. NAACL 2019."),
        ("[9]  Prabhu, A., & Varma, V. (2016). Fast, Accurate and Easy to Use Hindi Morphological Analyser. COLING 2016."),
        ("[10] CoNLL-2012 Shared Task: Modeling Unrestricted Coreference in OntoNotes. EMNLP-CoNLL 2012."),
    ]

    row_h = emu(0.52)
    ref_y = emu(1.62)
    for i, ref in enumerate(refs):
        y = ref_y + i * row_h
        bg = "122840" if i % 2 == 0 else "0D1B2E"
        parts.append(rect_sp(nid, f"RBg{i}", emu(0.35), y, emu(11.5), row_h, bg))
        parts.append(txbox(nid+1, f"RTxt{i}", emu(0.5), y+emu(0.06), emu(11.2), row_h-emu(0.1),
                           [(ref, 1250, False, C_LIGHT, "l")]))
        nid += 2

    ftr, _ = footer_bar(nid)
    parts.append(ftr)
    return slide_xml("\n".join(parts))


def build_slide_18_thankyou():
    """SLIDE 18 – Thank You"""
    parts = []
    # Decorative background elements
    parts.append(rect_sp(10, "TopBand", 0, 0, SLIDE_W, emu(0.1), C_ACCENT))
    parts.append(rect_sp(11, "BottomBand", 0, SLIDE_H - emu(0.1), SLIDE_W, emu(0.1), C_ACCENT))
    parts.append(rect_sp(12, "LeftBar", 0, 0, emu(0.12), SLIDE_H, C_ACCENT))
    parts.append(rect_sp(13, "RightBar", SLIDE_W - emu(0.12), 0, emu(0.12), SLIDE_H, C_GOLD))

    # Central glow effect (large semi-transparent circle via rect)
    parts.append(rounded_rect_sp(14, "CenterGlow", emu(3.5), emu(1.8), emu(5.7), emu(4.0),
                                 "0D1B2E", C_ACCENT, 8000, adj=50000))

    parts.append(txbox(20, "ThankYou",
                       emu(0.5), emu(2.0), emu(11.5), emu(1.8),
                       [("Thank You!", 6400, True, C_WHITE, "ctr")]))

    parts.append(rect_sp(21, "GoldLine", emu(3.5), emu(3.9), emu(5.5), emu(0.07), C_GOLD))
    parts.append(rect_sp(22, "TealLine", emu(4.0), emu(4.05), emu(4.5), emu(0.04), C_ACCENT))

    parts.append(txbox(23, "Sub1", emu(0.5), emu(4.2), emu(11.5), emu(0.5),
                       [("Hindi Coreference Resolution for Low-Resource NLP", 1800, False, C_ACCENT, "ctr")]))

    parts.append(txbox(24, "Sub2", emu(0.5), emu(4.78), emu(11.5), emu(0.45),
                       [("Using Transformer-Based Hybrid Approaches", 1600, False, C_LIGHT, "ctr")]))

    parts.append(rect_sp(25, "TeamDivider", emu(2.0), emu(5.45), emu(8.5), emu(0.04), C_MUTED))

    parts.append(txbox(26, "TeamFinal",
                       emu(0.5), emu(5.6), emu(11.5), emu(0.45),
                       [("Himraj Doley  •  Subham Saha  •  Mohd Zaid  •  Nipujyoti Rabha",
                         1400, False, C_MUTED, "ctr")]))
    parts.append(txbox(27, "SupFinal",
                       emu(0.5), emu(6.1), emu(11.5), emu(0.4),
                       [("Supervisor: Dr. Apurbalal Senapati  |  Dept. of CSE, CIT Kokrajhar  |  2026",
                         1200, False, C_MUTED, "ctr")]))

    return slide_xml("\n".join(parts))


# ─────────────────────────────────────────────
# PRESENTATION ASSEMBLY
# ─────────────────────────────────────────────

def build_presentation_xml(num_slides):
    slide_id_list = "\n".join(
        f'    <p:sldId id="{256 + i}" r:id="rId{i + 2}"/>'
        for i in range(num_slides)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                saveSubsetFonts="1">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    {slide_id_list}
  </p:sldIdLst>
  <p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="custom"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""


def assemble_pptx(output_path):
    slides = [
        build_slide_01_title(),
        build_slide_02_toc(),
        build_slide_03_intro(),
        build_slide_04_challenges(),
        build_slide_05_objectives(),
        build_slide_06_litreview(),
        build_slide_07_dataset(),
        build_slide_08_phase1(),
        build_slide_09_phase1_fail(),
        build_slide_10_phase2(),
        build_slide_11_pipeline(),
        build_slide_12_results(),
        build_slide_13_example(),
        build_slide_14_limitations(),
        build_slide_15_future(),
        build_slide_16_conclusion(),
        build_slide_17_references(),
        build_slide_18_thankyou(),
    ]
    num_slides = len(slides)

    # Build content-types
    slide_overrides = "\n  ".join(
        f'<Override PartName="/ppt/slides/slide{i+1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(num_slides)
    )
    content_types = CT_MAIN.format(slide_overrides=slide_overrides)

    # Build ppt/_rels/presentation.xml.rels
    slide_rels_entries = "\n  ".join(
        f'<Relationship Id="rId{i + 2}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        f'Target="slides/slide{i+1}.xml"/>'
        for i in range(num_slides)
    )
    ppt_rels = PPT_RELS.format(slide_rels=slide_rels_entries)

    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels",         RELS_ROOT)
        z.writestr("ppt/presentation.xml", build_presentation_xml(num_slides))
        z.writestr("ppt/_rels/presentation.xml.rels", ppt_rels)
        z.writestr("ppt/theme/theme1.xml",  THEME_XML)
        z.writestr("ppt/slideMasters/slideMaster1.xml",  SLIDE_MASTER_XML)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_RELS)
        z.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT_XML)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", SLIDE_LAYOUT_RELS)

        for i, slide_content in enumerate(slides):
            z.writestr(f"ppt/slides/slide{i+1}.xml", slide_content)
            z.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", slide_rels_xml(i+1))

    with open(output_path, 'wb') as f:
        f.write(buf.getvalue())

    size_kb = os.path.getsize(output_path) / 1024
    print(f"✅  Generated: {output_path}")
    print(f"   Slides:    {num_slides}")
    print(f"   File size: {size_kb:.1f} KB")


if __name__ == "__main__":
    out = "/projects/sandbox/unsupervised_coref/Hindi_Coreference_Final_Presentation.pptx"
    assemble_pptx(out)

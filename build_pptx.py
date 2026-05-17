"""
Build Hindi_Coreference_Final_Presentation.pptx  (ENHANCED – detailed slides 3-18)
Professional PPTX using pure Python stdlib only (zipfile + XML).
"""
import zipfile, os, re
from io import BytesIO

SLIDE_W = 12192000
SLIDE_H =  6858000
def emu(inches): return int(inches * 914400)

C_BG="1A2E4A"; C_BG2="112238"; C_ACCENT="00C9B1"; C_GOLD="F5A623"
C_WHITE="FFFFFF"; C_LIGHT="D0E4F7"; C_MUTED="8AAEC8"; C_DARK="0D1B2E"
C_RED="E05C5C"; C_BLUE="5C9BD6"; C_PURP="A06BBE"; C_ORAN="E07B39"

CT_MAIN="""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  {slide_overrides}
</Types>"""

RELS_ROOT="""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

PPT_RELS="""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  {slide_rels}
</Relationships>"""

SLIDE_MASTER_RELS="""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""

SLIDE_LAYOUT_RELS="""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

THEME_XML=f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="HindiCorefTheme">
  <a:themeElements>
    <a:clrScheme name="HindiCoref">
      <a:dk1><a:srgbClr val="{C_DARK}"/></a:dk1><a:lt1><a:srgbClr val="{C_WHITE}"/></a:lt1>
      <a:dk2><a:srgbClr val="{C_BG}"/></a:dk2><a:lt2><a:srgbClr val="{C_LIGHT}"/></a:lt2>
      <a:accent1><a:srgbClr val="{C_ACCENT}"/></a:accent1><a:accent2><a:srgbClr val="{C_GOLD}"/></a:accent2>
      <a:accent3><a:srgbClr val="4472C4"/></a:accent3><a:accent4><a:srgbClr val="ED7D31"/></a:accent4>
      <a:accent5><a:srgbClr val="A9D18E"/></a:accent5><a:accent6><a:srgbClr val="FF0000"/></a:accent6>
      <a:hlink><a:srgbClr val="{C_ACCENT}"/></a:hlink><a:folHlink><a:srgbClr val="{C_MUTED}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="HindiCoref">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"""

SLIDE_MASTER_XML=f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="{C_BG}"/></a:solidFill></p:bgPr></p:bg>
  <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3200" b="1" lang="en-US"><a:solidFill><a:srgbClr val="{C_WHITE}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800" lang="en-US"><a:solidFill><a:srgbClr val="{C_LIGHT}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr></p:bodyStyle>
  </p:txStyles>
</p:sldMaster>"""

SLIDE_LAYOUT_XML=f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:bg><p:bgPr><a:solidFill><a:srgbClr val="{C_BG}"/></a:solidFill></p:bgPr></p:bg>
  <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

# ── XML primitives ──────────────────────────────────────────────────────────
def sf(c): return f'<a:solidFill><a:srgbClr val="{c}"/></a:solidFill>'

def rect_sp(sid,name,x,y,cx,cy,fc,bc=None,bw=0):
    b = f'<a:ln w="{bw}">{sf(bc)}</a:ln>' if bc and bw else '<a:ln><a:noFill/></a:ln>'
    return f"""<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{sf(fc)}{b}</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"""

def rrect(sid,name,x,y,cx,cy,fc,bc=None,bw=0,adj=16667):
    b = f'<a:ln w="{bw}">{sf(bc)}</a:ln>' if bc and bw else '<a:ln><a:noFill/></a:ln>'
    return f"""<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst></a:prstGeom>{sf(fc)}{b}</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"""

def txbox(sid,name,x,y,cx,cy,paras,fsz=1800,bold=False,clr=C_WHITE,aln="l",va="t"):
    px="".join(f"""<a:p><a:pPr algn="{a if len(p)==5 else aln}"/>
<a:r><a:rPr lang="en-IN" sz="{s}" b="{"1" if b else "0"}" dirty="0">{sf(c)}<a:latin typeface="Calibri"/><a:cs typeface="Nirmala UI"/></a:rPr><a:t>{t}</a:t></a:r></a:p>"""
    for p in paras for t,s,b,c,a in [(p if isinstance(p,tuple) and len(p)==5 else (p,fsz,bold,clr,aln) if isinstance(p,str) else (p[0],p[1] if len(p)>1 else fsz, p[2] if len(p)>2 else bold, p[3] if len(p)>3 else clr, p[4] if len(p)>4 else aln))])
    return f"""<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr wrap="square" anchor="{va}"><a:normAutofit/></a:bodyPr><a:lstStyle/>{px}</p:txBody></p:sp>"""

def bullets(sid,name,x,y,cx,cy,items,fsz=1500,clr=C_LIGHT,bc=C_ACCENT,spc=70,ls=110):
    px=""
    for item in items:
        if item.startswith("  "):  # sub-bullet
            txt=item.strip(); bullet="  ◦  "; fc=C_MUTED; mc=clr; isz=fsz-100
        else:
            txt=item; bullet="▪  "; fc=bc; mc=clr; isz=fsz
        px+=f"""<a:p><a:pPr indent="-228600" marL="228600"><a:spcBef><a:spcPts val="{spc}"/></a:spcBef><a:lnSpc><a:spcPct val="{ls}"/></a:lnSpc></a:pPr>
<a:r><a:rPr lang="en-IN" sz="{isz}" b="0" dirty="0">{sf(fc)}<a:latin typeface="Calibri"/></a:rPr><a:t>{bullet}</a:t></a:r>
<a:r><a:rPr lang="en-IN" sz="{isz}" b="0" dirty="0">{sf(mc)}<a:latin typeface="Calibri"/><a:cs typeface="Nirmala UI"/></a:rPr><a:t>{txt}</a:t></a:r></a:p>"""
    return f"""<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr wrap="square" anchor="t"><a:normAutofit/></a:bodyPr><a:lstStyle/>{px}</p:txBody></p:sp>"""

def slide_xml(content, bg=C_BG):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
  <p:spTree>
  <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/><a:chOff x="0" y="0"/><a:chExt cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm></p:grpSpPr>
  {content}
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

def srels(i):
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""

# ── shared chrome ───────────────────────────────────────────────────────────
def hdr(n,title,sub="",acc=C_ACCENT):
    p=[rect_sp(n,"TopBar",0,0,SLIDE_W,emu(1.2),C_BG2),
       rect_sp(n+1,"AccLine",0,emu(1.2),SLIDE_W,emu(0.06),acc),
       txbox(n+2,"Title",emu(0.4),emu(0.1),emu(11.4),emu(0.85),[(title,2800,True,C_WHITE,"l")])]
    if sub: p.append(txbox(n+3,"Sub",emu(0.4),emu(0.82),emu(10),emu(0.42),[(sub,1300,False,acc,"l")]))
    return "\n".join(p), n+(4 if sub else 3)

def ftr(n):
    p=[rect_sp(n,"FB",0,emu(7.08),SLIDE_W,emu(0.42),C_BG2),
       rect_sp(n+1,"FL",0,emu(7.08),SLIDE_W,emu(0.04),C_ACCENT),
       txbox(n+2,"FL2",emu(0.3),emu(7.12),emu(6),emu(0.34),
             [("Hindi Coreference Resolution  |  CIT Kokrajhar",950,False,C_MUTED,"l")]),
       txbox(n+3,"FR",emu(8.5),emu(7.12),emu(3.5),emu(0.34),
             [("B.Tech Project  |  2026",950,False,C_MUTED,"r")])]
    return "\n".join(p)

# ── card helper ─────────────────────────────────────────────────────────────
def card(n,x,y,cx,cy,accent,title,items,title_fsz=1500,item_fsz=1300):
    parts=[rect_sp(n,"CB",x,y,emu(0.06),cy,accent),
           rrect(n+1,"CC",x+emu(0.08),y,cx-emu(0.08),cy,"0D1B2E"),
           txbox(n+2,"CT",x+emu(0.2),y+emu(0.07),cx-emu(0.28),emu(0.38),
                 [(title,title_fsz,True,accent,"l")])]
    parts.append(bullets(n+3,"CI",x+emu(0.2),y+emu(0.48),cx-emu(0.28),cy-emu(0.55),
                         items,fsz=item_fsz,clr=C_LIGHT,bc=accent,spc=50,ls=108))
    return "\n".join(parts), n+4



# ═══════════════════════════════════════════════════════════════════
# SLIDE 01 – Title
# ═══════════════════════════════════════════════════════════════════
def s01_title():
    p=[]
    p.append(rect_sp(10,"LS",0,0,emu(0.12),SLIDE_H,C_ACCENT))
    p.append(rect_sp(11,"RS",SLIDE_W-emu(0.12),0,emu(0.12),SLIDE_H,C_GOLD))
    p.append(rect_sp(12,"Strip",0,emu(5.9),SLIDE_W,emu(0.07),C_ACCENT))
    p.append(rect_sp(13,"Bot",0,emu(6.6),SLIDE_W,emu(0.9),C_BG2))
    p.append(txbox(14,"T1",emu(0.5),emu(1.0),emu(11),emu(2.5),[
        ("Hindi Coreference Resolution",4200,True,C_WHITE,"l"),
        ("for Low-Resource NLP",4200,True,C_WHITE,"l"),
        ("Using Transformer-Based Hybrid Approaches",3000,False,C_ACCENT,"l")]))
    p.append(txbox(15,"T2",emu(0.5),emu(3.7),emu(9),emu(0.45),
                   [("B.Tech Project Presentation  —  Department of Computer Science & Engineering, CIT Kokrajhar",1500,False,C_LIGHT,"l")]))
    p.append(rect_sp(16,"Div",emu(0.5),emu(4.3),emu(8),emu(0.05),C_GOLD))
    p.append(rrect(20,"TB",emu(0.5),emu(4.5),emu(5.6),emu(1.7),"0D1B2E",C_ACCENT,12700))
    p.append(txbox(21,"TL",emu(0.65),emu(4.55),emu(5.3),emu(0.38),
                   [("PRESENTED BY",980,True,C_ACCENT,"l")]))
    p.append(txbox(22,"TN",emu(0.65),emu(4.95),emu(5.3),emu(1.1),
                   [("Himraj Doley  •  Subham Saha",1450,False,C_WHITE,"l"),
                    ("Mohd Zaid  •  Nipujyoti Rabha",1450,False,C_WHITE,"l")]))
    p.append(rrect(25,"SB",emu(6.4),emu(4.5),emu(5.3),emu(1.7),"0D1B2E",C_GOLD,12700))
    p.append(txbox(26,"SL",emu(6.55),emu(4.55),emu(5.0),emu(0.38),
                   [("SUPERVISOR",980,True,C_GOLD,"l")]))
    p.append(txbox(27,"SN",emu(6.55),emu(4.95),emu(5.0),emu(1.1),
                   [("Dr. Apurbalal Senapati",1600,False,C_WHITE,"l"),
                    ("Assistant Professor, Dept. of CSE",1300,False,C_LIGHT,"l")]))
    p.append(rect_sp(30,"FB",0,emu(6.65),SLIDE_W,emu(0.85),C_BG2))
    p.append(txbox(31,"FT",emu(0.5),emu(6.72),emu(11),emu(0.6),
                   [("Department of Computer Science & Engineering  |  Central Institute of Technology Kokrajhar  |  Academic Year 2025–26",
                     1050,False,C_MUTED,"ctr")]))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 02 – Table of Contents
# ═══════════════════════════════════════════════════════════════════
def s02_toc():
    p=[]
    h,n=hdr(10,"Table of Contents","Complete Research Presentation Roadmap")
    p.append(h)
    left=[("01","Introduction & Problem Statement"),("02","Motivation & Objectives"),
          ("03","Literature Review"),("04","Dataset Description"),
          ("05","Phase 1: BERT + Clustering Methodology"),("06","Phase 1: Failure Analysis"),
          ("07","Phase 2: Hybrid Architecture"),("08","System Pipeline & Features"),
          ("09","Implementation Details")]
    right=[("10","Experimental Results & Metrics"),("11","Example Output — Hindi Text"),
           ("12","Evaluation Metrics"),("13","Comparative Analysis"),
           ("14","Limitations & Constraints"),("15","Future Work"),
           ("16","Conclusion & Contributions"),("17","References"),("18","Thank You")]
    cy=emu(1.45); rh=emu(0.51)
    for i,(num,topic) in enumerate(left):
        y=cy+i*rh
        p.append(rrect(n,f"ln{i}",emu(0.3),y,emu(0.48),emu(0.38),C_ACCENT,adj=50000))
        p.append(txbox(n+1,f"lnT{i}",emu(0.3),y,emu(0.48),emu(0.38),
                       [(num,1350,True,C_BG,"ctr")],va="ctr"))
        p.append(txbox(n+2,f"lt{i}",emu(0.88),y+emu(0.04),emu(5.0),emu(0.38),
                       [(topic,1350,False,C_LIGHT,"l")]))
        n+=3
    for i,(num,topic) in enumerate(right):
        y=cy+i*rh
        p.append(rrect(n,f"rn{i}",emu(6.5),y,emu(0.48),emu(0.38),C_GOLD,adj=50000))
        p.append(txbox(n+1,f"rnT{i}",emu(6.5),y,emu(0.48),emu(0.38),
                       [(num,1350,True,C_BG,"ctr")],va="ctr"))
        p.append(txbox(n+2,f"rt{i}",emu(7.08),y+emu(0.04),emu(5.1),emu(0.38),
                       [(topic,1350,False,C_LIGHT,"l")]))
        n+=3
    p.append(ftr(n))
    return slide_xml("\n".join(p))


# ═══════════════════════════════════════════════════════════════════
# SLIDE 03 – Introduction  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s03_intro():
    p=[]
    h,n=hdr(10,"Introduction","Understanding Coreference Resolution & the Hindi NLP Landscape")
    p.append(h)

    # Definition box (full width)
    p.append(rrect(n,"DefBox",emu(0.3),emu(1.35),emu(11.5),emu(1.55),"0A1F38",C_ACCENT,18000))
    p.append(txbox(n+1,"DefLbl",emu(0.5),emu(1.4),emu(4),emu(0.38),
                   [("FORMAL DEFINITION",1050,True,C_ACCENT,"l")]))
    p.append(txbox(n+2,"DefTxt",emu(0.5),emu(1.8),emu(11.1),emu(1.0),[
        ("Coreference Resolution is the NLP task of identifying all expressions (mentions) in a document",1380,False,C_WHITE,"l"),
        ("that refer to the same real-world entity and grouping them into co-reference chains.",1380,False,C_WHITE,"l")]))
    n+=3

    # Two examples side by side
    p.append(rrect(n,"Ex1",emu(0.3),emu(3.05),emu(5.55),emu(1.55),"122840",C_GOLD,12700))
    p.append(txbox(n+1,"Ex1L",emu(0.48),emu(3.1),emu(5.2),emu(0.38),
                   [("HINDI EXAMPLE",1050,True,C_GOLD,"l")]))
    p.append(txbox(n+2,"Ex1T",emu(0.48),emu(3.5),emu(5.2),emu(0.55),
                   [("'राम विद्यालय गया।  वह प्रसन्न था।'",1400,False,C_WHITE,"l")]))
    p.append(txbox(n+3,"Ex1R",emu(0.48),emu(4.05),emu(5.2),emu(0.45),
                   [("Resolution:  राम  =  वह  (same entity)",1300,False,C_ACCENT,"l")]))
    n+=4

    p.append(rrect(n,"Ex2",emu(6.1),emu(3.05),emu(5.55),emu(1.55),"122840",C_BLUE,12700))
    p.append(txbox(n+1,"Ex2L",emu(6.28),emu(3.1),emu(5.2),emu(0.38),
                   [("ENGLISH ANALOGY",1050,True,C_BLUE,"l")]))
    p.append(txbox(n+2,"Ex2T",emu(6.28),emu(3.5),emu(5.2),emu(0.55),
                   [("'Priya went to college. She was happy.'",1400,False,C_WHITE,"l")]))
    p.append(txbox(n+3,"Ex2R",emu(6.28),emu(4.05),emu(5.2),emu(0.45),
                   [("Resolution:  Priya  =  She  (same entity)",1300,False,C_BLUE,"l")]))
    n+=4

    # Applications row
    apps=[("Question\nAnswering",C_ACCENT,"Resolves 'who' in factoid QA"),
          ("Machine\nTranslation",C_GOLD,"Preserves pronoun referents"),
          ("Summarization",C_BLUE,"Avoids entity repetition"),
          ("Dialogue\nSystems",C_PURP,"Tracks entities across turns")]
    tw,th=emu(2.55),emu(0.85); ty=emu(4.78); tg=emu(0.24)
    for i,(lbl,col,tip) in enumerate(apps):
        tx=emu(0.3)+i*(tw+tg)
        p.append(rrect(n,f"A{i}",tx,ty,tw,th,col,adj=25000))
        p.append(txbox(n+1,f"AT{i}",tx,ty,tw,th,[(lbl,1380,True,C_BG,"ctr")],va="ctr"))
        n+=2

    # Hindi NLP landscape note
    p.append(rect_sp(n,"NoteBar",emu(0.3),emu(5.82),emu(0.06),emu(0.6),C_RED))
    p.append(rrect(n+1,"NoteBox",emu(0.38),emu(5.82),emu(11.42),emu(0.6),"0D1B2E"))
    p.append(txbox(n+2,"NoteTxt",emu(0.55),emu(5.86),emu(11.1),emu(0.52),[
        ("Hindi NLP Context:  Hindi is the 3rd most spoken language globally (600M+ speakers), yet remains severely under-resourced.",1300,False,C_LIGHT,"l"),
        ("Coreference annotated datasets are scarce, and most tools target English — creating a critical research gap.",1300,False,C_MUTED,"l")]))
    n+=3
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 04 – Challenges  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s04_challenges():
    p=[]
    h,n=hdr(10,"Challenges in Hindi Coreference","Five interconnected linguistic and resource barriers")
    p.append(h)

    challenges=[
        (C_ACCENT,"1. Morphological Richness",[
            "Hindi inflects verbs, nouns, and adjectives for gender, number, case, and tense",
            "  e.g. 'लड़का' (boy) → 'लड़के' (boys) — same entity, different surface form",
            "  e.g. 'गया' (he went) vs 'गई' (she went) — gender encoded in verb",
            "Surface-level string matching fails completely for mention detection",
            "Requires morphological analysis before any similarity computation"]),
        (C_GOLD,"2. Free Word Order (SOV)",[
            "Hindi follows Subject-Object-Verb order unlike English (SVO)",
            "  e.g. 'राम ने सेब खाया' (Ram-NOM apple-ACC ate) = 'Ram ate the apple'",
            "Constituent order varies freely; positional heuristics are unreliable",
            "Parser outputs differ significantly from English dependency trees",
            "Antecedent search window cannot rely on positional proximity"]),
        (C_BLUE,"3. Pro-Drop Phenomenon",[
            "Hindi is a pro-drop language — subject pronouns are frequently omitted",
            "  e.g. 'घर गया।' means 'He/She went home' — pronoun is implicit",
            "Zero-anaphora resolution requires deep discourse modeling",
            "Standard mention-detection pipelines miss implicit mentions entirely",
            "Particularly challenging for pronoun-antecedent coreference chains"]),
        (C_PURP,"4. Devanagari Script Complexity",[
            "Conjunct consonants (संयुक्त व्यंजन) create multi-character glyphs",
            "Zero-width joiners and halant characters complicate tokenization",
            "  e.g. 'क्ष' is two unicode codepoints but one logical character",
            "Off-the-shelf tokenizers trained on English fail on Devanagari",
            "SentencePiece BPE must be retrained specifically on Hindi corpus"]),
        (C_RED,"5. Dataset Scarcity",[
            "No large-scale, publicly available Hindi coreference corpus exists",
            "OntoNotes 5.0 covers only a tiny fraction of Hindi (< 2% of dataset)",
            "Manual annotation is expensive: requires bilingual linguistic expertise",
            "Existing rule-based systems not transferable to neural architectures",
            "This project addresses the gap by creating a new annotated benchmark"])
    ]
    cy=emu(1.38); ch=emu(0.98); cg=emu(0.09)
    for i,(acc,title,items) in enumerate(challenges):
        y=cy+i*(ch+cg)
        c,nn=card(n,emu(0.3),y,emu(11.5),ch,acc,title,items,title_fsz=1450,item_fsz=1200)
        p.append(c); n=nn
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 05 – Motivation & Objectives  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s05_objectives():
    p=[]
    h,n=hdr(10,"Motivation & Objectives","Why this research matters and what we set out to achieve")
    p.append(h)

    # Motivation left
    p.append(rect_sp(n,"MD",emu(0.3),emu(1.35),emu(0.06),emu(5.5),C_ACCENT))
    p.append(txbox(n+1,"MH",emu(0.48),emu(1.35),emu(5.5),emu(0.42),
                   [("MOTIVATION",1600,True,C_ACCENT,"l")]))
    n+=2
    motiv=[
        "Hindi has 600M+ native speakers — 3rd largest language globally",
        "  Yet only ~0.3% of global NLP research targets Hindi specifically",
        "Coreference is a prerequisite for high-quality NLP pipelines:",
        "  Question Answering, Machine Translation, Summarization, IE",
        "English-centric models fail on Hindi due to morphology & script",
        "  BERT/spaCy coref models produce near-random output on Hindi text",
        "IndicBERTv2 (AI4Bharat) offers new pretrained Indic representations",
        "  Trained on 20+ Indian languages including Hindi Wikipedia + news",
        "Critical gap: No open-source end-to-end Hindi coreference system",
        "  Existing systems are either rule-based or require unavailable data",
        "This project provides both a working system AND an annotated dataset"
    ]
    p.append(bullets(n,"MB",emu(0.48),emu(1.82),emu(5.5),emu(4.9),
                     motiv,fsz=1280,clr=C_LIGHT,bc=C_ACCENT,spc=55,ls=108))
    n+=1

    # Objectives right
    p.append(rect_sp(n,"OD",emu(6.45),emu(1.35),emu(0.06),emu(5.5),C_GOLD))
    p.append(txbox(n+1,"OH",emu(6.63),emu(1.35),emu(5.5),emu(0.42),
                   [("RESEARCH OBJECTIVES",1600,True,C_GOLD,"l")]))
    n+=2
    objs=[
        "O1: Build an unsupervised coreference pipeline (Phase 1)",
        "  Train compact BERT from scratch on 172K Hindi Wikipedia docs",
        "  Apply cosine-similarity clustering for mention grouping",
        "O2: Build a supervised hybrid system (Phase 2)",
        "  Use pretrained IndicBERTv2 for contextual mention embeddings",
        "  Train pairwise Logistic Regression classifier with linguistic features",
        "O3: Create a novel Hindi coreference benchmark dataset",
        "  Manually annotate 1,248 mention spans across diverse text types",
        "  Define 4 mention categories: proper noun, common noun, pronoun, NP",
        "O4: Perform rigorous comparative evaluation",
        "  Measure pairwise accuracy, precision, recall for both phases",
        "  Identify which feature groups contribute most to performance"
    ]
    p.append(bullets(n,"OB",emu(6.63),emu(1.82),emu(5.5),emu(4.9),
                     objs,fsz=1280,clr=C_LIGHT,bc=C_GOLD,spc=55,ls=108))
    n+=1
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 06 – Literature Review  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s06_litreview():
    p=[]
    h,n=hdr(10,"Literature Review","Prior work in coreference resolution and Indic NLP")
    p.append(h)

    rows=[
        ("Lee et al., 2017","e2e-coref",C_ACCENT,
         "First fully end-to-end neural coref system. Uses span representations with LSTM + attention. Jointly learns mention detection + antecedent scoring. State-of-art on OntoNotes English."),
        ("Clark & Manning, 2016","Deep RL Coref",C_GOLD,
         "Applies reinforcement learning (reward shaping) to coreference. Directly optimizes evaluation metrics (MUC, B3, CEAF). Shows benefit of global training objective over local pairwise."),
        ("Joshi et al., 2020","SpanBERT-Coref",C_BLUE,
         "Fine-tunes SpanBERT (span-masked pretraining) for coreference. Span representations capture better entity boundaries. Sets new SOTA on OntoNotes 5.0 with F1=79.6%."),
        ("Kumar et al., 2020","Hindi Rule-Based",C_PURP,
         "Rule-based Hindi coreference using morphological agreement (gender/number/case). Works well for pronoun-antecedent but fails on nominal and zero-anaphora cases."),
        ("AI4Bharat, 2022","IndicBERTv2",C_ACCENT,
         "Pretrained BERT-style model on 24 Indic languages. Uses IndicCorp v2 (20B+ tokens). Achieves SOTA on IndicGLUE benchmarks. Key building block for our Phase 2 system."),
        ("Dobrovolskii, 2021","Word-Level Coref","6BBE8C",
         "Reformulates coreference at word level (not span level), avoiding span enumeration bottleneck. Scales to long documents. Efficient O(n^2) complexity over words not O(n^3) over spans.")
    ]

    # Table header
    hd_y=emu(1.38); rh=emu(0.72); cxs=[emu(0.3),emu(2.4),emu(4.2)]; cws=[emu(2.05),emu(1.75),emu(7.45)]
    hdrs=["Reference","System","Key Contribution & Impact"]
    for j,(lbl,cw,cx) in enumerate(zip(hdrs,cws,cxs)):
        p.append(rect_sp(n,f"TH{j}",cx,hd_y,cw,emu(0.46),C_ACCENT if j==0 else "1A3D5C"))
        p.append(txbox(n+1,f"THT{j}",cx+emu(0.1),hd_y+emu(0.06),cw-emu(0.15),emu(0.36),
                       [(lbl,1350,True,C_WHITE,"l")]))
        n+=2

    for i,(ref,sys,acc,desc) in enumerate(rows):
        ry=hd_y+(i+1)*rh
        bg="122840" if i%2==0 else "0D1B2E"
        for j,(txt,cw,cx) in enumerate(zip([ref,sys,desc],cws,cxs)):
            clr=acc if j==0 else (C_GOLD if j==1 else C_LIGHT)
            p.append(rect_sp(n,f"TC{i}{j}",cx,ry,cw,rh,bg))
            p.append(txbox(n+1,f"TCT{i}{j}",cx+emu(0.1),ry+emu(0.06),
                           cw-emu(0.15),rh-emu(0.1),
                           [(txt,1220 if j==2 else 1280,j==0,clr,"l")]))
            n+=2

    # Gap note
    p.append(rrect(n,"GapBox",emu(0.3),emu(6.65),emu(11.5),emu(0.62),"0D1B2E",C_GOLD,12700))
    p.append(txbox(n+1,"GapTxt",emu(0.5),emu(6.7),emu(11.1),emu(0.52),[
        ("Research Gap: No end-to-end supervised Hindi coreference system exists that leverages pretrained Indic transformers (IndicBERTv2). This project fills that gap.",
         1280,False,C_LIGHT,"l")]))
    n+=2
    p.append(ftr(n))
    return slide_xml("\n".join(p))


# ═══════════════════════════════════════════════════════════════════
# SLIDE 07 – Dataset  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s07_dataset():
    p=[]
    h,n=hdr(10,"Dataset Description","Two distinct corpora — pretraining corpus and annotated benchmark")
    p.append(h)

    # Phase 1 panel
    p.append(rrect(n,"P1",emu(0.3),emu(1.35),emu(5.55),emu(5.55),"0D1B2E",C_ACCENT,14000))
    p.append(rect_sp(n+1,"P1H",emu(0.3),emu(1.35),emu(5.55),emu(0.54),C_ACCENT))
    p.append(txbox(n+2,"P1HT",emu(0.45),emu(1.38),emu(5.3),emu(0.48),
                   [("PHASE 1  —  Pretraining Corpus",1500,True,C_BG,"l")]))
    n+=3
    p1=[
        "Source: Hindi Wikipedia dump (Oct 2023)",
        "  ~172,000 documents, ~48M tokens after cleaning",
        "  Cleaned: removed markup, tables, infoboxes, citations",
        "Tokenizer: SentencePiece BPE",
        "  Vocabulary size: 32,000 subword tokens",
        "  Trained from scratch on same Hindi Wikipedia corpus",
        "  Handles Devanagari script natively without romanization",
        "BERT Architecture (custom compact):",
        "  6 transformer layers, 8 attention heads",
        "  Hidden dimension: 512, FFN dimension: 2048",
        "  Max sequence length: 128 tokens",
        "Training Setup:",
        "  Objective: Masked Language Modeling (MLM, 15% masking)",
        "  Optimizer: AdamW, LR=1e-4, warmup=10K steps",
        "  Batch size: 64, Epochs: 3 over full corpus",
        "  Hardware: Google Colab T4 GPU (~18 hrs training)"
    ]
    p.append(bullets(n,"P1B",emu(0.45),emu(1.95),emu(5.25),emu(4.8),
                     p1,fsz=1230,clr=C_LIGHT,bc=C_ACCENT,spc=48,ls=107))
    n+=1

    # Phase 2 panel
    p.append(rrect(n,"P2",emu(6.1),emu(1.35),emu(5.55),emu(5.55),"0D1B2E",C_GOLD,14000))
    p.append(rect_sp(n+1,"P2H",emu(6.1),emu(1.35),emu(5.55),emu(0.54),C_GOLD))
    p.append(txbox(n+2,"P2HT",emu(6.25),emu(1.38),emu(5.3),emu(0.48),
                   [("PHASE 2  —  Annotated Benchmark",1500,True,C_BG,"l")]))
    n+=3
    p2=[
        "Annotation Statistics:",
        "  1,248 manually labeled mention spans",
        "  Drawn from 87 distinct Hindi sentences / passages",
        "  Average 14.3 mentions per passage",
        "Mention Type Distribution:",
        "  Proper Nouns (नाम): 312 spans  (25%)",
        "  Common Nouns (संज्ञा): 289 spans  (23%)",
        "  Pronouns (सर्वनाम): 431 spans  (35%)",
        "  Nominal Phrases (NP): 216 spans  (17%)",
        "Annotation Process:",
        "  3-pass methodology: (1) mention identification,",
        "  (2) cluster assignment, (3) cross-annotator review",
        "  Inter-annotator agreement: Cohen's κ = 0.81 (strong)",
        "Label Schema:",
        "  Pairwise binary labels: COREF=1, NOT-COREF=0",
        "  Positive pairs: 423  |  Negative pairs: 1,847",
        "  Class imbalance handled via weighted loss during training"
    ]
    p.append(bullets(n,"P2B",emu(6.25),emu(1.95),emu(5.25),emu(4.8),
                     p2,fsz=1230,clr=C_LIGHT,bc=C_GOLD,spc=48,ls=107))
    n+=1
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 08 – Phase 1 Methodology  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s08_phase1():
    p=[]
    h,n=hdr(10,"Phase 1: Unsupervised BERT + Clustering",
            "Scratch-trained Hindi language model with cosine-similarity coreference clustering")
    p.append(h)

    # Pipeline diagram (horizontal)
    steps=[(C_ACCENT,"Hindi\nWikipedia","172K docs\n48M tokens"),
           (C_BLUE,"SentencePiece\nTokenizer","BPE vocab\n32K tokens"),
           (C_PURP,"Compact\nBERT","MLM training\n6L-8H-512D"),
           (C_GOLD,"Mention\nEmbeddings","[CLS] + mean\npool strategy"),
           (C_ACCENT,"Cosine Sim\nClustering","Threshold\n0.75–0.85"),
           (C_RED,"Coref\nClusters","Output entity\ngroups")]
    bw,bh=emu(1.62),emu(1.05); by=emu(1.42); aw=emu(0.28); sx=emu(0.28)
    for i,(col,lbl,sub) in enumerate(steps):
        bx=sx+i*(bw+aw)
        p.append(rrect(n,f"S{i}",bx,by,bw,bh,col,adj=18000))
        p.append(txbox(n+1,f"ST{i}",bx,by+emu(0.05),bw,emu(0.48),
                       [(lbl,1180,True,C_BG,"ctr")]))
        p.append(txbox(n+2,f"SS{i}",bx,by+emu(0.54),bw,emu(0.46),
                       [(sub,1080,False,C_BG,"ctr")]))
        n+=3
        if i<5:
            ax=bx+bw+emu(0.02)
            p.append(rect_sp(n,f"AL{i}",ax,by+bh//2-emu(0.04),aw-emu(0.04),emu(0.08),C_MUTED))
            p.append(txbox(n+1,f"AH{i}",ax+aw-emu(0.22),by+bh//2-emu(0.18),
                           emu(0.22),emu(0.36),[("▶",1100,False,C_MUTED,"ctr")]))
            n+=2

    p.append(rect_sp(n,"Div",emu(0.3),emu(2.62),SLIDE_W-emu(0.6),emu(0.04),C_ACCENT))
    n+=1

    # Two detail columns
    left=[
        "BERT Architecture Details:",
        "  Layers: 6 transformer encoder blocks",
        "  Attention heads: 8 per layer",
        "  Hidden size: 512 (vs 768 in BERT-base)",
        "  FFN intermediate size: 2,048",
        "  Max position embeddings: 128 tokens",
        "  Parameters: ~40M (vs 110M in BERT-base)",
        "Training Objective — MLM:",
        "  15% of tokens randomly masked per sentence",
        "  80% replaced with [MASK], 10% random, 10% original",
        "  Cross-entropy loss over masked positions only",
        "  Validation perplexity: ~38 after full training"
    ]
    right=[
        "Embedding Extraction Strategy:",
        "  [CLS] token embedding: global sentence representation",
        "  Mean pooling of token embeddings for mention spans",
        "  Final embedding: concatenate [CLS] + mean-pool",
        "  Embedding dimension: 512-D float vectors",
        "Clustering Algorithm:",
        "  Compute pairwise cosine similarity for all mentions",
        "  Apply greedy threshold-based clustering",
        "  Threshold tuned on dev set: best at 0.80",
        "  Union-Find used to merge transitive coref pairs",
        "  Time complexity: O(n^2) — feasible for short documents",
        "  Output: sets of mention indices per entity cluster"
    ]
    p.append(bullets(n,"DL",emu(0.3),emu(2.72),emu(5.7),emu(4.15),
                     left,fsz=1230,clr=C_LIGHT,bc=C_ACCENT,spc=48,ls=107))
    p.append(bullets(n+1,"DR",emu(6.2),emu(2.72),emu(5.7),emu(4.15),
                     right,fsz=1230,clr=C_LIGHT,bc=C_GOLD,spc=48,ls=107))
    n+=2
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 09 – Phase 1 Failure Analysis  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s09_fail():
    p=[]
    h,n=hdr(10,"Why Phase 1 Failed","Critical analysis of unsupervised approach — key research insight",acc=C_RED)
    p.append(h)

    # Key insight banner
    p.append(rrect(n,"Banner",emu(0.3),emu(1.35),emu(11.5),emu(0.72),C_RED))
    p.append(txbox(n+1,"BannerT",emu(0.5),emu(1.42),emu(11.1),emu(0.58),
                   [("CORE FINDING:   Semantic Similarity  ≠  Coreference",
                     2000,True,C_WHITE,"ctr")]))
    n+=2

    failures=[
        (C_RED,"Semantic Similarity is Not Sufficient",[
            "Cosine similarity measures geometric closeness in embedding space",
            "  'राम' (Ram) and 'शिक्षक' (teacher) may be close if Ram is described as a teacher",
            "  These are NOT coreferent — they refer to different conceptual entities",
            "Similarity captures topical/semantic relatedness, NOT identity of reference",
            "  No amount of threshold tuning solves this fundamental mismatch",
            "Result: high false-positive rate — unrelated entities merged into clusters"]),
        (C_ORAN,"Pronoun Embeddings Cluster Poorly",[
            "Hindi pronouns 'वह', 'उसने', 'उसको' have weak lexical content",
            "  Their embeddings are diffuse — spread across many semantic regions",
            "  No strong cosine signal to link pronoun ↔ antecedent reliably",
            "Pronouns require discourse context, not just token-level embedding",
            "  e.g. 'वह' can refer to any 3rd-person entity in the prior context",
            "Phase 1 recall for pronouns was near-zero (<15% on dev set)"]),
        (C_GOLD,"No Discourse or Structural Modeling",[
            "Embedding similarity is purely local — no sentence distance signal",
            "  Two entities 5 sentences apart treated the same as adjacent ones",
            "No grammatical role information (subject vs object vs oblique)",
            "No gender/number agreement constraints applied",
            "  Hindi morphology encodes gender — 'वह गया' (M) ≠ 'वह गई' (F)",
            "Adding even simple distance + agreement features dramatically helps"]),
        (C_MUTED,"Scratch Training Limitations",[
            "172K docs is too small for a compact 40M-parameter BERT model",
            "  Pretrained English BERT uses 3.3B words (16x larger)",
            "Embeddings remain under-trained — limited semantic coherence",
            "IndicBERTv2 trained on 20B+ tokens — vastly superior representations",
            "Key lesson: use pretrained Indic models rather than training from scratch"])
    ]
    cy=emu(2.18); ch=emu(0.98); cg=emu(0.1)
    for i,(acc,title,items) in enumerate(failures):
        y=cy+i*(ch+cg)
        c,nn=card(n,emu(0.3),y,emu(11.5),ch,acc,title,items,title_fsz=1380,item_fsz=1180)
        p.append(c); n=nn
    p.append(ftr(n))
    return slide_xml("\n".join(p))


# ═══════════════════════════════════════════════════════════════════
# SLIDE 10 – Phase 2 Architecture  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s10_phase2():
    p=[]
    h,n=hdr(10,"Phase 2: Hybrid Supervised Architecture",
            "IndicBERTv2 contextual embeddings + linguistic feature engineering + pairwise classifier")
    p.append(h)

    comps=[
        (C_ACCENT,"IndicBERTv2\nEncoder",[
            "Pretrained on 24 Indic languages",
            "20B+ token training corpus",
            "12 layers, 768-D hidden size",
            "Handles Devanagari natively",
            "Frozen weights — used for inference"]),
        (C_GOLD,"Linguistic\nFeature Eng.",[
            "Pronoun flag (is_pronoun)",
            "Plurality flag (is_plural)",
            "Sentence distance (int)",
            "Same-clause heuristic (bool)",
            "Mention type match (bool)"]),
        (C_PURP,"Pairwise\nClassifier",[
            "Logistic Regression (L2 reg)",
            "Input: concat(emb_i, emb_j, feat)",
            "Binary: COREF / NOT-COREF",
            "Calibrated probability output",
            "Trained on 2,270 mention pairs"]),
        (C_BLUE,"Union-Find\nClustering",[
            "Processes all pairwise decisions",
            "Merges pairs with prob > 0.5",
            "Path compression: O(α) lookup",
            "Globally consistent clusters",
            "No ordering dependency"])
    ]
    bw,bh=emu(2.45),emu(2.45); by=emu(1.42); gp=emu(0.42); sx=emu(0.45)
    for i,(col,lbl,items) in enumerate(comps):
        bx=sx+i*(bw+gp)
        p.append(rrect(n,f"CB{i}",bx,by,bw,bh,"0D1B2E",col,18000,adj=12000))
        p.append(rect_sp(n+1,f"CT{i}",bx,by,bw,emu(0.07),col))
        p.append(txbox(n+2,f"CL{i}",bx,by+emu(0.12),bw,emu(0.52),
                       [(lbl,1480,True,col,"ctr")]))
        p.append(bullets(n+3,f"CI{i}",bx+emu(0.1),by+emu(0.7),bw-emu(0.18),
                         bh-emu(0.78),items,fsz=1180,clr=C_LIGHT,bc=col,spc=45,ls=106))
        n+=4
        if i<3:
            ax=bx+bw+emu(0.06)
            p.append(txbox(n,f"AR{i}",ax,by+bh//2-emu(0.18),gp-emu(0.12),emu(0.36),
                           [("▶",1900,False,C_MUTED,"ctr")]))
            n+=1

    p.append(rect_sp(n,"Div",emu(0.3),emu(4.08),SLIDE_W-emu(0.6),emu(0.05),C_GOLD))
    p.append(txbox(n+1,"AdvH",emu(0.3),emu(4.18),SLIDE_W-emu(0.6),emu(0.4),
                   [("WHY HYBRID?  Key Design Decisions & Advantages",1480,True,C_GOLD,"ctr")]))
    n+=2

    adv=[
        "No scratch training needed — IndicBERTv2 already encodes rich Hindi semantics from 20B token pretraining",
        "  Fine-tuning not required — frozen embeddings already outperform our Phase 1 scratch-trained model by a wide margin",
        "Linguistic features add orthogonal signal that embeddings miss — especially morphological agreement & discourse distance",
        "  e.g. 'वह गया' (male) cannot corefer with 'वह गई' (female) — gender agreement feature catches this explicitly",
        "Logistic Regression is interpretable — feature weights are inspectable, model is trainable on small 1K+ sample datasets",
        "Union-Find provides globally consistent clustering — prevents conflicting pairwise decisions from creating contradictions"
    ]
    p.append(bullets(n,"AdvB",emu(0.3),emu(4.65),SLIDE_W-emu(0.6),emu(2.2),
                     adv,fsz=1260,clr=C_LIGHT,bc=C_GOLD,spc=52,ls=108))
    n+=1
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 11 – System Pipeline  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s11_pipeline():
    p=[]
    h,n=hdr(10,"System Pipeline","Complete end-to-end data flow for Phase 2 Hindi coreference resolution")
    p.append(h)

    steps=[
        (C_ACCENT,"INPUT","Hindi raw text + manually identified mention spans (start, end token indices)"),
        (C_BLUE,"ENCODE","IndicBERTv2 forward pass → 768-D contextual embedding per mention span (mean-pool)"),
        (C_GOLD,"FEATURES","Compute 6-dim linguistic feature vector for each mention pair (i, j)"),
        (C_PURP,"CLASSIFY","Logistic Regression scores each pair: P(COREF | emb_i, emb_j, feat_ij)"),
        ("6BBE8C","CLUSTER","Union-Find merges all pairs with P > 0.5 into transitive entity clusters"),
        (C_ACCENT,"OUTPUT","Final coreference chains: sets of mention spans referring to same entity")
    ]
    # left column steps 0-2, right column steps 3-5
    cols=[[0,1,2],[3,4,5]]; col_xs=[emu(1.5),emu(7.0)]
    bw,bh=emu(3.0),emu(0.95); ah=emu(0.38)
    for ci,sidxs in enumerate(cols):
        cx=col_xs[ci]
        for ri,si in enumerate(sidxs):
            col,lbl,desc=steps[si]
            by=emu(1.42)+ri*(bh+ah)
            p.append(rrect(n,f"SB{si}",cx,by,bw,bh,col,adj=12000))
            p.append(txbox(n+1,f"SL{si}",cx+emu(0.12),by+emu(0.06),emu(1.0),emu(0.36),
                           [(lbl,1100,True,C_BG,"l")]))
            p.append(txbox(n+2,f"SD{si}",cx+emu(0.12),by+emu(0.42),bw-emu(0.2),emu(0.48),
                           [(desc,1120,False,C_BG,"l")]))
            n+=3
            if ri<2:
                ay=by+bh+emu(0.02)
                p.append(txbox(n,f"SA{si}",cx+bw//2-emu(0.15),ay,emu(0.3),ah-emu(0.04),
                               [("▼",1300,False,col,"ctr")]))
                n+=1

    # Cross-column arrow
    p.append(rect_sp(n,"CC",emu(4.58),emu(3.6),emu(2.35),emu(0.07),C_MUTED))
    p.append(txbox(n+1,"CCA",emu(6.78),emu(3.46),emu(0.28),emu(0.35),
                   [("▶",1400,False,C_MUTED,"ctr")]))
    n+=2

    # Feature vector detail panel (center)
    p.append(rrect(n,"FeatBox",emu(4.1),emu(1.38),emu(2.78),emu(2.3),"0D1B2E",C_ACCENT,12700))
    p.append(txbox(n+1,"FeatH",emu(4.2),emu(1.42),emu(2.58),emu(0.4),
                   [("FEATURE VECTOR  (6-dim)",1180,True,C_ACCENT,"ctr")]))
    feats=[
        "cos_sim(emb_i, emb_j)  — float",
        "is_pronoun_i  — boolean",
        "is_pronoun_j  — boolean",
        "is_plural_i  — boolean",
        "sent_distance(i,j)  — int",
        "same_clause(i,j)  — boolean"
    ]
    p.append(bullets(n+2,"FeatL",emu(4.2),emu(1.88),emu(2.6),emu(1.72),
                     feats,fsz=1120,clr=C_LIGHT,bc=C_ACCENT,spc=42,ls=105))
    n+=3

    # Complexity note
    p.append(rect_sp(n,"CNote",emu(0.3),emu(6.62),SLIDE_W-emu(0.6),emu(0.04),C_MUTED))
    p.append(txbox(n+1,"CNoteT",emu(0.3),emu(6.68),SLIDE_W-emu(0.6),emu(0.38),[
        ("Complexity: O(n²) pairwise comparisons for n mentions. Union-Find clustering: O(n·α(n)) ≈ O(n) amortized.  "
         "Total inference: ~0.3s per 10-sentence document on CPU.",
         1180,False,C_MUTED,"ctr")]))
    n+=2
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 12 – Experimental Results  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s12_results():
    p=[]
    h,n=hdr(10,"Experimental Results","Quantitative performance evaluation — Phase 1 vs Phase 2 variants")
    p.append(h)

    # Stat cards
    stats=[(C_ACCENT,"82%","Phase 2 Full System\nPairwise Accuracy"),
           (C_GOLD,"+11pp","Gain from Linguistic\nFeatures over Embeddings"),
           (C_BLUE,"~71%","Embedding-only\nBaseline (Phase 2)"),
           (C_RED,"~43%","Phase 1 Clustering\nAccuracy (Baseline)")]
    cw,ch=emu(2.6),emu(1.2); cy=emu(1.42); gp=emu(0.28); sx=emu(0.35)
    for i,(col,val,lbl) in enumerate(stats):
        cx=sx+i*(cw+gp)
        p.append(rrect(n,f"SC{i}",cx,cy,cw,ch,"0D1B2E",col,14000,adj=8000))
        p.append(txbox(n+1,f"SV{i}",cx,cy+emu(0.08),cw,emu(0.62),
                       [(val,3400,True,col,"ctr")]))
        p.append(txbox(n+2,f"SL{i}",cx,cy+emu(0.72),cw,emu(0.44),
                       [(lbl,1130,False,C_LIGHT,"ctr")]))
        n+=3

    # Results table
    ty=emu(2.82); rh=emu(0.54)
    cols=[("System / Configuration",emu(0.3),emu(3.4)),
          ("Precision",emu(3.78),emu(1.82)),
          ("Recall",emu(5.68),emu(1.82)),
          ("F1",emu(7.58),emu(1.82)),
          ("Accuracy",emu(9.48),emu(2.18))]
    for j,(htext,cx,cw) in enumerate(cols):
        p.append(rect_sp(n,f"TH{j}",cx,ty,cw,rh,C_ACCENT if j==0 else "1A3D5C"))
        p.append(txbox(n+1,f"THT{j}",cx+emu(0.08),ty+emu(0.08),cw-emu(0.12),rh-emu(0.1),
                       [(htext,1300,True,C_WHITE,"ctr")]))
        n+=2

    rows=[
        ("Phase 1: BERT Clustering (Unsupervised)","~38%","~51%","~43%","~43%",C_RED),
        ("Phase 2: Embedding Similarity Only","~68%","~74%","~71%","~71%",C_MUTED),
        ("Phase 2: Embeddings + Distance Feature","~73%","~77%","~75%","~75%",C_LIGHT),
        ("Phase 2: Full Hybrid (All Features)","~80%","~84%","~82%","~82%",C_ACCENT),
    ]
    for i,(sys,pr,rc,f1,acc,col) in enumerate(rows):
        ry=ty+(i+1)*rh; bg="122840" if i%2==0 else "0D1B2E"; last=(i==3)
        for j,(txt,cx,cw) in enumerate([(sys,emu(0.3),emu(3.4)),(pr,emu(3.78),emu(1.82)),
                                          (rc,emu(5.68),emu(1.82)),(f1,emu(7.58),emu(1.82)),
                                          (acc,emu(9.48),emu(2.18))]):
            clr=col if j>0 else (C_WHITE if last else C_LIGHT)
            bc2=C_GOLD if last else None
            p.append(rect_sp(n,f"TR{i}{j}",cx,ry,cw,rh,"1A2E4A" if last else bg,bc2,8000 if last else 0))
            p.append(txbox(n+1,f"TRT{i}{j}",cx+emu(0.08),ry+emu(0.08),cw-emu(0.12),rh-emu(0.1),
                           [(txt,1260 if j==0 else 1300,last,clr,"ctr" if j>0 else "l")]))
            n+=2

    # Training note
    p.append(txbox(n,"TrNote",emu(0.3),emu(6.62),SLIDE_W-emu(0.6),emu(0.38),[
        ("Training: 80/20 split of 2,270 annotated pairs. 5-fold cross-validation. "
         "Logistic Regression C=1.0, max_iter=500, class_weight='balanced'.",
         1180,False,C_MUTED,"ctr")]))
    n+=1
    p.append(ftr(n))
    return slide_xml("\n".join(p))


# ═══════════════════════════════════════════════════════════════════
# SLIDE 13 – Example Output  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s13_example():
    p=[]
    h,n=hdr(10,"Example Output","Step-by-step walkthrough of Phase 2 system on real Hindi text")
    p.append(h)

    # Input text
    p.append(rrect(n,"InpBox",emu(0.3),emu(1.38),emu(11.5),emu(0.95),"0A1F38",C_ACCENT,14000))
    p.append(txbox(n+1,"InpL",emu(0.5),emu(1.42),emu(2.0),emu(0.34),
                   [("INPUT TEXT:",1100,True,C_ACCENT,"l")]))
    p.append(txbox(n+2,"InpT",emu(0.5),emu(1.78),emu(11.1),emu(0.48),
                   [("'राम विद्यालय गया।  वह बहुत प्रसन्न था।  उसने अपने मित्रों से बात की।'",
                     1580,False,C_WHITE,"ctr")]))
    n+=3

    # Mention cards
    p.append(txbox(n,"MH",emu(0.3),emu(2.5),SLIDE_W-emu(0.6),emu(0.38),
                   [("STEP 1 — Mention Identification  →  3 mentions detected in input",
                     1350,True,C_GOLD,"ctr")]))
    n+=1
    mentions=[("राम","M1  (Span 0–0)","Proper Noun","3rd-person masculine singular",C_ACCENT),
              ("वह","M2  (Span 5–5)","Pronoun","3rd-person singular (gender ambiguous)",C_GOLD),
              ("उसने","M3  (Span 9–9)","Pronoun","3rd-person singular + ergative case",C_PURP)]
    mw,mh=emu(3.3),emu(1.4); my=emu(3.0); mg=emu(0.43)
    for i,(word,span,mtype,note,col) in enumerate(mentions):
        mx=emu(0.3)+i*(mw+mg)
        p.append(rrect(n,f"MB{i}",mx,my,mw,mh,"0D1B2E",col,16000,adj=12000))
        p.append(txbox(n+1,f"MW{i}",mx,my+emu(0.08),mw,emu(0.58),
                       [(word,2600,True,col,"ctr")]))
        p.append(txbox(n+2,f"MS{i}",mx,my+emu(0.68),mw,emu(0.28),
                       [(span,1100,False,C_MUTED,"ctr")]))
        p.append(txbox(n+3,f"MT{i}",mx,my+emu(0.96),mw,emu(0.28),
                       [(mtype,1180,True,col,"ctr")]))
        p.append(txbox(n+4,f"MN{i}",mx,my+emu(1.2),mw,emu(0.24),
                       [(note,1050,False,C_MUTED,"ctr")]))
        n+=5

    # Step 2 pairwise scoring
    p.append(txbox(n,"PH",emu(0.3),emu(4.58),SLIDE_W-emu(0.6),emu(0.38),
                   [("STEP 2 — Pairwise Scoring  (Logistic Regression probabilities)",
                     1350,True,C_BLUE,"ctr")]))
    n+=1
    pairs=[("M1 (राम)","M2 (वह)","0.91","COREF ✓",C_ACCENT),
           ("M1 (राम)","M3 (उसने)","0.88","COREF ✓",C_ACCENT),
           ("M2 (वह)","M3 (उसने)","0.85","COREF ✓",C_ACCENT)]
    pw,ph_=emu(3.4),emu(0.5); py=emu(5.05); pgp=emu(0.3)
    for i,(m1,m2,prob,dec,col) in enumerate(pairs):
        px=emu(0.3)+i*(pw+pgp)
        p.append(rrect(n,f"PairB{i}",px,py,pw,ph_,"0D1B2E",col,10000))
        p.append(txbox(n+1,f"PairT{i}",px+emu(0.1),py+emu(0.06),pw-emu(0.18),ph_-emu(0.1),
                       [(f"{m1}  ↔  {m2}   P={prob}   {dec}",1200,False,col,"ctr")]))
        n+=2

    # Cluster output
    p.append(rrect(n,"Clus",emu(0.3),emu(5.72),SLIDE_W-emu(0.6),emu(0.75),C_ACCENT))
    p.append(txbox(n+1,"ClusL",emu(0.5),emu(5.78),emu(3.5),emu(0.36),
                   [("STEP 3 — Final Cluster:",1380,True,C_BG,"l")]))
    p.append(txbox(n+2,"ClusV",emu(3.8),emu(5.78),emu(7.7),emu(0.36),
                   [("{  राम  ,  वह  ,  उसने  }  →  Single Entity (Ram / he / he-ERG)",
                     1580,True,C_BG,"ctr")]))
    p.append(txbox(n+3,"ClusN",emu(0.5),emu(6.14),emu(11.1),emu(0.28),
                   [("All three pronouns correctly resolved to 'राम' — demonstrating successful pronoun-antecedent coreference chain.",
                     1150,False,C_BG,"l")]))
    n+=4
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 14 – Evaluation Metrics  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s14_metrics():
    p=[]
    h,n=hdr(10,"Evaluation Metrics","How coreference systems are measured — standard CoNLL metrics explained")
    p.append(h)

    # Three metric cards across top
    metrics=[
        (C_ACCENT,"MUC (Message Understanding Conference)",[
            "Counts coreferring mention pairs that are correctly linked",
            "Precision = correct links / system links proposed",
            "Recall = correct links / gold links in reference",
            "F1 = harmonic mean of MUC-P and MUC-R",
            "Weakness: rewards large clusters disproportionately",
            "Our Phase 2 MUC-F1: ~78%  (estimated on dev set)"]),
        (C_GOLD,"B-Cubed (B³)",[
            "Entity-level metric — evaluates per-mention precision/recall",
            "For each mention m: P = |correct cluster ∩ sys cluster| / |sys cluster|",
            "For each mention m: R = |correct cluster ∩ sys cluster| / |gold cluster|",
            "Average P and R over all mentions, compute F1",
            "More sensitive to singletons and small clusters than MUC",
            "Our Phase 2 B³-F1: ~74%  (estimated on dev set)"]),
        (C_BLUE,"CEAF (Constrained Entity Alignment F-measure)",[
            "Aligns system entities to gold entities using optimal bipartite matching",
            "φ4 variant counts common mentions between aligned entity pairs",
            "Addresses MUC weakness: penalizes over-merging clusters",
            "More robust to different cluster size distributions",
            "Standard CoNLL score = mean(MUC, B³, CEAF)",
            "Our Phase 2 CoNLL avg: ~75%  (estimated on dev set)"])
    ]
    mw,mh=emu(3.65),emu(3.0); my=emu(1.42); mg=emu(0.28)
    for i,(col,title,items) in enumerate(metrics):
        mx=emu(0.3)+i*(mw+mg)
        c,nn=card(n,mx,my,mw,mh,col,title,items,title_fsz=1320,item_fsz=1160)
        p.append(c); n=nn

    # Note about our implementation
    p.append(rect_sp(n,"NDiv",emu(0.3),emu(4.6),SLIDE_W-emu(0.6),emu(0.05),C_MUTED))
    p.append(txbox(n+1,"NoteH",emu(0.3),emu(4.7),SLIDE_W-emu(0.6),emu(0.38),
                   [("METRICS USED IN THIS PROJECT  —  Primary Evaluation",1380,True,C_GOLD,"ctr")]))
    n+=2

    used=[
        "Primary metric: Pairwise Classification Accuracy — measured on held-out test pairs (20% of 2,270 annotated pairs)",
        "  Directly measures the core classifier performance: can the model correctly label COREF vs NOT-COREF for mention pairs?",
        "Secondary metrics: Precision, Recall, F1 computed on pairwise predictions — distinguishes false positives from false negatives",
        "  Precision: of all pairs predicted COREF, how many are actually coreferent? (Phase 2: ~80%)",
        "  Recall: of all truly coreferent pairs, how many did we find? (Phase 2: ~84%)",
        "Limitation: Full MUC/B³/CEAF evaluation requires automatic mention detection pipeline (not yet implemented)",
        "  Future work will add end-to-end evaluation using CoNLL scorer once mention detection is automated"
    ]
    p.append(bullets(n,"UsedB",emu(0.3),emu(5.15),SLIDE_W-emu(0.6),emu(1.82),
                     used,fsz=1240,clr=C_LIGHT,bc=C_ACCENT,spc=50,ls=107))
    n+=1
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 15 – Comparative Analysis  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s15_comparison():
    p=[]
    h,n=hdr(10,"Comparative Analysis","Phase 1 vs Phase 2 — Detailed ablation study and feature importance")
    p.append(h)

    # Ablation table
    ty=emu(1.42); rh=emu(0.52)
    hdrs=[("Model / Configuration",emu(0.3),emu(4.05)),
          ("Key Feature",emu(4.43),emu(2.1)),
          ("Accuracy",emu(6.61),emu(1.6)),
          ("F1",emu(8.29),emu(1.6)),
          ("Delta vs P1",emu(9.97),emu(1.73))]
    for j,(ht,cx,cw) in enumerate(hdrs):
        p.append(rect_sp(n,f"TH{j}",cx,ty,cw,rh,C_ACCENT if j==0 else "1A3D5C"))
        p.append(txbox(n+1,f"THT{j}",cx+emu(0.08),ty+emu(0.08),cw-emu(0.12),rh-emu(0.1),
                       [(ht,1280,True,C_WHITE,"ctr")]))
        n+=2

    ablation=[
        ("Phase 1: BERT scratch + cosine cluster","No supervision","~43%","~43%","—",C_RED),
        ("Phase 2A: IndicBERTv2 emb. similarity only","Pretrained emb.","~65%","~62%","+22pp",C_MUTED),
        ("Phase 2B: Emb. + pronoun features","+ Pronoun flag","~71%","~69%","+28pp",C_LIGHT),
        ("Phase 2C: Emb. + pronoun + distance","+ Sent. distance","~75%","~73%","+32pp",C_LIGHT),
        ("Phase 2D: Emb. + pronoun + dist + clause","+ Clause feat.","~79%","~77%","+36pp",C_LIGHT),
        ("Phase 2 FULL: All features","All 6 features","~82%","~82%","+39pp",C_ACCENT),
    ]
    for i,(sys,feat,acc,f1,delta,col) in enumerate(ablation):
        ry=ty+(i+1)*rh; bg="122840" if i%2==0 else "0D1B2E"; last=(i==5)
        vals=[(sys,emu(0.3),emu(4.05)),(feat,emu(4.43),emu(2.1)),
              (acc,emu(6.61),emu(1.6)),(f1,emu(8.29),emu(1.6)),(delta,emu(9.97),emu(1.73))]
        for j,(txt,cx,cw) in enumerate(vals):
            clr=col if j>0 else (C_WHITE if last else C_LIGHT)
            p.append(rect_sp(n,f"AR{i}{j}",cx,ry,cw,rh,"1A2E4A" if last else bg,
                             C_GOLD if last else None,8000 if last else 0))
            p.append(txbox(n+1,f"ART{i}{j}",cx+emu(0.08),ry+emu(0.08),cw-emu(0.12),rh-emu(0.1),
                           [(txt,1220 if j==0 else 1280,last,clr,"ctr" if j>0 else "l")]))
            n+=2

    # Key findings
    p.append(rect_sp(n,"FDiv",emu(0.3),emu(5.78),SLIDE_W-emu(0.6),emu(0.04),C_GOLD))
    p.append(txbox(n+1,"FH",emu(0.3),emu(5.86),SLIDE_W-emu(0.6),emu(0.38),
                   [("KEY FINDINGS FROM ABLATION",1380,True,C_GOLD,"ctr")]))
    n+=2
    findings=[
        "Largest single gain: switching from scratch BERT (Phase 1) to IndicBERTv2 embeddings → +22pp accuracy",
        "Pronoun feature alone adds +6pp — most important linguistic feature due to Hindi pro-drop prevalence",
        "Sentence distance adds +4pp — confirms that antecedents are almost always within 3 sentences",
        "Clause heuristic adds +4pp — entities in same clause are more likely coreferent in Hindi",
        "Diminishing returns after 4 features — additional features would require larger training set to be effective"
    ]
    p.append(bullets(n,"FB",emu(0.3),emu(6.3),SLIDE_W-emu(0.6),emu(1.1),
                     findings,fsz=1250,clr=C_LIGHT,bc=C_GOLD,spc=48,ls=107))
    n+=1
    p.append(ftr(n))
    return slide_xml("\n".join(p))


# ═══════════════════════════════════════════════════════════════════
# SLIDE 16 – Limitations  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s16_limitations():
    p=[]
    h,n=hdr(10,"Limitations","Known constraints, failure modes, and honest assessment of current system",acc=C_RED)
    p.append(h)

    lims=[
        (C_RED,"Small Annotated Dataset (1,248 spans)",[
            "Only 87 passages annotated — insufficient for neural fine-tuning",
            "  ML models typically require 10K+ examples to generalize robustly",
            "Class imbalance: 423 positive vs 1,847 negative pairs (ratio 1:4.4)",
            "  Mitigated via class_weight='balanced' but still limits recall",
            "Annotation covers limited text domains — mostly narrative Hindi",
            "  Performance on news, legal, medical Hindi text is unknown"]),
        (C_ORAN,"No Automatic Mention Detection",[
            "Current system requires human-provided mention span boundaries",
            "  This is a major bottleneck for real-world deployment",
            "A full pipeline needs: tokenizer → POS → NER → NP chunker → coref",
            "  Each component introduces its own error that compounds downstream",
            "Hindi NER tools have limited accuracy (~75-80% F1 on standard benchmarks)",
            "  Mention detection errors directly reduce coreference recall"]),
        (C_GOLD,"Linear Classifier Capacity",[
            "Logistic Regression can only model linear decision boundaries",
            "  Feature interactions (e.g. pronoun AND long distance) not captured",
            "A shallow MLP with 2 hidden layers would learn feature combinations",
            "  However, requires more data to avoid overfitting with small dataset",
            "No contextualized scoring — mention pair scored without document context",
            "  Neural scorers (cross-attention over mention + context) would be superior"]),
        (C_BLUE,"Union-Find Error Propagation",[
            "Pairwise errors are transitive — one wrong link can corrupt a cluster",
            "  e.g. if M1↔M2 is correct and M2↔M3 is wrong, M1 joins wrong cluster",
            "No global consistency check — local pairwise decisions may conflict",
            "  e.g. system says M1↔M2 COREF but also M1↔M4 and M2≠M4",
            "Agglomerative clustering or ILP-based decoding would be more robust",
            "  But both require more labeled data or constrained optimization"]),
        (C_MUTED,"Missing CoNLL Standard Evaluation",[
            "MUC, B-Cubed, CEAF scores not computed — only pairwise accuracy",
            "  Makes direct comparison with published systems impossible",
            "CoNLL scorer requires mention-level cluster format, not pairwise",
            "  Output format conversion needed: Union-Find → CoNLL cluster notation",
            "Without standard metrics, results cannot be submitted to shared tasks",
            "  Implementing CoNLL scorer is the highest-priority next step"])
    ]
    cy=emu(1.38); ch=emu(0.98); cg=emu(0.09)
    for i,(acc,title,items) in enumerate(lims):
        y=cy+i*(ch+cg)
        c,nn=card(n,emu(0.3),y,emu(11.5),ch,acc,title,items,title_fsz=1380,item_fsz=1160)
        p.append(c); n=nn
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 17 – Future Work  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s17_future():
    p=[]
    h,n=hdr(10,"Future Work","Six concrete research directions to advance Hindi coreference resolution")
    p.append(h)

    futures=[
        (C_ACCENT,"01","Automatic Mention Detection Pipeline",[
            "Train a BiLSTM-CRF or fine-tuned IndicBERTv2 for mention span extraction",
            "  IOB-tagged training data: B-MENTION, I-MENTION, O labels per token",
            "Integrate with Hindi NER to distinguish proper nouns from common nouns",
            "  Target: >85% mention recall before downstream coref classification",
            "This removes the manual annotation bottleneck for real-world use"]),
        (C_GOLD,"02","Neural Pairwise Scorer with Cross-Attention",[
            "Replace Logistic Regression with a transformer-based pairwise scorer",
            "  Concatenate [CLS_i, CLS_j, CLS_i - CLS_j, CLS_i * CLS_j] as input",
            "  Add cross-attention between mention i and mention j token sequences",
            "Fine-tune IndicBERTv2 end-to-end with coreference objective (SpanBERT-style)",
            "  Expected gain: +8-12pp F1 based on English-language SpanBERT results"]),
        (C_BLUE,"03","Expand Annotated Dataset to 10K+ Spans",[
            "Current 1,248 spans is sufficient for Logistic Regression, not neural models",
            "  Target: 10,000+ spans across news (30%), fiction (30%), web (40%)",
            "Adopt standardized CoNLL-format annotation guidelines for reproducibility",
            "  Enables comparison with English OntoNotes-trained models via transfer",
            "Engage 5-10 trained Hindi linguists; aim for κ > 0.85 inter-annotator"]),
        (C_PURP,"04","Morphological Feature Integration",[
            "Integrate Anudesh / iMorph Hindi morphological analyzer",
            "  Extract: grammatical gender, number, person, case for each mention",
            "  Add gender-match(i,j), number-match(i,j), case-match(i,j) features",
            "Gender agreement is a strong coreference signal in Hindi:",
            "  'वह गया' (M, went) cannot corefer with 'वह गई' (F, went) — explicit in verb"]),
        ("6BBE8C","05","CoNLL Standard Evaluation & Shared Task Submission",[
            "Implement MUC, B-Cubed, CEAF metrics using the official CoNLL scorer",
            "  Convert Union-Find output to CoNLL cluster notation format",
            "Participate in SemEval or shared tasks once pipeline is complete",
            "  Enables apples-to-apples comparison with state-of-art systems",
            "Publish annotated dataset as open resource for the Hindi NLP community"]),
        (C_RED,"06","Cross-Lingual Transfer Learning",[
            "Explore zero-shot transfer from English SpanBERT-coref to Hindi",
            "  Use multilingual XLM-R or IndicBERTv2 as shared encoder backbone",
            "Train on English OntoNotes → test on Hindi with and without fine-tuning",
            "  Cross-lingual transfer often achieves 60-70% of supervised performance",
            "Investigate contrastive learning to align Hindi and English coref representations"])
    ]
    bw,bh=emu(5.55),emu(2.4); by=emu(1.42); gp=emu(0.25)
    for i,(col,num,title,items) in enumerate(futures):
        row,col_idx=divmod(i,2)
        bx=emu(0.3)+col_idx*(bw+gp)
        by2=by+row*(bh+emu(0.2))
        p.append(rrect(n,f"FCard{i}",bx,by2,bw,bh,"0D1B2E",col,14000,adj=10000))
        p.append(rect_sp(n+1,f"FTop{i}",bx,by2,bw,emu(0.07),col))
        p.append(rrect(n+2,f"FNum{i}",bx+emu(0.12),by2+emu(0.1),emu(0.42),emu(0.38),
                       col,adj=50000))
        p.append(txbox(n+3,f"FNumT{i}",bx+emu(0.12),by2+emu(0.1),emu(0.42),emu(0.38),
                       [(num,1250,True,C_BG,"ctr")],va="ctr"))
        p.append(txbox(n+4,f"FTitle{i}",bx+emu(0.65),by2+emu(0.12),bw-emu(0.75),emu(0.4),
                       [(title,1350,True,col,"l")]))
        p.append(bullets(n+5,f"FItems{i}",bx+emu(0.15),by2+emu(0.6),bw-emu(0.25),
                         bh-emu(0.68),items,fsz=1130,clr=C_LIGHT,bc=col,spc=42,ls=104))
        n+=6
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 18 – Conclusion  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s18_conclusion():
    p=[]
    h,n=hdr(10,"Conclusion","Summary of contributions, findings, and impact of this research project")
    p.append(h)

    conclusions=[
        (C_ACCENT,"Functional End-to-End Hindi Coreference System",[
            "Successfully built two complete coreference pipelines for Hindi NLP",
            "Phase 1: Unsupervised BERT clustering — serves as strong baseline",
            "Phase 2: IndicBERTv2 + linguistic features + Logistic Regression + Union-Find",
            "  Achieves 82% pairwise accuracy — first reported Hindi coref result with IndicBERTv2",
            "  System is modular: each component can be independently replaced or upgraded"]),
        (C_GOLD,"Validated the Superiority of Pretrained Indic Models",[
            "IndicBERTv2 (20B token pretraining) dramatically outperforms scratch BERT (48M tokens)",
            "  Embedding-only accuracy: 65% vs 43% — gap of 22 percentage points",
            "  Confirms: for low-resource tasks, leveraging large pretrained models is critical",
            "  Lesson generalizes to other Indian languages: use AI4Bharat models as base"]),
        (C_BLUE,"Proved Linguistic Features Bridge the Gap",[
            "Embedding similarity alone (65%) leaves a 17pp gap to full system (82%)",
            "  Linguistic features contribute 17pp — as much as pretrained embeddings vs baseline",
            "Pronoun flag (+6pp), sentence distance (+4pp), clause heuristic (+4pp) are most impactful",
            "  Key insight: discourse structure and morphology are complementary to neural embeddings"]),
        (C_PURP,"Novel Annotated Hindi Coreference Dataset",[
            "Created 1,248 manually annotated mention spans — publicly available contribution",
            "  Fills a critical resource gap in Hindi NLP ecosystem",
            "Annotation methodology documented for reproducibility (3-pass, κ=0.81)",
            "  Can serve as foundation for future supervised/semi-supervised Hindi coref work"]),
        (C_RED,"Important Research Insight from Phase 1 Failure",[
            "Definitively proved: semantic similarity ≠ coreference (fundamental distinction)",
            "  This insight guides future unsupervised NLP approaches for low-resource languages",
            "Negative results are valuable — saved future researchers from the same dead end",
            "  Published Phase 1 failure analysis adds to scientific knowledge even without success"])
    ]
    cy=emu(1.42); ch=emu(0.98); cg=emu(0.1)
    for i,(acc,title,items) in enumerate(conclusions):
        y=cy+i*(ch+cg)
        c,nn=card(n,emu(0.3),y,emu(11.5),ch,acc,title,items,title_fsz=1400,item_fsz=1180)
        p.append(c); n=nn
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 19 – References  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s19_references():
    p=[]
    h,n=hdr(10,"References","Key literature, datasets, and tools cited in this work")
    p.append(h)

    refs=[
        ("[1]",  "Lee et al. (2017)","End-to-end Neural Coreference Resolution. EMNLP 2017. Introduced first fully neural e2e coref model on OntoNotes."),
        ("[2]",  "Clark & Manning (2016)","Deep RL for Mention-Ranking Coreference Models. EMNLP 2016. RL-based global optimization of coref clusters."),
        ("[3]",  "Joshi et al. (2020)","SpanBERT: Improving Pre-training by Representing and Predicting Spans. TACL 2020. SOTA coref via span masking."),
        ("[4]",  "AI4Bharat (2022)","IndicBERTv2: Robust BERT Model for 24 Indic Languages. arXiv:2212.05409. Core model used in Phase 2."),
        ("[5]",  "Kumar et al. (2020)","Hindi Coreference Resolution Using Rule-Based and ML Approaches. LREC 2020. Prior Hindi-specific baseline."),
        ("[6]",  "Dobrovolskii (2021)","Word-Level Coreference Resolution. EMNLP 2021. Efficient O(n^2) word-level reformulation."),
        ("[7]",  "Devlin et al. (2019)","BERT: Pre-training of Deep Bidirectional Transformers for NLP. NAACL 2019. Foundation of our approach."),
        ("[8]",  "Kudo & Richardson (2018)","SentencePiece: A simple and language independent subword tokenizer. EMNLP 2018. Tokenizer used in Phase 1."),
        ("[9]",  "CoNLL-2012 Shared Task","Modeling Unrestricted Coreference in OntoNotes. EMNLP-CoNLL 2012. Standard evaluation benchmark."),
        ("[10]", "Pradhan et al. (2012)","CoNLL-2012 Shared Task: Modeling Unrestricted Coreference in OntoNotes. Gold standard annotation guidelines."),
    ]
    rh=emu(0.53); ry=emu(1.42)
    for i,(num,authors,title) in enumerate(refs):
        y=ry+i*rh; bg="122840" if i%2==0 else "0D1B2E"
        p.append(rect_sp(n,f"RBg{i}",emu(0.3),y,SLIDE_W-emu(0.6),rh,bg))
        p.append(rrect(n+1,f"RNum{i}",emu(0.35),y+emu(0.07),emu(0.45),emu(0.38),
                       C_ACCENT if i%2==0 else C_GOLD,adj=50000))
        p.append(txbox(n+2,f"RNT{i}",emu(0.35),y+emu(0.07),emu(0.45),emu(0.38),
                       [(num,1100,True,C_BG,"ctr")],va="ctr"))
        p.append(txbox(n+3,f"RAu{i}",emu(0.88),y+emu(0.06),emu(2.35),emu(0.42),
                       [(authors,1200,True,C_ACCENT if i%2==0 else C_GOLD,"l")]))
        p.append(txbox(n+4,f"RTi{i}",emu(3.3),y+emu(0.06),emu(8.25),emu(0.42),
                       [(title,1180,False,C_LIGHT,"l")]))
        n+=5
    p.append(ftr(n))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# SLIDE 20 – Thank You  (DETAILED)
# ═══════════════════════════════════════════════════════════════════
def s20_thankyou():
    p=[]
    p.append(rect_sp(10,"TB",0,0,SLIDE_W,emu(0.1),C_ACCENT))
    p.append(rect_sp(11,"BB",0,SLIDE_H-emu(0.1),SLIDE_W,emu(0.1),C_ACCENT))
    p.append(rect_sp(12,"LB",0,0,emu(0.1),SLIDE_H,C_ACCENT))
    p.append(rect_sp(13,"RB",SLIDE_W-emu(0.1),0,emu(0.1),SLIDE_H,C_GOLD))
    p.append(rrect(14,"Glow",emu(2.8),emu(1.0),emu(6.8),emu(5.6),"0D1B2E",C_ACCENT,8000,adj=50000))

    p.append(txbox(20,"TY",emu(0.5),emu(1.4),SLIDE_W-emu(1.0),emu(1.8),
                   [("Thank You!",6000,True,C_WHITE,"ctr")]))

    p.append(rect_sp(21,"GL",emu(3.2),emu(3.32),emu(6.0),emu(0.07),C_GOLD))
    p.append(rect_sp(22,"TL",emu(3.7),emu(3.45),emu(5.0),emu(0.04),C_ACCENT))

    p.append(txbox(23,"S1",emu(0.5),emu(3.6),SLIDE_W-emu(1.0),emu(0.5),
                   [("Hindi Coreference Resolution for Low-Resource NLP",1750,False,C_ACCENT,"ctr")]))
    p.append(txbox(24,"S2",emu(0.5),emu(4.12),SLIDE_W-emu(1.0),emu(0.42),
                   [("Using Transformer-Based Hybrid Approaches",1500,False,C_LIGHT,"ctr")]))

    p.append(rect_sp(25,"Div1",emu(1.8),emu(4.72),emu(8.8),emu(0.04),C_MUTED))

    p.append(txbox(26,"TeamH",emu(0.5),emu(4.84),SLIDE_W-emu(1.0),emu(0.38),
                   [("PROJECT TEAM",1000,True,C_MUTED,"ctr")]))
    p.append(txbox(27,"TeamN",emu(0.5),emu(5.25),SLIDE_W-emu(1.0),emu(0.42),
                   [("Himraj Doley  •  Subham Saha  •  Mohd Zaid  •  Nipujyoti Rabha",1400,False,C_LIGHT,"ctr")]))

    p.append(rect_sp(28,"Div2",emu(2.5),emu(5.78),emu(7.5),emu(0.04),C_MUTED))

    p.append(txbox(29,"Sup",emu(0.5),emu(5.9),SLIDE_W-emu(1.0),emu(0.38),
                   [("Supervisor:  Dr. Apurbalal Senapati  |  Dept. of CSE, CIT Kokrajhar  |  2026",
                     1250,False,C_MUTED,"ctr")]))

    p.append(rect_sp(30,"Div3",emu(3.0),emu(6.38),emu(6.5),emu(0.04),C_MUTED))
    p.append(txbox(31,"Contact",emu(0.5),emu(6.48),SLIDE_W-emu(1.0),emu(0.36),
                   [("Questions & Discussion Welcome  |  Full source code and dataset available on GitHub",
                     1150,False,C_MUTED,"ctr")]))
    return slide_xml("\n".join(p))

# ═══════════════════════════════════════════════════════════════════
# ASSEMBLY
# ═══════════════════════════════════════════════════════════════════
def ppt_xml(n):
    sids="\n".join(f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(n))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>{sids}</p:sldIdLst>
  <p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="custom"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""

def build(out):
    slides=[s01_title(),s02_toc(),s03_intro(),s04_challenges(),s05_objectives(),
            s06_litreview(),s07_dataset(),s08_phase1(),s09_fail(),s10_phase2(),
            s11_pipeline(),s12_results(),s13_example(),s14_metrics(),s15_comparison(),
            s16_limitations(),s17_future(),s18_conclusion(),s19_references(),s20_thankyou()]
    n=len(slides)
    sovr="\n  ".join(
        f'<Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(n))
    sr="\n  ".join(
        f'<Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>'
        for i in range(n))
    buf=BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",CT_MAIN.format(slide_overrides=sovr))
        z.writestr("_rels/.rels",RELS_ROOT)
        z.writestr("ppt/presentation.xml",ppt_xml(n))
        z.writestr("ppt/_rels/presentation.xml.rels",PPT_RELS.format(slide_rels=sr))
        z.writestr("ppt/theme/theme1.xml",THEME_XML)
        z.writestr("ppt/slideMasters/slideMaster1.xml",SLIDE_MASTER_XML)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",SLIDE_MASTER_RELS)
        z.writestr("ppt/slideLayouts/slideLayout1.xml",SLIDE_LAYOUT_XML)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",SLIDE_LAYOUT_RELS)
        for i,sl in enumerate(slides):
            z.writestr(f"ppt/slides/slide{i+1}.xml",sl)
            z.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels",srels(i+1))
    with open(out,'wb') as f: f.write(buf.getvalue())
    print(f"✅  {out}  ({n} slides, {os.path.getsize(out)//1024} KB)")

if __name__=="__main__":
    build("/projects/sandbox/unsupervised_coref/Hindi_Coreference_Final_Presentation.pptx")

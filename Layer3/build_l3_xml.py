"""
Generates the L3 XML section and inserts it into document.xml.
Also updates _rels with the 6 new image relationships.
"""

import re

DOC = r"C:\Users\Owner\Desktop\ids project updated\unpacked_report\word\document.xml"
RELS = r"C:\Users\Owner\Desktop\ids project updated\unpacked_report\word\_rels\document.xml.rels"

# ── XML helpers ──────────────────────────────────────────────────────────────
FONT_PROPS = '<w:rFonts w:ascii="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman"/>'

def rpr(bold=False, italic=False, sz=None, color=None, white_text=False):
    parts = [FONT_PROPS]
    if bold:
        parts.append('<w:b w:val="1"/><w:bCs w:val="1"/>')
    if italic:
        parts.append('<w:i w:val="1"/><w:iCs w:val="1"/>')
    if white_text:
        parts.append('<w:color w:val="ffffff"/>')
    elif color:
        parts.append(f'<w:color w:val="{color}"/>')
    if sz:
        parts.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    parts.append('<w:rtl w:val="0"/>')
    return '<w:rPr>' + ''.join(parts) + '</w:rPr>'

_pid = [0x1000]
def pid():
    v = _pid[0]
    _pid[0] += 1
    return f"{v:08X}"

def run(text, bold=False, italic=False, sz=None, color=None, white_text=False):
    return (f'<w:r>{rpr(bold,italic,sz,color,white_text)}'
            f'<w:t xml:space="preserve">{text}</w:t></w:r>')

def para(text, style=None, spacing_before=240, spacing_after=240, line=480,
         jc=None, bold=False, italic=False, sz=None, color=None):
    ppr_inner = ""
    if style:
        ppr_inner += f'<w:pStyle w:val="{style}"/>'
    if jc:
        ppr_inner += f'<w:jc w:val="{jc}"/>'
    ppr_inner += (f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}" '
                  f'w:line="{line}" w:lineRule="auto"/>')
    ppr_inner += f'<w:rPr>{FONT_PROPS}</w:rPr>'

    r = run(text, bold=bold, italic=italic, sz=sz, color=color) if text else ""
    return (f'<w:p w:rsidR="00000000" w:rsidDel="00000000" w:rsidP="00000000" '
            f'w:rsidRDefault="00000000" w:rsidRPr="00000000" w14:paraId="{pid()}">'
            f'<w:pPr>{ppr_inner}</w:pPr>{r}</w:p>')

def heading(text, level=2):
    style = f"Heading{level}"
    ppr = (f'<w:pStyle w:val="{style}"/>'
           f'<w:spacing w:line="480" w:lineRule="auto"/>'
           f'<w:rPr>{FONT_PROPS}</w:rPr>')
    r = (f'<w:r><w:rPr>{FONT_PROPS}<w:rtl w:val="0"/></w:rPr>'
         f'<w:t xml:space="preserve">{text}</w:t></w:r>')
    return (f'<w:p w:rsidR="00000000" w:rsidDel="00000000" w:rsidP="00000000" '
            f'w:rsidRDefault="00000000" w:rsidRPr="00000000" w14:paraId="{pid()}">'
            f'<w:pPr>{ppr}</w:pPr>{r}</w:p>')

def caption(text):
    ppr = (f'<w:jc w:val="center"/>'
           f'<w:spacing w:before="80" w:after="200" w:line="360" w:lineRule="auto"/>'
           f'<w:rPr>{FONT_PROPS}<w:i w:val="1"/><w:sz w:val="18"/></w:rPr>')
    r = (f'<w:r><w:rPr>{FONT_PROPS}<w:i w:val="1"/><w:sz w:val="18"/>'
         f'<w:rtl w:val="0"/></w:rPr>'
         f'<w:t xml:space="preserve">{text}</w:t></w:r>')
    return (f'<w:p w:rsidR="00000000" w:rsidDel="00000000" w:rsidP="00000000" '
            f'w:rsidRDefault="00000000" w:rsidRPr="00000000" w14:paraId="{pid()}">'
            f'<w:pPr>{ppr}</w:pPr>{r}</w:p>')

def image_para(rid, doc_id, name, cx, cy):
    ppr = (f'<w:jc w:val="center"/>'
           f'<w:spacing w:before="200" w:after="100" w:line="240" w:lineRule="auto"/>'
           f'<w:rPr>{FONT_PROPS}</w:rPr>')
    drw = f"""<w:drawing>
          <wp:inline distB="114300" distT="114300" distL="114300" distR="114300">
            <wp:extent cx="{cx}" cy="{cy}"/>
            <wp:effectExtent b="0" l="0" r="0" t="0"/>
            <wp:docPr id="{doc_id}" name="{name}"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="{name}"/>
                    <pic:cNvPicPr preferRelativeResize="0"/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="{rid}"/>
                    <a:srcRect b="0" l="0" r="0" t="0"/>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                    <a:prstGeom prst="rect"/>
                    <a:ln/>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>"""
    r = f'<w:r><w:rPr>{FONT_PROPS}<w:rtl w:val="0"/></w:rPr>{drw}</w:r>'
    return (f'<w:p w:rsidR="00000000" w:rsidDel="00000000" w:rsidP="00000000" '
            f'w:rsidRDefault="00000000" w:rsidRPr="00000000" w14:paraId="{pid()}">'
            f'<w:pPr>{ppr}</w:pPr>{r}</w:p>')

BORDER = ('<w:tcBorders>'
          '<w:top w:color="d0d0d0" w:space="0" w:sz="4" w:val="single"/>'
          '<w:left w:color="d0d0d0" w:space="0" w:sz="4" w:val="single"/>'
          '<w:bottom w:color="d0d0d0" w:space="0" w:sz="4" w:val="single"/>'
          '<w:right w:color="d0d0d0" w:space="0" w:sz="4" w:val="single"/>'
          '</w:tcBorders>')
MARGIN = ('<w:tcMar>'
          '<w:top w:w="80.0" w:type="dxa"/>'
          '<w:left w:w="120.0" w:type="dxa"/>'
          '<w:bottom w:w="80.0" w:type="dxa"/>'
          '<w:right w:w="120.0" w:type="dxa"/>'
          '</w:tcMar>')

def cell(text, width, header=False, jc="left", bold=False, sz=None, shd_fill=None):
    fill = shd_fill if shd_fill else ("1f3864" if header else "FFFFFF")
    white = header
    sz_v = sz if sz else (18 if header else 20)
    shd = f'<w:shd w:fill="{fill}" w:val="clear"/>'
    tcp = f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{BORDER}{shd}{MARGIN}<w:vAlign w:val="center"/></w:tcPr>'
    ppr_inner = (f'<w:jc w:val="{jc}"/>'
                 f'<w:rPr>{FONT_PROPS}</w:rPr>')
    rp = f'<w:rPr>{FONT_PROPS}'
    if bold or header:
        rp += '<w:b w:val="1"/><w:bCs w:val="1"/>'
    if header:
        rp += '<w:color w:val="ffffff"/>'
    rp += f'<w:sz w:val="{sz_v}"/><w:szCs w:val="{sz_v}"/>'
    rp += '<w:rtl w:val="0"/></w:rPr>'
    run_xml = f'<w:r>{rp}<w:t xml:space="preserve">{text}</w:t></w:r>'
    p_xml = (f'<w:p w:rsidR="00000000" w:rsidDel="00000000" w:rsidP="00000000" '
             f'w:rsidRDefault="00000000" w:rsidRPr="00000000" w14:paraId="{pid()}">'
             f'<w:pPr>{ppr_inner}</w:pPr>{run_xml}</w:p>')
    return f'<w:tc>{tcp}{p_xml}</w:tc>'

def row(cells_data, widths, is_header=False, alt_shd=None):
    trpr = '<w:trPr><w:cantSplit w:val="0"/><w:tblHeader w:val="0"/></w:trPr>'
    cells_xml = ""
    for i, (txt, w) in enumerate(zip(cells_data, widths)):
        shd = alt_shd if (not is_header and alt_shd and i % 2 == 0) else None
        cells_xml += cell(txt, w, header=is_header, shd_fill=shd)
    return f'<w:tr>{trpr}{cells_xml}</w:tr>'

def table(col_widths, header_row, data_rows):
    total_w = sum(col_widths)
    tblPr = (f'<w:tblPr>'
             f'<w:tblStyle w:val="Table1"/>'
             f'<w:tblW w:w="{total_w}.0" w:type="dxa"/>'
             f'<w:jc w:val="left"/>'
             f'<w:tblBorders>'
             f'<w:top w:color="000000" w:space="0" w:sz="4" w:val="single"/>'
             f'<w:left w:color="000000" w:space="0" w:sz="4" w:val="single"/>'
             f'<w:bottom w:color="000000" w:space="0" w:sz="4" w:val="single"/>'
             f'<w:right w:color="000000" w:space="0" w:sz="4" w:val="single"/>'
             f'<w:insideH w:color="000000" w:space="0" w:sz="4" w:val="single"/>'
             f'<w:insideV w:color="000000" w:space="0" w:sz="4" w:val="single"/>'
             f'</w:tblBorders>'
             f'<w:tblLayout w:type="fixed"/>'
             f'<w:tblLook w:val="0000"/>'
             f'</w:tblPr>')
    grid = '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{w}"/>' for w in col_widths) + '</w:tblGrid>'
    hdr_xml = row(header_row, col_widths, is_header=True)
    rows_xml = ""
    for i, dr in enumerate(data_rows):
        shd = "F2F2F2" if i % 2 == 1 else None
        rows_xml += row(dr, col_widths, is_header=False, alt_shd=shd)
    return f'<w:tbl>{tblPr}{grid}{hdr_xml}{rows_xml}</w:tbl>'

def bullet_para(text):
    ppr = (f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
           f'<w:spacing w:before="100" w:after="100" w:line="480" w:lineRule="auto"/>'
           f'<w:rPr>{FONT_PROPS}</w:rPr>')
    r = (f'<w:r><w:rPr>{FONT_PROPS}<w:rtl w:val="0"/></w:rPr>'
         f'<w:t xml:space="preserve">{text}</w:t></w:r>')
    return (f'<w:p w:rsidR="00000000" w:rsidDel="00000000" w:rsidP="00000000" '
            f'w:rsidRDefault="00000000" w:rsidRPr="00000000" w14:paraId="{pid()}">'
            f'<w:pPr>{ppr}</w:pPr>{r}</w:p>')

# ── Image dimensions (EMUs) ──────────────────────────────────────────────────
# All images at 150 DPI. Content width = 5943600 EMU (6.5 in).
# Scale each image so cx = 5943600.
def img_dims(px_w, px_h, target_cx=5943600):
    EMU_PER_PX = 914400 / 150.0
    orig_cx = int(px_w * EMU_PER_PX)
    orig_cy = int(px_h * EMU_PER_PX)
    scale = target_cx / orig_cx
    return target_cx, int(orig_cy * scale)

CX1, CY1 = img_dims(1785, 1219)   # HHI scorecard
CX2, CY2 = img_dims(1390, 1035)   # EL/EAD mismatch
CX3, CY3 = img_dims(1335, 885)    # efficient frontier
CX4, CY4 = img_dims(1785, 660)    # EWI dashboard
CX5, CY5 = img_dims(1133, 810)    # roll-rate heatmap
CX6, CY6 = img_dims(1935, 772)    # concentration limits

# ── Build L3 XML ─────────────────────────────────────────────────────────────
parts = []

# Section heading (Heading 1 in the doc uses Heading2 style per the existing structure)
parts.append(heading("L3 -- Portfolio Strategy, Concentration &amp; Early Warning Indicators", level=2))

parts.append(para(
    "This layer addresses the core portfolio management question: what book should we "
    "actively hold, and how do we detect trouble before it materialises? L3 integrates "
    "two analytical blocks -- strategic portfolio construction (Part A) and early warning "
    "monitoring (Part B) -- and draws directly on the segment-level expected loss (EL) "
    "estimates from L1 and the repricing recommendations from L2."
))

# ── PART A ──────────────────────────────────────────────────────────────────
parts.append(heading("Part A: Strategic Portfolio Construction", level=3))

# A.1 HHI
parts.append(heading("A.1  Concentration Diagnostic (HHI)", level=3))
parts.append(para(
    "We measure portfolio concentration using the Herfindahl-Hirschman Index "
    "(HHI = sum of squared market shares, scaled to 10,000). An HHI above 2,500 indicates "
    "high concentration; below 1,000 indicates competitive diversity. "
    "We compute HHI across four dimensions using 2018 Q4 portfolio data."
))

# HHI scorecard table
# Cols: Dimension | HHI | Industry Benchmark | Breach | Largest Segment | Largest Share (%)
hhi_cols = [1500, 1000, 1600, 900, 2860, 1500]
hhi_hdr  = ["Dimension", "HHI", "Industry Benchmark", "Breach?",
            "Largest Segment", "Largest Share (%)"]
hhi_data = [
    ["Purpose",  "4,074", "1,200", "YES -- BREACH", "debt_consolidation", "58.4%"],
    ["Grade",    "2,363",   "900", "YES -- BREACH", "C",                  "29.9%"],
    ["Term",     "5,003", "5,500", "No",            "36m",                "72.6%"],
    ["State",      "586",   "600", "No",            "CA",                 "13.1%"],
]
parts.append(table(hhi_cols, hhi_hdr, hhi_data))

parts.append(para(
    "Two dimensions breach their benchmarks. Purpose concentration is most severe: "
    "debt_consolidation alone accounts for 58.4% of EAD ($5.56B), producing an HHI of "
    "4,074 -- more than three times the industry benchmark of 1,200. Grade concentration "
    "is also elevated (HHI 2,363 vs. benchmark 900), driven by grade C representing "
    "29.9% of EAD. Term and geographic concentration are within acceptable bounds."
))

parts.append(image_para("rId32", 30, "image26.png", CX1, CY1))
parts.append(caption(
    "Figure L3-1. HHI Concentration Scorecard -- LendingClub 2018 Q4 active portfolio. "
    "Red bars indicate breach of industry benchmark."
))

# A.2 Concentration Limits
parts.append(heading("A.2  Concentration Limits (Policy Layer)", level=3))
parts.append(para(
    "Based on the concentration diagnostic, we propose four binding policy limits. "
    "These are calibrated to reduce purpose HHI below 2,500 and grade HHI below 1,500 "
    "over a 6-12 month re-origination cycle."
))

lim_cols = [3000, 1500, 3060, 1800]
lim_hdr  = ["Constraint", "Proposed Limit", "Current Value", "Breach by (pp)"]
lim_data = [
    ["No single purpose / EAD",    "40% or less", "debt_consol 58.4%",   "+18.4 pp"],
    ["No single grade / EAD",      "25% or less", "grade C 29.9%",        "+4.9 pp"],
    ["No single state / EAD",      "10% or less", "CA 13.1%",             "+3.1 pp"],
    ["Grade D-G combined / EAD",   "less than 20%","20.9%",               "+0.9 pp"],
]
parts.append(table(lim_cols, lim_hdr, lim_data))

parts.append(para(
    "The most consequential constraint is the purpose cap. Reducing debt_consolidation "
    "from 58.4% to 40% requires shifting approximately $1.75B of new originations toward "
    "home improvement, credit card, and car loans -- all of which show strong RAROC "
    "profiles in the L2 analysis. The grade constraint requires moderate reduction in "
    "grade C growth and reallocation toward grade B, consistent with L2's Grow "
    "recommendation for 60-month grade B loans."
))

parts.append(image_para("rId37", 35, "image31.png", CX6, CY6))
parts.append(caption(
    "Figure L3-2. Current EAD distribution by purpose and grade vs. proposed policy "
    "limits (dashed red line). Red bars breach the proposed limits."
))

# A.3 EL/EAD Mismatch
parts.append(heading("A.3  EL/EAD Contribution Mismatch", level=3))
parts.append(para(
    "For each grade x purpose segment, we compute the ratio of its EL share to its EAD "
    "share. A ratio above 1.0 indicates a segment that contributes disproportionately more "
    "to expected losses than to portfolio exposure -- these are the portfolio's hidden loss "
    "drivers. A ratio below 1.0 indicates an efficient, loss-light exposure. "
    "Formula: EL/EAD Ratio = (EL_i / Total EL) / (EAD_i / Total EAD)."
))

mis_cols = [700, 1700, 1060, 1000, 1400, 1300, 2200]
mis_hdr  = ["Grade", "Purpose", "EAD ($M)", "EL ($M)",
            "EAD Share (%)", "EL Share (%)", "EL/EAD Ratio"]
mis_data = [
    ["G", "small_business",   "0.54", "0.27", "0.01%", "0.07%", "5.96"],
    ["G", "renewable_energy", "0.17", "0.07", "0.00%", "0.02%", "5.47"],
    ["G", "medical",          "0.25", "0.10", "0.00%", "0.02%", "5.42"],
    ["G", "moving",           "0.21", "0.08", "0.00%", "0.02%", "5.12"],
    ["F", "vacation",         "0.62", "0.20", "0.01%", "0.05%", "4.65"],
    ["F", "small_business",   "1.10", "0.35", "0.01%", "0.08%", "4.37"],
    ["F", "moving",           "0.63", "0.20", "0.01%", "0.04%", "4.32"],
    ["E", "small_business",   "3.98", "1.08", "0.04%", "0.23%", "5.77"],
    ["D", "moving",           "5.90", "1.26", "0.06%", "0.27%", "4.45"],
    ["C", "small_business",   "12.4", "2.31", "0.13%", "0.50%", "3.81"],
]
parts.append(table(mis_cols, mis_hdr, mis_data))

parts.append(para(
    "Grade G and F segments consistently show EL/EAD ratios of 4-6x, confirming that "
    "these small-exposure segments punch far above their weight in expected losses. "
    "Consistent with L2's Decline recommendation for grade E-G across high-risk purposes, "
    "these segments should be excluded from new originations. The large "
    "debt_consolidation x grade C segment (EAD $765M, 36m term) -- identified in L2 as "
    "Hold/Monitor -- also has an above-average EL/EAD ratio of 1.8x, representing the "
    "single largest absolute loss contribution at approximately $32M per year."
))

parts.append(image_para("rId33", 31, "image27.png", CX2, CY2))
parts.append(caption(
    "Figure L3-3. EL/EAD Contribution Mismatch. Points above the diagonal contribute "
    "disproportionately to expected losses relative to their portfolio exposure. "
    "Bubble size proportional to EAD."
))

# A.4 Efficient Frontier
parts.append(heading("A.4  Target Portfolio Mix &amp; Efficient Frontier", level=3))
parts.append(para(
    "We construct a grade-level efficient frontier by solving a linear programme: for each "
    "target loss rate, maximise portfolio yield subject to portfolio loss rate being at or "
    "below target, weights summing to 1, and all weights being non-negative. "
    "This identifies the optimal grade mix at each risk tolerance level."
))
parts.append(para(
    "The current portfolio sits at yield = 13.06%, loss rate = 4.92%. The efficient "
    "frontier shows that the current mix is sub-optimal: by shifting grade weights from "
    "F-G (reducing from 2.1% to approximately 0.5%) into grade B-C (increasing from 57% "
    "to approximately 64-65%), the portfolio can reduce its loss rate to 3.8% while "
    "sacrificing only approximately 40-50 basis points of yield. Specifically:"
))
parts.append(bullet_para("Reduce grade F-G combined from 2.1% to approximately 0.5% of EAD"))
parts.append(bullet_para(
    "Increase grade B-C combined from 57% to approximately 64-65% of EAD"))
parts.append(bullet_para(
    "Expected outcome: yield approximately 12.6-12.7%, loss rate drops from 4.9% to 3.8%"))
parts.append(bullet_para(
    "This is consistent with L2's Grow recommendations for 60-month grade B and C loans"))

parts.append(image_para("rId34", 32, "image28.png", CX3, CY3))
parts.append(caption(
    "Figure L3-4. Portfolio Efficient Frontier (grade-level optimisation). The current "
    "portfolio (red dot) lies inside the frontier; the recommended reallocation (green star) "
    "achieves a 1.1pp lower loss rate at modest yield cost. Grey path shows LendingClub "
    "historical trajectory 2014-2018."
))

# ── PART B ──────────────────────────────────────────────────────────────────
parts.append(heading("Part B: Early Warning Indicators (EWI)", level=3))

# B.1 EWI Dashboard
parts.append(heading("B.1  Four-Tier EWI Dashboard", level=3))
parts.append(para(
    "We design a four-tier system of leading indicators to detect portfolio deterioration "
    "3-6 months before charge-offs materialise. Each tier targets a distinct early signal "
    "in the default process: loan-level delinquency, vintage performance, segment model "
    "accuracy, and macroeconomic environment. Each indicator has Yellow and Red alert "
    "thresholds and a prescribed response."
))

# EWI table -- 7 columns
ewi_cols = [1000, 1700, 1200, 1260, 1000, 820, 2380]
ewi_hdr  = ["Tier", "Indicator", "Current Value",
            "Yellow Threshold", "Red Threshold", "Status", "Action if Red"]
ewi_data = [
    ["Loan-level", "30 DPD share",
     "2.95%", "> 4%", "> 6%", "GREEN", "Tighten collections strategy"],
    ["Loan-level", "60 DPD share",
     "1.50%", "> 2%", "> 3.5%", "GREEN", "Tighten collections strategy"],
    ["Vintage", "6-month default rate vs same-vintage historical",
     "+6% YoY", "> +20% YoY", "> +40% YoY", "GREEN", "Freeze originations in affected grade"],
    ["Segment", "Grade-level EL realised vs predicted",
     "gap = 6%", "> 10%", "> 20%", "GREEN", "Re-estimate PD; escalate to credit committee"],
    ["Macro", "Unemployment rate YoY change",
     "Delta +0.3pp", "> +0.5pp", "> +1pp", "GREEN", "Activate L4 Adverse stress scenario"],
    ["Macro", "LC application volume YoY",
     "-5%", "&lt; -10% YoY", "&lt; -20% YoY", "GREEN", "Activate L4 Adverse stress scenario"],
]
parts.append(table(ewi_cols, ewi_hdr, ewi_data))

parts.append(para(
    "As of 2018 Q4, all six indicators are in the GREEN zone, consistent with a healthy "
    "credit environment. However, the segment-tier gap of 6% (realised EL exceeds "
    "model-predicted EL by 6%, as documented in L1's backtest calibration) is worth "
    "monitoring -- a further widening to 10% would trigger Yellow, indicating that the "
    "underlying PD model requires recalibration."
))
parts.append(para(
    "Each tier's Red state triggers a specific escalation: Loan-level Red -- immediate "
    "tightening of collections outreach cadence and review of servicer performance. "
    "Vintage Red -- freeze new originations in the affected grade for one cycle and review "
    "underwriting criteria. Segment Red -- re-estimate PDs using recent 6-month cohort and "
    "escalate to credit risk committee. Macro Red -- activate L4 Adverse scenario and "
    "reassess capital adequacy buffers."
))

parts.append(image_para("rId35", 33, "image29.png", CX4, CY4))
parts.append(caption(
    "Figure L3-5. Four-Tier EWI Dashboard visual summary. All tiers in GREEN as of 2018 Q4."
))

# B.2 Roll-Rate Analysis
parts.append(heading("B.2  Roll-Rate Analysis", level=3))
parts.append(para(
    "Roll-rate analysis measures the monthly probability of a loan transitioning between "
    "delinquency states: Current, 30 DPD, 60 DPD, 90 DPD, Charge-off. This is the "
    "mechanical engine underlying the EWI Loan-level tier: an unexpected surge in the "
    "30-to-60 DPD roll rate is typically visible 2-3 months before a charge-off spike."
))

rr_cols = [1360, 1334, 1334, 1334, 1334, 1334, 1330]
rr_hdr  = ["From \\ To", "Current", "30DPD", "60DPD", "90DPD", "Charge-off", "Paid-off"]
rr_data = [
    ["Current",     "0.951", "0.040", "0.000", "0.000", "0.000", "0.009"],
    ["30DPD",       "0.450", "0.050", "0.450", "0.000", "0.000", "0.050"],
    ["60DPD",       "0.150", "0.100", "0.080", "0.570", "0.000", "0.100"],
    ["90DPD",       "0.050", "0.050", "0.050", "0.100", "0.700", "0.050"],
    ["Charge-off",  "0.000", "0.000", "0.000", "0.000", "1.000", "0.000"],
    ["Paid-off",    "0.000", "0.000", "0.000", "0.000", "0.000", "1.000"],
]
parts.append(table(rr_cols, rr_hdr, rr_data))

parts.append(para(
    "Key roll rates in this environment: 30-to-60 DPD roll rate = 45% (meaning 45% of "
    "loans that enter 30-day delinquency progress to 60-day delinquency rather than "
    "curing); 60-to-90 DPD roll rate = 57%; 90 DPD to Charge-off rate = 70%; cure rate "
    "from 30 DPD = 45%. The EWI system monitors the 30-to-60 roll rate monthly -- a "
    "sudden increase (e.g., from 45% to 55%+) would signal rising credit stress "
    "2-3 months ahead of charge-off recognition."
))

parts.append(image_para("rId36", 34, "image30.png", CX5, CY5))
parts.append(caption(
    "Figure L3-6. Monthly Roll-Rate Transition Matrix heatmap. Darker colours indicate "
    "higher transition probabilities. The 90DPD-to-Charge-off cell (0.70) and "
    "30-to-60DPD cell (0.45) are key EWI monitoring targets."
))

# ── L3 Summary ───────────────────────────────────────────────────────────────
parts.append(heading("L3 Summary -- Story Line", level=3))
parts.append(para(
    "The 2018 Q4 LendingClub portfolio exhibits severe concentration in two dimensions: "
    "purpose (debt_consolidation = 58.4% of EAD, HHI 4,074) and grade "
    "(grade C = 29.9% of EAD, HHI 2,363). Approximately 30% of EAD is concentrated in "
    "a single cell -- debt_consolidation x grade C -- which, while not the worst "
    "loss-rate segment, represents the single largest source of absolute expected loss "
    "at approximately $32M per year."
))
parts.append(para(
    "Our proposed four-dimension concentration limits, if implemented over a 6-12 month "
    "origination cycle, would reduce purpose HHI from 4,074 to approximately 2,100 and "
    "grade HHI from 2,363 to approximately 1,400. The efficient frontier analysis shows "
    "the portfolio can reduce its annualised loss rate from 4.9% to 3.8% by reallocating "
    "1-2pp of EAD from grade F-G into grade B-C, at a yield cost of approximately "
    "40-50 basis points -- a highly favourable risk-adjusted trade-off."
))
parts.append(para(
    "The four-tier EWI system provides a structured early warning mechanism, enabling "
    "portfolio-level action 3-6 months before losses materialise. As of 2018 Q4, all "
    "indicators are in the GREEN zone, but the 6% segment-level model gap (L1 backtest "
    "finding) warrants close monitoring: if it widens to 10%, portfolio-level reserve and "
    "pricing assumptions should be revisited in coordination with L1 and L2 outputs."
))
parts.append(para(
    "Dependencies: L3 relies on L1 (EL by segment, backtest gap) for the EL/EAD mismatch "
    "analysis and EWI segment tier. L3 connects to L2 via concentration limits -- the "
    "repricing and growth recommendations from L2 will change the origination mix and, if "
    "implemented as suggested, will partially resolve the grade and purpose concentration "
    "breaches identified here. L4 stress scenarios are triggered by the macro tier of the "
    "EWI dashboard when unemployment rises by more than 1pp."
))

L3_XML = "\n".join(parts)

# ── Insert into document.xml ─────────────────────────────────────────────────
with open(DOC, "r", encoding="utf-8") as f:
    content = f.read()

MARKER = '<w:p w:rsidR="00000000" w:rsidDel="00000000" w:rsidP="00000000" w:rsidRDefault="00000000" w:rsidRPr="00000000" w14:paraId="00000298">'
if MARKER not in content:
    raise RuntimeError("Could not find L4 insertion marker in document.xml")

content = content.replace(MARKER, L3_XML + "\n    " + MARKER, 1)

with open(DOC, "w", encoding="utf-8") as f:
    f.write(content)
print("Inserted L3 XML into document.xml")

# ── Update _rels ─────────────────────────────────────────────────────────────
with open(RELS, "r", encoding="utf-8") as f:
    rels = f.read()

new_rels = (
    '  <Relationship Id="rId32" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image26.png"/>\n'
    '  <Relationship Id="rId33" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image27.png"/>\n'
    '  <Relationship Id="rId34" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image28.png"/>\n'
    '  <Relationship Id="rId35" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image29.png"/>\n'
    '  <Relationship Id="rId36" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image30.png"/>\n'
    '  <Relationship Id="rId37" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image31.png"/>\n'
)
rels = rels.replace("</Relationships>", new_rels + "</Relationships>")
with open(RELS, "w", encoding="utf-8") as f:
    f.write(rels)
print("Updated _rels with L3 image relationships")
print("Done!")

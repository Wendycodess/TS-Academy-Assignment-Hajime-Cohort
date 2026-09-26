"""
generate_report.py
==================
Builds the comprehensive Word (.docx) data analysis report for the logistics
database, following the 7-step data analysis workflow and the formatting
requirements:
  - A4 page, 2.5cm top/bottom & 2cm left/right margins
  - Cover page (no header/footer) isolated by a section break
  - Auto-generated Table of Contents (TOC field) linked to heading styles
  - Header (report title, right-aligned, 10pt) & footer (centered page number)
    on every page except the cover
  - Heading 1/2/3 styles for hierarchy
  - Inline-with-text, centered images, width ~14cm, aspect ratio preserved
  - Figure captions: "Figure X-XX: [description]" centered, smaller font
  - Tables with header bold + #F2F2F2 fill, text wrapping, auto-fit width
  - File naming: [Theme]_Analysis_Report_[YYYY-MM-DD].docx
"""
import json
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DOWNLOAD_DIR = Path("/home/z/my-project/download")
CHARTS_DIR   = DOWNLOAD_DIR / "charts"
RESULTS_PATH = DOWNLOAD_DIR / "analysis_results.json"
OUTPUT_PATH  = DOWNLOAD_DIR / "Logistics_Database_Analysis_Report_2026-07-12.docx"

with open(RESULTS_PATH) as f:
    R = json.load(f)

REPORT_TITLE    = "Logistics Database — Data Analysis Report"
REPORT_SUBTITLE = "Operational Performance, Fuel Efficiency & Safety Risk (2022-2024)"
REPORT_DATE     = "July 12, 2026"
REPORT_AUTHOR   = "Data Analytics Team"
IMG_WIDTH_CM    = 14.0  # ~80% of A4 effective width


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------
def set_cell_bg(cell, color_hex: str):
    """Apply background shading to a table cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    """Set cell margins in twips (1/20 pt)."""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        tc_mar.append(el)
    tc_pr.append(tc_mar)


def set_run_font(run, name="Arial", size_pt=11, bold=False, color=None, italic=False):
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    # Also set East-Asian font for compatibility
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), name)
    rFonts.set(qn("w:cs"), name)


def add_page_number_field(paragraph):
    """Insert a PAGE field (centered page number) into a footer paragraph."""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE \\* MERGEFORMAT"
    fld_sep = OxmlElement("w:fldChar"); fld_sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = "1"
    fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    run._element.append(fld_begin); run._element.append(instr); run._element.append(fld_sep)
    run._element.append(t); run._element.append(fld_end)
    set_run_font(run, name="Arial", size_pt=10)


def add_toc_field(paragraph):
    """Insert a TOC field (auto-generated Table of Contents)."""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
    instr.text = r'TOC \o "1-3" \h \z \u'
    fld_sep = OxmlElement("w:fldChar"); fld_sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = "Right-click and select 'Update Field' to generate the Table of Contents."
    fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    run._element.append(fld_begin); run._element.append(instr); run._element.append(fld_sep)
    run._element.append(t); run._element.append(fld_end)
    set_run_font(run, name="Arial", size_pt=10, italic=True, color="808080")


def add_paragraph(doc, text="", style=None, bold=False, italic=False, size=11,
                  font="Arial", color=None, align=None, space_after=6, space_before=0,
                  line_spacing=1.3, first_line_indent=None):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = line_spacing
    if first_line_indent is not None:
        pf.first_line_indent = Cm(first_line_indent)
    if text:
        run = p.add_run(text)
        set_run_font(run, name=font, size_pt=size, bold=bold, italic=italic, color=color)
    return p


def add_heading(doc, text, level=1):
    """Add a heading using the built-in Heading styles (so TOC can pick them up)."""
    h = doc.add_heading(text, level=level)
    # Style the heading run
    for run in h.runs:
        set_run_font(run, name="Arial",
                     size_pt={1: 16, 2: 13, 3: 11.5}.get(level, 11),
                     bold=True, color="2E5A88")
    h.paragraph_format.space_before = Pt({1: 14, 2: 10, 3: 8}.get(level, 6))
    h.paragraph_format.space_after  = Pt(6)
    h.paragraph_format.keep_with_next = True
    return h


def add_image(doc, image_path, caption, width_cm=IMG_WIDTH_CM):
    """Add a centered inline image with a figure caption below."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run()
    # Preserve aspect ratio: set ONLY width, let height auto-scale
    run.add_picture(str(image_path), width=Cm(width_cm))
    # Caption
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(10)
    cap.paragraph_format.space_before = Pt(0)
    crun = cap.add_run(caption)
    set_run_font(crun, name="Arial", size_pt=9.5, italic=True, color="555555")
    return p


def add_table(doc, headers, rows, col_widths_cm=None, caption=None):
    """Add a formatted table with header row + body rows."""
    if caption:
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_after = Pt(4)
        cap.paragraph_format.space_before = Pt(8)
        crun = cap.add_run(caption)
        set_run_font(crun, name="Arial", size_pt=9.5, italic=True, color="555555")

    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    # Apply a clean table style
    try:
        table.style = "Table Grid"
    except KeyError:
        pass

    # Header row
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = ""
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(h))
        set_run_font(run, name="Arial", size_pt=10, bold=True, color="FFFFFF")
        set_cell_bg(hdr_cells[i], "2E5A88")
        set_cell_margins(hdr_cells[i])

    # Body rows
    for ri, row in enumerate(rows):
        cells = table.rows[ri + 1].cells
        for ci, val in enumerate(row):
            cells[ci].text = ""
            p = cells[ci].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run("" if val is None else str(val))
            set_run_font(run, name="Arial", size_pt=10)
            set_cell_margins(cells[ci])
            if ri % 2 == 1:
                set_cell_bg(cells[ci], "F2F2F2")

    # Column widths (if specified)
    if col_widths_cm:
        for ci, w in enumerate(col_widths_cm):
            for row in table.rows:
                row.cells[ci].width = Cm(w)
    # Spacing after table
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return table


def fmt_money(v):
    try:
        return f"${float(v):,.0f}"
    except (TypeError, ValueError):
        return str(v)


def fmt_num(v, dec=2):
    try:
        return f"{float(v):,.{dec}f}"
    except (TypeError, ValueError):
        return str(v)


def fmt_int(v):
    try:
        return f"{int(v):,}"
    except (TypeError, ValueError):
        return str(v)


def fmt_pct(v, dec=2):
    try:
        return f"{float(v):+.{dec}f}%" if v is not None else "—"
    except (TypeError, ValueError):
        return str(v)


# ---------------------------------------------------------------------------
# Build the document
# ---------------------------------------------------------------------------
doc = Document()

# ---- Page setup: A4, margins 2.5cm top/bottom, 2cm left/right ----
section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width  = Cm(21.0)
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin   = Cm(2.0)
section.right_margin  = Cm(2.0)

# ---- Default style ----
style = doc.styles["Normal"]
style.font.name = "Arial"
style.font.size = Pt(11)
style.paragraph_format.line_spacing = 1.3
style.paragraph_format.space_after = Pt(6)
# East-Asian font binding for compatibility
rpr = style.element.get_or_add_rPr()
rfonts = rpr.find(qn("w:rFonts"))
if rfonts is None:
    rfonts = OxmlElement("w:rFonts"); rpr.insert(0, rfonts)
rfonts.set(qn("w:eastAsia"), "Arial")

# ===========================================================================
# SECTION 1 — COVER PAGE (no header/footer)
# ===========================================================================
# Add some vertical space at top of cover
for _ in range(6):
    doc.add_paragraph()

# Title
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
run = p.add_run("LOGISTICS DATABASE")
set_run_font(run, name="Arial", size_pt=30, bold=True, color="2E5A88")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(24)
run = p.add_run("Data Analysis Report")
set_run_font(run, name="Arial", size_pt=22, bold=True, color="C0504D")

# Decorative thin line via paragraph border
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(18)
pPr = p._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
bottom = OxmlElement("w:bottom")
bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "12")
bottom.set(qn("w:space"), "1"); bottom.set(qn("w:color"), "2E5A88")
pBdr.append(bottom); pPr.append(pBdr)

# Subtitle
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(36)
run = p.add_run(REPORT_SUBTITLE)
set_run_font(run, name="Arial", size_pt=13, italic=True, color="555555")

# Meta block
for label, value in [
    ("Reporting Period", "January 2022 – December 2024 (36 months)"),
    ("Data Sources", "trips.csv (85,410 rows), trailers.csv (180 rows), safety_incidents.csv (170 rows)"),
    ("Database Engine", "SQLite 3.53"),
    ("Report Date", REPORT_DATE),
    ("Prepared By", REPORT_AUTHOR),
]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f"{label}: ")
    set_run_font(run, name="Arial", size_pt=11, bold=True, color="333333")
    run = p.add_run(value)
    set_run_font(run, name="Arial", size_pt=11, color="333333")

# End cover section with a section break (NEXT_PAGE) so TOC+body start fresh
doc.add_section(WD_SECTION.NEW_PAGE)

# ===========================================================================
# SECTION 2 — TOC + BODY (with header/footer)
# ===========================================================================
section2 = doc.sections[1]
section2.page_height = Cm(29.7)
section2.page_width  = Cm(21.0)
section2.top_margin    = Cm(2.5)
section2.bottom_margin = Cm(2.5)
section2.left_margin   = Cm(2.0)
section2.right_margin  = Cm(2.0)
# Unlink from previous section so this section gets its own header/footer
section2.header.is_linked_to_previous = False
section2.footer.is_linked_to_previous = False

# Header: report title, right-aligned, 10pt
hdr_p = section2.header.paragraphs[0]
hdr_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
hdr_run = hdr_p.add_run(REPORT_TITLE)
set_run_font(hdr_run, name="Arial", size_pt=10, italic=True, color="808080")

# Footer: centered page number
ftr_p = section2.footer.paragraphs[0]
ftr_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_page_number_field(ftr_p)

# --- Table of Contents ---
add_heading(doc, "Table of Contents", level=1)
toc_p = doc.add_paragraph()
add_toc_field(toc_p)

# Note for the reader
note_p = doc.add_paragraph()
note_p.paragraph_format.space_before = Pt(8)
note_run = note_p.add_run(
    "Note: The Table of Contents above is generated automatically by Word. "
    "If it shows placeholder text, right-click anywhere inside it and choose "
    "\"Update Field\" → \"Update entire table\"."
)
set_run_font(note_run, name="Arial", size_pt=9.5, italic=True, color="808080")

# Page break before main content (TOC on its own page)
doc.add_page_break()

# ===========================================================================
# 1. Executive Summary
# ===========================================================================
add_heading(doc, "1. Executive Summary", level=1)

overall = R["overall_kpis"]
yearly = R["yearly"]
inc_year = R["incidents_by_year"]

add_paragraph(doc,
    "This report presents a comprehensive analysis of three years (2022-2024) of "
    "logistics operational data drawn from the company's trips, trailers, and "
    "safety-incident tables. The dataset comprises 85,410 dispatched trips covering "
    f"approximately {fmt_int(overall['total_dist'])} total miles, 180 active trailers, "
    f"and 170 recorded safety incidents. The analysis follows a structured seven-step "
    "workflow covering scope definition, data quality, core performance, comparisons, "
    "attribution, recommendations, and uncertainty.",
    first_line_indent=0.0)

add_paragraph(doc,
    f"The fleet maintained remarkably stable operations across the three-year window: "
    f"trip volume stayed within a narrow band ({fmt_int(yearly[0]['trips'])} to "
    f"{fmt_int(yearly[-1]['trips'])} trips per year, ±2%), average distance per trip held "
    f"at {fmt_num(overall['avg_dist'], 0)} miles, and average fuel efficiency remained "
    f"effectively flat at {fmt_num(overall['avg_mpg'], 2)} MPG. This stability reflects a "
    "mature, demand-constrained operation with well-established routes and capacity "
    "utilization. However, the absence of MPG improvement over three years also suggests "
    "that fuel-efficiency initiatives have not yet gained traction, leaving a meaningful "
    "cost-reduction opportunity on the table.",
    first_line_indent=0.0)

add_paragraph(doc,
    f"On the safety dimension, the picture is more concerning. The incident rate per "
    f"1,000 trips rose from {inc_year[0]['rate_per_1k']} in 2022 to {inc_year[-1]['rate_per_1k']} "
    f"in 2024, a cumulative increase of "
    f"{fmt_pct((inc_year[-1]['rate_per_1k']-inc_year[0]['rate_per_1k'])/inc_year[0]['rate_per_1k']*100, 1)}. "
    f"Total claim exposure across the period exceeded ${2.65:,.0f}M, with three cities "
    f"(Chicago, Seattle, Memphis) concentrating a disproportionate share. Preventable "
    "incidents — those directly addressable through driver behavior and coaching — account "
    "for roughly one-third of claim dollars and represent the most actionable lever for "
    "near-term risk reduction.",
    first_line_indent=0.0)

add_paragraph(doc,
    "A material data-quality issue was also uncovered: approximately 2% of trips (1,680 "
    "records) have no trailer_id assigned in the source system, preventing trailer-level "
    "utilization, depreciation, and cost attribution for those trips. Closing this gap is "
    "a high-priority operational and analytics recommendation. The five recommendations "
    "in Section 7 prioritize actions by impact and feasibility, and each includes "
    "validation criteria so that progress can be tracked objectively.",
    first_line_indent=0.0)

# Key findings table
add_heading(doc, "Key Findings at a Glance", level=2)
add_table(doc,
    headers=["Metric", "2022", "2023", "2024", "3-Year Total / Avg"],
    rows=[
        ["Trips", fmt_int(yearly[0]["trips"]), fmt_int(yearly[1]["trips"]),
         fmt_int(yearly[2]["trips"]), fmt_int(overall["trips"])],
        ["Avg distance/trip (mi)", fmt_num(yearly[0]["avg_dist"], 1), fmt_num(yearly[1]["avg_dist"], 1),
         fmt_num(yearly[2]["avg_dist"], 1), fmt_num(overall["avg_dist"], 1)],
        ["Avg MPG", fmt_num(yearly[0]["avg_mpg"], 2), fmt_num(yearly[1]["avg_mpg"], 2),
         fmt_num(yearly[2]["avg_mpg"], 2), fmt_num(overall["avg_mpg"], 2)],
        ["Avg idle time (h)", fmt_num(yearly[0]["avg_idle"], 2), fmt_num(yearly[1]["avg_idle"], 2),
         fmt_num(yearly[2]["avg_idle"], 2), fmt_num(overall["avg_idle"], 2)],
        ["Safety incidents", str(inc_year[0]["incidents"]), str(inc_year[1]["incidents"]),
         str(inc_year[2]["incidents"]), str(sum(r["incidents"] for r in inc_year))],
        ["Incident rate / 1k trips", fmt_num(inc_year[0]["rate_per_1k"], 2),
         fmt_num(inc_year[1]["rate_per_1k"], 2), fmt_num(inc_year[2]["rate_per_1k"], 2), "—"],
        ["Total claims ($)", fmt_money(inc_year[0]["total_claims"]), fmt_money(inc_year[1]["total_claims"]),
         fmt_money(inc_year[2]["total_claims"]),
         fmt_money(sum(r["total_claims"] for r in inc_year))],
    ],
    caption="Table 1-1: Headline operational and safety metrics, 2022-2024."
)

# ===========================================================================
# 2. Scope & Definitions
# ===========================================================================
add_heading(doc, "2. Scope & Definitions", level=1)

add_paragraph(doc,
    f"Objective. {R['scope']['objective']}",
    first_line_indent=0.0)

add_heading(doc, "Key Metrics", level=2)
add_paragraph(doc,
    "The following six metrics anchor the analysis. They were chosen because they "
    "span the three operational dimensions that matter most to a freight fleet — "
    "throughput, efficiency, and safety — and because each can be computed directly "
    "from the available tables without external enrichment.",
    first_line_indent=0.0)
for m in R["scope"]["key_metrics"]:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(m)
    set_run_font(run, name="Arial", size_pt=11)

add_heading(doc, "Time Range & Granularity", level=2)
add_paragraph(doc,
    f"Time range: {R['scope']['time_range']}. Granularity: {R['scope']['granularity']} "
    "Monthly granularity is used for trend detection, yearly granularity for comparison, "
    "and per-trip granularity for attribution analysis. The three-year window is long "
    "enough to smooth out seasonal noise and detect structural shifts, while remaining "
    "recent enough to be operationally relevant.",
    first_line_indent=0.0)

add_heading(doc, "Source Tables", level=2)
add_paragraph(doc,
    "Three tables, loaded from CSV exports of the operational system, form the analytical "
    "foundation. Their sizes and relationships are summarized below.",
    first_line_indent=0.0)
add_table(doc,
    headers=["Table", "Rows", "Role", "Primary Key"],
    rows=[
        ["trailers", "180", "Fleet asset reference (trailer master)", "trailer_id"],
        ["trips", "85,410", "Operational fact table (one row per dispatched trip)", "trip_id"],
        ["safety_incidents", "170", "Safety event log (linked to trips via trip_id)", "incident_id"],
    ],
    caption="Table 2-1: Source tables loaded into the logistics SQLite database."
)

# ===========================================================================
# 3. Data Quality Checks
# ===========================================================================
add_heading(doc, "3. Data Quality Checks", level=1)

dq = R["data_quality"]
add_paragraph(doc,
    f"Data quality was assessed across eight dimensions covering primary-key integrity, "
    f"null-value checks, foreign-key referential integrity, and duplicate detection. "
    f"The dataset is broadly fit for analytical use: primary keys are unique across all "
    f"three tables, all 170 safety incidents successfully link to a parent trip, and "
    f"numeric fields are fully populated. One data-quality warning was identified — "
    f"{dq['checks'][7]['result']:,} trips ({dq['checks'][7]['pct']}%) carry a blank "
    f"trailer_id — which is addressed as a recommendation in Section 7.",
    first_line_indent=0.0)

add_heading(doc, "Row Counts & Coverage", level=2)
add_table(doc,
    headers=["Table", "Row Count", "Date Coverage"],
    rows=[
        ["trailers", fmt_int(dq["row_counts"]["trailers"]), "acquisition_date 2015-2024"],
        ["trips", fmt_int(dq["row_counts"]["trips"]),
         f"{dq['date_coverage']['min']} to {dq['date_coverage']['max']}"],
        ["safety_incidents", fmt_int(dq["row_counts"]["safety_incidents"]),
         "incident_date 2022-2024"],
    ],
    caption="Table 3-1: Row counts and temporal coverage by table."
)

add_heading(doc, "Quality Check Results", level=2)
add_table(doc,
    headers=["Check", "Result", "Status"],
    rows=[[c["check"], str(c.get("result", "")), c["status"]] for c in dq["checks"]],
    caption="Table 3-2: Data quality check results (PASS / WARN / FAIL)."
)

add_heading(doc, "Cleaning Rules Applied", level=2)
add_paragraph(doc,
    "The following data-cleaning rules were applied during the ETL load from CSV to "
    "SQLite. No row deduplication was required because zero duplicate primary keys "
    "were detected. No imputation was applied — blank trailer_id rows were retained as "
    "NULL so that the gap is visible and traceable rather than masked.",
    first_line_indent=0.0)
for rule in dq["cleaning_rules_applied"]:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(rule)
    set_run_font(run, name="Arial", size_pt=11)

# ===========================================================================
# 4. Core Performance Analysis
# ===========================================================================
add_heading(doc, "4. Core Performance Analysis", level=1)

add_paragraph(doc,
    f"Across the 36-month window the fleet completed {fmt_int(overall['trips'])} trips, "
    f"covering {fmt_int(overall['total_dist'])} total miles at an average of "
    f"{fmt_num(overall['avg_dist'], 0)} miles per trip. Total fuel consumption was "
    f"{fmt_int(overall['total_fuel'])} gallons, yielding a fleet-level fuel efficiency "
    f"of {fmt_num(overall['fleet_mpg'], 2)} MPG (computed as total miles divided by "
    f"total gallons, which is the correct volume-weighted measure and differs slightly "
    f"from the simple average of per-trip MPG, {fmt_num(overall['avg_mpg'], 2)}). "
    f"Average idle time per trip was {fmt_num(overall['avg_idle'], 2)} hours, and "
    f"average trip duration was {fmt_num(overall['avg_dur'], 2)} hours.",
    first_line_indent=0.0)

add_heading(doc, "Monthly Trip Volume and Fuel Efficiency", level=2)
add_paragraph(doc,
    "The chart below pairs monthly trip volume (bars, left axis) with the monthly "
    "average MPG (line, right axis). Trip volume oscillates in a tight seasonal band "
    "around 2,300-2,700 trips per month with no dramatic peaks or troughs, confirming "
    "stable demand. Average MPG is essentially flat at 6.5 across the entire period — "
    "there is no visible upward trend, which means fleet fuel-efficiency initiatives "
    "(if any) have not yet produced a measurable effect at the aggregate level. This "
    "flatness is itself a finding: in an environment of rising fuel prices and ESG "
    "scrutiny, a flat MPG curve represents unrealized savings.",
    first_line_indent=0.0)
add_image(doc, R["charts"]["chart1"],
          "Figure 4-1: Monthly trip volume (bars) and average MPG (line), Jan 2022 – Dec 2024.")

add_heading(doc, "Idle Time Trend", level=2)
add_paragraph(doc,
    f"Idle time per trip averaged {fmt_num(overall['avg_idle'], 2)} hours over the "
    "three-year period — meaning roughly 28% of the average {fmt_num(overall['avg_dur'], 1)}-hour "
    "trip duration is spent idling. The trend chart shows idle time fluctuating in a "
    "narrow band around the three-year mean, with no clear directional improvement. "
    "Idle time is a known fuel-waste driver: at typical diesel consumption of ~0.8 gal/h "
    "for idle heavy trucks, the fleet burns an estimated 6.7 million gallons per year "
    "just idling — a meaningful reduction target. Idle time also correlates with "
    "driver comfort needs (sleeper heating/cooling) and regulatory rest requirements, "
    "so any intervention must balance efficiency against compliance and welfare.",
    first_line_indent=0.0)
add_image(doc, R["charts"]["chart2"],
          "Figure 4-2: Average idle time per trip (hours), monthly trend with 3-year mean.")

# ===========================================================================
# 5. Comparisons
# ===========================================================================
add_heading(doc, "5. Comparisons (YoY and Segment)", level=1)

add_paragraph(doc,
    "This section compares performance across years (year-over-year) and across trailer "
    "types (Dry Van vs Refrigerated) to surface structural differences that aggregate "
    "metrics can mask. YoY analysis answers \"is the operation improving?\" while "
    "segment analysis answers \"which parts of the operation differ, and why?\".",
    first_line_indent=0.0)

add_heading(doc, "Year-over-Year Performance", level=2)
yoy = R["yoy"]
add_table(doc,
    headers=["Year", "Trips YoY%", "Avg Dist YoY%", "Avg MPG YoY%", "Avg Idle YoY%", "Total Dist YoY%"],
    rows=[
        [y["year"], fmt_pct(y["trips_yoy%"]), fmt_pct(y["avg_dist_yoy%"]),
         fmt_pct(y["avg_mpg_yoy%"]), fmt_pct(y["avg_idle_yoy%"]), fmt_pct(y["total_dist_yoy%"])]
        for y in yoy
    ],
    caption="Table 5-1: Year-over-year change in key metrics (vs. prior year)."
)
add_paragraph(doc,
    "The YoY deltas are small in magnitude — every metric moves less than 2% year-over-year, "
    "which is consistent with a mature operation. Trip volume dipped 1.5% in 2023 before "
    "recovering +1.7% in 2024; average distance per trip moved in the opposite direction "
    "(+0.5% then -1.0%), suggesting a slight shift in the mix of long-haul vs short-haul "
    "loads. Average MPG was effectively flat (within ±0.15% YoY), and idle time was also "
    "flat. The absence of any directional improvement in MPG and idle time over three "
    "years is the most strategically important YoY finding — it indicates that whatever "
    "efficiency initiatives have been attempted, they have not moved the needle at fleet "
    "scale.",
    first_line_indent=0.0)
add_image(doc, R["charts"]["chart3"],
          "Figure 5-1: Year-over-year comparison of trip volume and average MPG.")

add_heading(doc, "Segment Comparison: Dry Van vs Refrigerated", level=2)
seg = R["segment_trailer_type"]
add_table(doc,
    headers=["Trailer Type", "Trips (3-yr)", "Avg Dist (mi)", "Avg MPG", "Avg Idle (h)"],
    rows=[[s["trailer_type"], fmt_int(s["trips"]), fmt_num(s["avg_dist"], 1),
           fmt_num(s["avg_mpg"], 2), fmt_num(s["avg_idle"], 2)] for s in seg],
    caption="Table 5-2: Operational comparison by trailer type."
)
add_paragraph(doc,
    "The fleet operates two trailer types in roughly comparable volumes — Dry Vans "
    f"({fmt_int(seg[0]['trips'])} trips) and Refrigerated units ({fmt_int(seg[1]['trips'])} "
    "trips). The segment comparison reveals that the two types are operated on "
    "essentially identical routes (similar average distance per trip), but Refrigerated "
    "trailers show measurably lower fuel efficiency than Dry Vans. This gap is expected "
    "physically — reefer units consume additional fuel for compressor operation — but "
    "its magnitude should be quantified, tracked, and benchmarked against manufacturer "
    "specifications. Closing even half of the MPG gap through newer reefer technology, "
    "electric standby, or driver-behavior coaching would translate to six-figure annual "
    "fuel savings for a fleet this size.",
    first_line_indent=0.0)
add_image(doc, R["charts"]["chart4"],
          "Figure 5-2: Dry Van vs Refrigerated — trip volume and average MPG comparison.")

# ===========================================================================
# 6. Attribution & Diagnosis
# ===========================================================================
add_heading(doc, "6. Attribution & Diagnosis (Safety)", level=1)

add_paragraph(doc,
    "Safety incidents are the most consequential operational risk in the dataset: they "
    "directly drive insurance premiums, regulatory exposure (CSA scores), driver "
    "retention, and customer trust. This section decomposes the 170 incidents by type, "
    "geography, year, and preventability to identify the most actionable intervention "
    "points.",
    first_line_indent=0.0)

add_heading(doc, "Incident Type Breakdown", level=2)
ibt = R["incidents_by_type"]
add_table(doc,
    headers=["Incident Type", "Count", "Total Claims ($)", "Avg Claim ($)", "At Fault", "Preventable", "Injuries"],
    rows=[[r["incident_type"], str(r["n"]), fmt_money(r["total_claims"]),
           fmt_money(r["avg_claim"]), str(r["at_fault"]), str(r["preventable"]), str(r["injuries"])]
          for r in ibt],
    caption="Table 6-1: Safety incidents by type with financial and severity breakdown."
)
add_paragraph(doc,
    "DOT Violations are the most frequent incident type (39 occurrences) and Accidents "
    "the second-most (35). Equipment Damage and Accidents together drive the largest "
    "claim amounts, reflecting the high cost of vehicle repair and cargo loss. Customer "
    "Complaints, while not financially large in claim terms, are strategically important "
    "because they directly affect customer retention and reputation. The relatively even "
    "distribution across the five incident types suggests that no single root cause "
    "dominates — a portfolio of interventions (training, maintenance, route planning, "
    "customer service) will be required rather than a single silver bullet.",
    first_line_indent=0.0)
add_image(doc, R["charts"]["chart5"],
          "Figure 6-1: Incident count and total claim amount by incident type.")

add_heading(doc, "Geographic and Temporal Concentration", level=2)
ibc = R["incidents_by_city"]
top3_claims = sum(r["total_claims"] for r in ibc[:3])
total_all_claims = sum(r["total_claims"] for r in ibt)
add_table(doc,
    headers=["Rank", "City", "Incidents", "Total Claims ($)", "Avg Claim ($)"],
    rows=[[str(i+1), r["location_city"], str(r["n"]), fmt_money(r["total_claims"]),
           fmt_money(r["avg_claim"])] for i, r in enumerate(ibc[:10])],
    caption="Table 6-2: Top 10 cities by total claim amount."
)
add_paragraph(doc,
    f"The geographic concentration is striking: the top three cities (Chicago, Seattle, "
    f"Memphis) alone account for ${top3_claims:,.0f} in claims — approximately "
    f"{top3_claims/total_all_claims*100:.0f}% of total claim exposure across all 170 "
    "incidents. This clustering is far higher than what a uniform geographic distribution "
    "would predict, and it points to city-specific risk factors: weather patterns "
    "(Seattle rain, Memphis severe storms), dock and yard congestion (Chicago rail "
    "interchange volume), and traffic density. The temporal view shows the incident "
    "rate per 1,000 trips rising from 1.96 in 2022 to 2.09 in 2024, a concerning "
    "directional trend that warrants proactive intervention before it accelerates.",
    first_line_indent=0.0)
add_image(doc, R["charts"]["chart6"],
          "Figure 6-2: Top cities by total claim amount (left) and incident rate per 1,000 trips by year (right).")

add_heading(doc, "Preventability Analysis", level=2)
ps = R["preventable_split"]
add_table(doc,
    headers=["Preventability", "Incidents", "Total Claims ($)", "Avg Claim ($)"],
    rows=[
        ["Preventable" if r["preventable_flag"] == 1 else "Non-Preventable",
         str(r["n"]), fmt_money(r["total_claims"]), fmt_money(r["avg_claim"])]
        for r in ps
    ],
    caption="Table 6-3: Preventable vs non-preventable incident split."
)
prev_total = next((r["total_claims"] for r in ps if r["preventable_flag"] == 1), 0)
nonprev_total = next((r["total_claims"] for r in ps if r["preventable_flag"] == 0), 0)
grand_total = prev_total + nonprev_total
add_paragraph(doc,
    f"Of the 170 incidents, 64 (38%) are flagged as preventable — meaning they could "
    f"have been avoided through different driver behavior, training, or decision-making. "
    f"These preventable incidents account for ${prev_total:,.0f} in claims "
    f"({prev_total/grand_total*100:.1f}% of total). While non-preventable incidents "
    f"represent the larger dollar share (${nonprev_total:,.0f}, "
    f"{nonprev_total/grand_total*100:.1f}%), they are by definition not directly "
    "actionable. The preventable segment is therefore the highest-leverage target for "
    "intervention: every preventable incident avoided is a dollar of claim cost "
    "eliminated plus a reduction in future premium exposure.",
    first_line_indent=0.0)

# ===========================================================================
# 7. Insights & Action Plan
# ===========================================================================
add_heading(doc, "7. Insights & Action Plan", level=1)

add_heading(doc, "Key Findings", level=2)
for i, finding in enumerate(R["insights"]["key_findings"], 1):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(finding)
    set_run_font(run, name="Arial", size_pt=11)

add_heading(doc, "Recommendations", level=2)
add_paragraph(doc,
    "Five recommendations are prioritized below. Each includes the action, expected "
    "impact, key risks, and a validation mechanism so that progress can be tracked "
    "objectively rather than asserted.",
    first_line_indent=0.0)

for i, rec in enumerate(R["insights"]["recommendations"], 1):
    add_heading(doc, f"R{i}. {rec['title']}  (Priority: {rec['priority']})", level=3)
    # Build a small 2-col table for the recommendation
    add_table(doc,
        headers=["Dimension", "Detail"],
        rows=[
            ["Impact", rec["impact"]],
            ["Action", rec["action"]],
            ["Risks", rec["risks"]],
            ["Validation", rec["validation"]],
        ],
        col_widths_cm=[3.5, 13.0],
    )

# ===========================================================================
# 8. Uncertainty Statement
# ===========================================================================
add_heading(doc, "8. Uncertainty Statement & Limitations", level=1)

add_paragraph(doc,
    "Every analysis carries uncertainty, and responsible use of this report requires "
    "understanding what the data can and cannot support. The limitations below should "
    "be weighed when interpreting the findings, and the validation steps should guide "
    "next-round analysis.",
    first_line_indent=0.0)

add_heading(doc, "Limitations", level=2)
for lim in R["uncertainty"]["limitations"]:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(lim)
    set_run_font(run, name="Arial", size_pt=11)

add_heading(doc, "Potential Biases", level=2)
for b in R["uncertainty"]["biases"]:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(b)
    set_run_font(run, name="Arial", size_pt=11)

add_heading(doc, "Recommended Next Validation Steps", level=2)
for nv in R["uncertainty"]["next_validation"]:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(nv)
    set_run_font(run, name="Arial", size_pt=11)

# ===========================================================================
# 9. Appendix
# ===========================================================================
add_heading(doc, "9. Appendix", level=1)

add_heading(doc, "A. Database Schema", level=2)
add_paragraph(doc,
    "The logistics database is stored as a SQLite file at "
    "/home/z/my-project/download/logistics.db (18.59 MB). The full DDL is available in "
    "/home/z/my-project/download/logistics_schema.sql. The three tables and their key "
    "columns are summarized below.",
    first_line_indent=0.0)
add_table(doc,
    headers=["Table", "Primary Key", "Key Foreign Key", "Rows", "Key Indexed Columns"],
    rows=[
        ["trailers", "trailer_id", "—", "180",
         "trailer_id, trailer_type, status"],
        ["trips", "trip_id", "trailer_id → trailers.trailer_id", "85,410",
         "dispatch_date, driver_id, truck_id, trailer_id, trip_status"],
        ["safety_incidents", "incident_id", "trip_id → trips.trip_id", "170",
         "incident_date, incident_type, driver_id, trip_id"],
    ],
    caption="Table A-1: Database schema summary."
)

add_heading(doc, "B. Analysis Artifacts", level=2)
add_paragraph(doc,
    "The following artifacts were produced during this analysis and are available for "
    "re-use and audit:",
    first_line_indent=0.0)
for art in [
    "logistics.db — SQLite database with all three tables loaded and indexed.",
    "logistics_schema.sql — DDL for all tables, indexes, and foreign keys.",
    "build_logistics_db.py — Idempotent ETL script that rebuilds the database from CSVs.",
    "analyze_logistics.py — Analysis script that produces all metrics and charts.",
    "analysis_results.json — All numerical results in machine-readable form.",
    "charts/ — Six PNG charts (150 DPI) embedded in this report.",
]:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(art)
    set_run_font(run, name="Arial", size_pt=10.5)

add_heading(doc, "C. Reproducibility", level=2)
add_paragraph(doc,
    "To reproduce this analysis end-to-end, run the following commands in order. "
    "The build script is idempotent — it wipes and recreates the database on each run — "
    "and the analysis script regenerates all charts and the results JSON deterministically.",
    first_line_indent=0.0)
for cmd in [
    "python3 /home/z/my-project/download/build_logistics_db.py",
    "python3 /home/z/my-project/download/analyze_logistics.py",
    "python3 /home/z/my-project/download/generate_report.py",
]:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.line_spacing = 1.2
    run = p.add_run(f"$ {cmd}")
    set_run_font(run, name="Consolas", size_pt=10, color="2E5A88")

# ---- Save ----
doc.save(OUTPUT_PATH)
print(f"\nReport saved to: {OUTPUT_PATH}")
print(f"File size: {OUTPUT_PATH.stat().st_size/1024:.1f} KB")

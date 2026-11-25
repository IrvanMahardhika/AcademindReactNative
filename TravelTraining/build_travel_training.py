#!/usr/bin/env python3
"""Build Travel Training source workbook, CC0-style globe logo, gold chart PNG, and gold PDF."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
import cairosvg
from PIL import Image
from reportlab.lib.colors import Color, HexColor, white
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import cm, inch
from reportlab.pdfgen import canvas
from reportlab.platypus import Table as RLTable
from reportlab.platypus import TableStyle

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
NAVY = "1B4F72"
NAVY_HEX = HexColor("#1B4F72")
PAGE = landscape((13.333 * inch, 7.5 * inch))

ROWS = [
    # Employee, Paris, Tokyo, Bali, Cancun, Rome
    ("Avery Cole", 94, 91, 88, 92, 90),
    ("Blake Diaz", 86, 84, 81, 88, 82),
    ("Casey Ng", 96, 93, 95, 90, 91),
    ("Drew Patel", 72, 78, 70, 74, 76),
    ("Eden Park", 62, 68, 65, 60, 64),
    ("Finley Ross", 80, 85, 83, 79, 81),
    ("Gray Kim", 71, 69, 73, 75, 72),
    ("Harper Lee", 98, 95, 97, 94, 96),
    ("Indigo Shah", 58, 64, 61, 55, 59),
    ("Jamie Wu", 88, 82, 85, 87, 83),
]

BANDS = ["Expert", "Advanced", "Proficient", "Needs Improvement"]


def band_for(avg: float) -> str:
    if avg >= 90:
        return "Expert"
    if avg >= 80:
        return "Advanced"
    if avg >= 70:
        return "Proficient"
    return "Needs Improvement"


def computed_rows():
    out = []
    for name, *scores in ROWS:
        avg = round(sum(scores) / 5, 1)
        out.append((name, *scores, avg, band_for(avg)))
    return out


# Public domain schematic globe (Sarang, 2012):
# https://commons.wikimedia.org/wiki/File:Globe_icon-white.svg
GLOBE_SVG = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="420" height="420">
  <circle cx="210" cy="210" r="208" fill="#FFFFFF"/>
  <path stroke="#1B4F72" fill="none" stroke-width="20"
    d="M209,15a195,195 0 1,0 2,0zm1,0v390m195-195H15M59,90a260,260 0 0,0 302,0 m0,240 a260,260 0 0,0-302,0M195,20a250,250 0 0,0 0,382 m30,0 a250,250 0 0,0 0-382"/>
</svg>
"""


def draw_logo(path: Path) -> None:
    """Rasterize Wikimedia File:Globe_icon-white.svg (public domain) in navy."""
    svg_path = path.with_name("Globe_icon-white.svg")
    svg_path.write_text(GLOBE_SVG, encoding="utf-8")
    cairosvg.svg2png(bytestring=GLOBE_SVG.encode("utf-8"), write_to=str(path), output_width=512)


def style_header(ws, cells):
    fill = PatternFill("solid", fgColor=NAVY)
    font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    for c in cells:
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal="center", wrap_text=True)


def build_source_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Scores"
    headers = ["Employee", "Paris", "Tokyo", "Bali", "Cancun", "Rome"]
    ws.append(headers)
    for row in ROWS:
        ws.append(list(row))
    style_header(ws, [ws.cell(1, i) for i in range(1, 7)])
    for col in range(1, 7):
        ws.column_dimensions[chr(64 + col)].width = 16 if col == 1 else 12
    tab = Table(displayName="ScoresTable", ref="A1:F11")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    ws.sheet_properties.tabColor = NAVY
    wb.save(path)


def build_gold_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Scores"
    headers = [
        "Employee",
        "Paris",
        "Tokyo",
        "Bali",
        "Cancun",
        "Rome",
        "Average",
        "Band",
    ]
    ws.append(headers)
    for i, row in enumerate(ROWS, start=2):
        ws.append([row[0], *row[1:]])
        ws.cell(i, 7).value = f"=AVERAGE(B{i}:F{i})"
        ws.cell(i, 7).number_format = "0.0"
        ws.cell(
            i, 8
        ).value = (
            f'=IF(G{i}>=90,"Expert",IF(G{i}>=80,"Advanced",IF(G{i}>=70,"Proficient","Needs Improvement")))'
        )
    style_header(ws, [ws.cell(1, i) for i in range(1, 9)])
    for col, w in enumerate([16, 12, 12, 12, 12, 12, 12, 20], start=1):
        ws.column_dimensions[chr(64 + col)].width = w
    tab = Table(displayName="ScoresTable", ref="A1:H11")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)

    gs = wb.create_sheet("GradeSummary")
    gs.append(["Band", "Count"])
    for i, band in enumerate(BANDS, start=2):
        gs.cell(i, 1).value = band
        gs.cell(i, 2).value = f'=COUNTIF(ScoresTable[Band],A{i})'
    style_header(gs, [gs.cell(1, 1), gs.cell(1, 2)])
    gs.column_dimensions["A"].width = 22
    gs.column_dimensions["B"].width = 12
    gst = Table(displayName="GradeSummaryTable", ref="A1:B5")
    gst.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    gs.add_table(gst)

    chart = BarChart()
    chart.type = "col"
    chart.title = "Grade distribution"
    chart.y_axis.title = "Count"
    chart.x_axis.title = None
    data = Reference(gs, min_col=2, min_row=1, max_row=5)
    cats = Reference(gs, min_col=1, min_row=2, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.legend = None
    chart.series[0].graphicalProperties.solidFill = NAVY
    gs.add_chart(chart, "D2")
    ws.sheet_properties.tabColor = NAVY
    gs.sheet_properties.tabColor = NAVY
    wb.save(path)


def build_chart_png(path: Path) -> dict:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    computed = computed_rows()
    counts = {b: 0 for b in BANDS}
    for r in computed:
        counts[r[-1]] += 1
    fig, ax = plt.subplots(figsize=(10.2, 5.4), dpi=140)
    xs = BANDS
    ys = [counts[b] for b in BANDS]
    bars = ax.bar(xs, ys, color="#1B4F72", width=0.62)
    ax.set_title("Grade distribution", fontsize=16, color="#1B4F72", pad=12)
    ax.set_ylabel("Count")
    ax.set_ylim(0, max(ys) + 1.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for bar, y in zip(bars, ys):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y + 0.08,
            str(y),
            ha="center",
            va="bottom",
            color="#1B4F72",
            fontsize=11,
        )
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return counts


def footer(c: canvas.Canvas, w, h):
    c.setFillColor(NAVY_HEX)
    c.rect(0, 0, w, 0.42 * inch, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica", 9)
    c.drawString(0.5 * inch, 0.16 * inch, "Confidential Training Operations")
    c.drawRightString(w - 0.5 * inch, 0.16 * inch, c.getPageNumber().__str__())


def place_logo(c: canvas.Canvas, logo: Path, w, h):
    size = 2.0 * cm
    c.drawImage(
        str(logo),
        w - size - 0.4 * cm,
        h - size - 0.4 * cm,
        width=size,
        height=size,
        mask="auto",
        preserveAspectRatio=True,
        anchor="c",
    )


def build_pdf(path: Path, logo: Path, chart: Path, counts: dict) -> None:
    computed = computed_rows()
    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle("Destination Specialist Results")
    c.setAuthor("Training Operations")

    # Slide 1 — title
    c.setFillColor(NAVY_HEX)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    place_logo(c, logo, w, h)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 36)
    c.drawString(0.7 * inch, h / 2 + 0.35 * inch, "Destination Specialist Results")
    c.setFont("Helvetica", 18)
    c.drawString(0.7 * inch, h / 2 - 0.15 * inch, "Certification cohort — five destination modules")
    c.setFont("Helvetica", 9)
    c.drawString(0.7 * inch, 0.55 * inch, "Confidential Training Operations")
    c.showPage()

    # Slide 2 — grade rule
    c.setFillColor(white)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    place_logo(c, logo, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(0.7 * inch, h - 1.15 * inch, "Grade rule")
    c.setFont("Helvetica", 14)
    lines = [
        "Average is the mean of Paris, Tokyo, Bali, Cancun, and Rome, shown to one decimal place.",
        "Bands use nested thresholds on that average:",
        "    Expert — 90 and above",
        "    Advanced — 80 up to but not including 90",
        "    Proficient — 70 up to but not including 80",
        "    Needs Improvement — below 70",
        "GradeSummary counts each band with COUNTIF against the Scores Band column,",
        "listed in the order Expert, Advanced, Proficient, Needs Improvement.",
    ]
    y = h - 1.7 * inch
    for line in lines:
        c.drawString(0.7 * inch, y, line)
        y -= 0.38 * inch
    footer(c, w, h)
    c.showPage()

    # Slide 3 — chart
    c.setFillColor(white)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    place_logo(c, logo, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(0.7 * inch, h - 1.15 * inch, "Grade distribution")
    c.drawImage(
        str(chart),
        0.7 * inch,
        0.7 * inch,
        width=w - 1.6 * inch,
        height=h - 2.15 * inch,
        preserveAspectRatio=True,
        mask="auto",
    )
    footer(c, w, h)
    c.showPage()

    # Slide 4 — one roster table by band
    c.setFillColor(white)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    place_logo(c, logo, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(0.7 * inch, h - 1.15 * inch, "Roster by band")
    order = {b: i for i, b in enumerate(BANDS)}
    roster = sorted(computed, key=lambda r: (order[r[-1]], r[0]))
    data = [["Employee", "Paris", "Tokyo", "Bali", "Cancun", "Rome", "Average", "Band"]]
    for r in roster:
        data.append(
            [
                r[0],
                str(r[1]),
                str(r[2]),
                str(r[3]),
                str(r[4]),
                str(r[5]),
                f"{r[6]:.1f}",
                r[7],
            ]
        )
    col_w = [1.7 * inch, 0.85 * inch, 0.85 * inch, 0.85 * inch, 0.95 * inch, 0.85 * inch, 0.95 * inch, 1.7 * inch]
    table = RLTable(data, colWidths=col_w, rowHeights=[22] * len(data))
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY_HEX),
                ("TEXTCOLOR", (0, 0), (-1, 0), white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (1, 0), (6, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.4, NAVY_HEX),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [HexColor("#EAF0F6"), white]),
            ]
        )
    )
    tw, th = table.wrap(w, h)
    table.drawOn(c, 0.7 * inch, h - 1.4 * inch - th)
    footer(c, w, h)
    c.showPage()
    c.save()


def write_sources(path: Path) -> None:
    path.write_text(
        """Travel Training / Destination Specialist — asset sources

agency_logo.png
  Derived from Wikimedia Commons File:Globe_icon-white.svg
  https://commons.wikimedia.org/wiki/File:Globe_icon-white.svg
  Author: Sarang (own work, 2012)
  License: Public domain
  Rasterized to PNG and recolored stroke #1B4F72 for the slide master.
  No photograph of a person.

certification_results.xlsx
  Original synthetic scores for ten employees and five destination modules.
  Names are fictional. No PII.

Gold files (not in the source zip)
  certification_results_GOLD.xlsx — Average / Band formulas, GradeSummary COUNTIF, navy column chart.
  grade_distribution.png — rendered GradeSummary counts.
  Destination_Specialist_Results.pdf — four 16:9 pages (title, grade rule, chart, one roster table).
""",
        encoding="utf-8",
    )


def zip_dir(zip_path: Path, files: list[tuple[Path, str]]) -> None:
    with ZipFile(zip_path, "w", ZIP_DEFLATED) as z:
        for src, arc in files:
            z.write(src, arc)


def main() -> None:
    ROOT.mkdir(parents=True, exist_ok=True)
    logo = ROOT / "agency_logo.png"
    src_xlsx = ROOT / "certification_results.xlsx"
    gold_xlsx = ROOT / "certification_results_GOLD.xlsx"
    chart = ROOT / "grade_distribution.png"
    pdf = ROOT / "Destination_Specialist_Results.pdf"
    sources = ROOT / "SOURCES.txt"

    draw_logo(logo)
    build_source_xlsx(src_xlsx)
    build_gold_xlsx(gold_xlsx)
    counts = build_chart_png(chart)
    build_pdf(pdf, logo, chart, counts)
    write_sources(sources)

    zip_dir(
        REPO / "TravelTraining_Assets.zip",
        [
            (src_xlsx, "TravelTraining/certification_results.xlsx"),
            (logo, "TravelTraining/agency_logo.png"),
            (ROOT / "Globe_icon-white.svg", "TravelTraining/Globe_icon-white.svg"),
            (sources, "TravelTraining/SOURCES.txt"),
        ],
    )
    zip_dir(
        REPO / "TravelTraining_Expected.zip",
        [
            (gold_xlsx, "TravelTraining/certification_results_GOLD.xlsx"),
            (chart, "TravelTraining/grade_distribution.png"),
            (pdf, "TravelTraining/Destination_Specialist_Results.pdf"),
        ],
    )
    print("counts", counts)
    print("wrote", ROOT)


if __name__ == "__main__":
    main()

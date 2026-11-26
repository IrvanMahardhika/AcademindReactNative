#!/usr/bin/env python3
"""Build Travel Training source files and gold Excel / PNG / 5-slide PDF."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import cairosvg
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import cm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Table as RLTable
from reportlab.platypus import TableStyle

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
FONT_DIR = Path("/tmp/roboto")
NAVY = "1B4F72"
NAVY_HEX = HexColor("#1B4F72")
CARD = HexColor("#F4F7FA")
BODY = HexColor("#202124")
MUTED = HexColor("#5F6368")
PAGE = landscape((13.333 * inch, 7.5 * inch))

ROWS = [
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
DESTS = ["Paris", "Tokyo", "Bali", "Cancun", "Rome"]
BANDS = ["Expert", "Advanced", "Proficient", "Needs Improvement"]

GLOBE_SVG = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="420" height="420">
  <circle cx="210" cy="210" r="208" fill="#FFFFFF"/>
  <path stroke="#1B4F72" fill="none" stroke-width="20"
    d="M209,15a195,195 0 1,0 2,0zm1,0v390m195-195H15M59,90a260,260 0 0,0 302,0 m0,240 a260,260 0 0,0-302,0M195,20a250,250 0 0,0 0,382 m30,0 a250,250 0 0,0 0-382"/>
</svg>
"""


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


def stats():
    rows = computed_rows()
    counts = {b: 0 for b in BANDS}
    for r in rows:
        counts[r[-1]] += 1
    means = {d: round(sum(r[i + 1] for r in rows) / len(rows), 1) for i, d in enumerate(DESTS)}
    avgs = [r[6] for r in rows]
    cohort = round(sum(avgs) / len(avgs), 1)
    cert = round(100 * sum(1 for a in avgs if a >= 70) / len(avgs), 1)
    sorted_rows = sorted(rows, key=lambda r: -r[6])
    hi = max(means, key=means.get)
    lo_val = min(means.values())
    lo = [d for d, v in means.items() if v == lo_val]
    spread = round(max(means.values()) - lo_val, 1)
    return {
        "rows": rows,
        "sorted": sorted_rows,
        "counts": counts,
        "means": means,
        "cohort": cohort,
        "cert": cert,
        "hi": hi,
        "lo": lo,
        "spread": spread,
    }


def register_fonts() -> None:
    pdfmetrics.registerFont(TTFont("Roboto", str(FONT_DIR / "Roboto-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("Roboto-Bold", str(FONT_DIR / "Roboto-Bold.ttf")))


def draw_logo(path: Path) -> None:
    svg_path = path.with_name("Globe_icon-white.svg")
    svg_path.write_text(GLOBE_SVG, encoding="utf-8")
    cairosvg.svg2png(bytestring=GLOBE_SVG.encode("utf-8"), write_to=str(path), output_width=512)


def style_header(ws, n):
    fill = PatternFill("solid", fgColor=NAVY)
    font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    for i in range(1, n + 1):
        c = ws.cell(1, i)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal="center", wrap_text=True)


def build_source_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Employee", *DESTS])
    for row in ROWS:
        ws.append(list(row))
    style_header(ws, 6)
    ws.column_dimensions["A"].width = 16
    for col in "BCDEF":
        ws.column_dimensions[col].width = 12
    wb.save(path)


def navy_col_chart(title, y_title, cats_ref, data_ref, ymax=None) -> BarChart:
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = "Band" if "Grade" in title else "Destination"
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.legend = None
    chart.y_axis.scaling.min = 0
    if ymax is not None:
        chart.y_axis.scaling.max = ymax
    chart.series[0].graphicalProperties.solidFill = NAVY
    chart.series[0].dLbls = DataLabelList()
    chart.series[0].dLbls.showVal = True
    chart.series[0].dLbls.showCatName = False
    chart.series[0].dLbls.showSerName = False
    return chart


def build_working_xlsx(path: Path) -> None:
    ordered = stats()["sorted"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Scores"
    ws.append(["Employee", *DESTS, "Average", "Band"])
    for i, r in enumerate(ordered, start=2):
        ws.append([r[0], *r[1:6]])
        ws.cell(i, 7).value = f"=AVERAGE([@Paris],[@Tokyo],[@Bali],[@Cancun],[@Rome])"
        ws.cell(i, 7).number_format = "0.0"
        ws.cell(i, 8).value = (
            '=IF([@Average]>=90,"Expert",IF([@Average]>=80,"Advanced",'
            'IF([@Average]>=70,"Proficient","Needs Improvement")))'
        )
    style_header(ws, 8)
    for col, w in zip("ABCDEFGH", [16, 12, 12, 12, 12, 12, 12, 22]):
        ws.column_dimensions[col].width = w
    tab = Table(displayName="Scores", ref="A1:H11")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    ws.conditional_formatting.add(
        "G2:G11",
        DataBarRule(start_type="min", end_type="max", color=NAVY, showValue=None, minLength=None, maxLength=None),
    )

    gs = wb.create_sheet("GradeSummary")
    gs.append(["Band", "Count"])
    for i, band in enumerate(BANDS, start=2):
        gs.cell(i, 1).value = band
        gs.cell(i, 2).value = f"=COUNTIF(Scores[Band],A{i})"
    style_header(gs, 2)
    gs.column_dimensions["A"].width = 22
    gs.column_dimensions["B"].width = 12
    gst = Table(displayName="GradeSummaryTable", ref="A1:B5")
    gst.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    gs.add_table(gst)
    data = Reference(gs, min_col=2, min_row=1, max_row=5)
    cats = Reference(gs, min_col=1, min_row=2, max_row=5)
    chart = navy_col_chart("Grade distribution", "Employees", cats, data)
    chart.x_axis.title = "Band"
    gs.add_chart(chart, "D2")

    ms = wb.create_sheet("ModuleSummary")
    ms.append(["Destination", "Mean"])
    for i, dest in enumerate(DESTS, start=2):
        ms.cell(i, 1).value = dest
        ms.cell(i, 2).value = f"=AVERAGE(Scores[{dest}])"
        ms.cell(i, 2).number_format = "0.0"
    ms["A8"] = "Cohort average"
    ms["B8"] = "=AVERAGE(Scores[Average])"
    ms["B8"].number_format = "0.0"
    ms["A9"] = "Certification rate"
    ms["B9"] = '=COUNTIF(Scores[Average],">=70")/COUNTA(Scores[Employee])'
    ms["B9"].number_format = "0.0%"
    style_header(ms, 2)
    ms.column_dimensions["A"].width = 22
    ms.column_dimensions["B"].width = 14
    data2 = Reference(ms, min_col=2, min_row=1, max_row=6)
    cats2 = Reference(ms, min_col=1, min_row=2, max_row=6)
    chart2 = navy_col_chart("Destination module means", "Mean score", cats2, data2, ymax=100)
    chart2.x_axis.title = "Destination"
    ms.add_chart(chart2, "D2")
    wb.save(path)


def bar_png(path: Path, title: str, xlabel: str, ylabel: str, labels, values, ymax=None) -> None:
    fig, ax = plt.subplots(figsize=(7.2, 4.6), dpi=140)
    bars = ax.bar(labels, values, color="#1B4F72", width=0.62)
    ax.set_title(title, fontsize=14, color="#1B4F72", pad=10, fontweight="bold")
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    top = ymax if ymax is not None else max(values) + 1.4
    ax.set_ylim(0, top)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fmt = "{:.1f}" if any(isinstance(v, float) and v != int(v) for v in values) or ymax == 100 else "{:.0f}"
    for bar, y in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, y + (top * 0.02), fmt.format(y), ha="center", va="bottom", color="#1B4F72", fontsize=10)
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def master_chrome(c, w, h, logo: Path) -> None:
    c.setFillColor(white)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    c.setFillColor(NAVY_HEX)
    c.rect(0, h - 0.12 * cm, w, 0.12 * cm, fill=1, stroke=0)
    size = 2.0 * cm
    c.drawImage(str(logo), w - size - 0.4 * cm, h - size - 0.4 * cm, width=size, height=size, mask="auto", preserveAspectRatio=True)
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.55 * inch, 0.28 * inch, "Confidential Training Operations")
    c.drawRightString(w - 0.55 * inch, 0.28 * inch, str(c.getPageNumber()))


def rounded_card(c, x, y, cw, ch):
    c.setFillColor(CARD)
    c.setStrokeColor(NAVY_HEX)
    c.setLineWidth(1)
    c.roundRect(x, y, cw, ch, 8, fill=1, stroke=1)


def build_pdf(path: Path, logo: Path, grade_png: Path, module_png: Path) -> None:
    s = stats()
    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle("Destination specialist results")
    c.setAuthor("Training Operations")

    # Slide 1
    master_chrome(c, w, h, logo)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 32)
    c.drawString(0.7 * inch, h / 2 + 0.2 * inch, "Destination specialist results")
    c.setFillColor(MUTED)
    c.setFont("Roboto", 18)
    c.drawString(0.7 * inch, h / 2 - 0.25 * inch, "Average scores, grade bands, and destination modules")
    c.showPage()

    # Slide 2
    master_chrome(c, w, h, logo)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Certification grade policy")
    card_w, card_h = 5.6 * inch, 1.15 * inch
    gap = 0.35 * inch
    x0 = 0.7 * inch
    y_card = h - 2.35 * inch
    rounded_card(c, x0, y_card, card_w, card_h)
    rounded_card(c, x0 + card_w + gap, y_card, card_w, card_h)
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(x0 + 0.2 * inch, y_card + 0.75 * inch, "Cohort average")
    c.drawString(x0 + card_w + gap + 0.2 * inch, y_card + 0.75 * inch, "Certification rate")
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(x0 + 0.2 * inch, y_card + 0.28 * inch, f"{s['cohort']:.1f}")
    c.drawString(x0 + card_w + gap + 0.2 * inch, y_card + 0.28 * inch, f"{s['cert']:.1f}%")

    data = [
        ["Band", "Minimum average", "Rule"],
        ["Expert", "90", "Average 90 and above"],
        ["Advanced", "80", "Average 80 up to 90"],
        ["Proficient", "70", "Average 70 up to 80"],
        ["Needs Improvement", "Below 70", "Average below 70"],
    ]
    table = RLTable(data, colWidths=[2.4 * inch, 2.2 * inch, 6.4 * inch], rowHeights=[22] * 5)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY_HEX),
                ("TEXTCOLOR", (0, 0), (-1, 0), white),
                ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Roboto"),
                ("FONTSIZE", (0, 0), (-1, -1), 14),
                ("TEXTCOLOR", (0, 1), (-1, -1), BODY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, NAVY_HEX),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, CARD]),
            ]
        )
    )
    tw, th = table.wrap(w, h)
    table.drawOn(c, 0.7 * inch, y_card - 0.25 * inch - th)
    c.setFillColor(BODY)
    c.setFont("Roboto", 14)
    c.drawString(0.7 * inch, 0.7 * inch, "Band is computed with the nested IF formula on Average.")
    c.showPage()

    # Slide 3
    master_chrome(c, w, h, logo)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Grade distribution")
    img_w, img_h = 7.1 * inch, 4.4 * inch
    c.setStrokeColor(NAVY_HEX)
    c.setLineWidth(1)
    c.rect(0.7 * inch, 1.35 * inch, img_w, img_h, fill=0, stroke=1)
    c.drawImage(str(grade_png), 0.75 * inch, 1.4 * inch, width=img_w - 0.1 * inch, height=img_h - 0.1 * inch, preserveAspectRatio=True, mask="auto")
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.7 * inch, 1.1 * inch, "Employee counts by grade band")
    bx, by, bw, bh = 8.05 * inch, 1.35 * inch, 4.5 * inch, 4.4 * inch
    rounded_card(c, bx, by, bw, bh)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 14)
    c.drawString(bx + 0.2 * inch, by + bh - 0.4 * inch, "Key findings")
    c.setFillColor(BODY)
    c.setFont("Roboto", 14)
    findings = [
        "Largest grade band: Expert and Advanced (3 employees each)",
        "Combined Expert and Advanced: 6 employees",
        "Needs Improvement: 2 employees",
    ]
    fy = by + bh - 0.9 * inch
    for line in findings:
        c.drawString(bx + 0.2 * inch, fy, "")
        textobject = c.beginText(bx + 0.2 * inch, fy)
        textobject.setFont("Roboto", 14)
        textobject.setFillColor(BODY)
        for part in _wrap(line, 32):
            textobject.textLine(part)
        c.drawText(textobject)
        fy -= 1.05 * inch
    c.showPage()

    # Slide 4
    master_chrome(c, w, h, logo)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Destination module means")
    c.setStrokeColor(NAVY_HEX)
    c.rect(0.7 * inch, 1.35 * inch, img_w, img_h, fill=0, stroke=1)
    c.drawImage(str(module_png), 0.75 * inch, 1.4 * inch, width=img_w - 0.1 * inch, height=img_h - 0.1 * inch, preserveAspectRatio=True, mask="auto")
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.7 * inch, 1.1 * inch, "Mean score by destination module")
    rounded_card(c, bx, by, bw, bh)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 14)
    c.drawString(bx + 0.2 * inch, by + bh - 0.4 * inch, "Key findings")
    lo = " and ".join(s["lo"])
    findings4 = [
        f"Highest-scoring destination: {s['hi']} ({s['means'][s['hi']]:.1f})",
        f"Lowest-scoring destination: {lo} ({s['means'][s['lo'][0]]:.1f})",
        f"Point spread across modules: {s['spread']:.1f}",
    ]
    fy = by + bh - 0.95 * inch
    for line in findings4:
        textobject = c.beginText(bx + 0.2 * inch, fy)
        textobject.setFont("Roboto", 14)
        textobject.setFillColor(BODY)
        for part in _wrap(line, 32):
            textobject.textLine(part)
        c.drawText(textobject)
        fy -= 1.05 * inch
    c.showPage()

    # Slide 5
    master_chrome(c, w, h, logo)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Roster by band")
    data = [["Employee", "Average", "Band"]]
    for r in s["sorted"]:
        data.append([r[0], f"{r[6]:.1f}", r[7]])
    table = RLTable(data, colWidths=[4.2 * inch, 2.4 * inch, 4.6 * inch], rowHeights=[22] * 11)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY_HEX),
                ("TEXTCOLOR", (0, 0), (-1, 0), white),
                ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Roboto"),
                ("FONTSIZE", (0, 0), (-1, -1), 14),
                ("TEXTCOLOR", (0, 1), (-1, -1), BODY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, NAVY_HEX),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, CARD]),
            ]
        )
    )
    tw, th = table.wrap(w, h)
    table.drawOn(c, 0.7 * inch, h - 1.2 * inch - th)
    c.showPage()
    c.save()


def _wrap(text: str, width: int) -> list[str]:
    words = text.split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if len(trial) <= width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def write_sources(path: Path) -> None:
    path.write_text(
        """Travel Training / Destination Specialist — asset sources

agency_logo.png
  Derived from Wikimedia Commons File:Globe_icon-white.svg
  https://commons.wikimedia.org/wiki/File:Globe_icon-white.svg
  Author: Sarang (own work, 2012)
  License: Public domain
  Rasterized to PNG, stroke #1B4F72.

certification_results.xlsx
  Original synthetic scores. Fictional names. No PII.

Expected (not in the source zip)
  certification_working.xlsx — Table Scores, Average / Band formulas, Data Bars,
    GradeSummary COUNTIF, ModuleSummary means / cohort / certification rate, two navy charts.
  grade_distribution.png
  module_means.png
  Destination_Specialist_Results.pdf — five 16:9 slides.
""",
        encoding="utf-8",
    )


def zip_dir(zip_path: Path, files: list[tuple[Path, str]]) -> None:
    with ZipFile(zip_path, "w", ZIP_DEFLATED) as z:
        for src, arc in files:
            z.write(src, arc)


def main() -> None:
    register_fonts()
    ROOT.mkdir(parents=True, exist_ok=True)
    logo = ROOT / "agency_logo.png"
    src_xlsx = ROOT / "certification_results.xlsx"
    work_xlsx = ROOT / "certification_working.xlsx"
    grade_png = ROOT / "grade_distribution.png"
    module_png = ROOT / "module_means.png"
    pdf = ROOT / "Destination_Specialist_Results.pdf"
    sources = ROOT / "SOURCES.txt"

    draw_logo(logo)
    build_source_xlsx(src_xlsx)
    build_working_xlsx(work_xlsx)
    s = stats()
    bar_png(grade_png, "Grade distribution", "Band", "Employees", BANDS, [s["counts"][b] for b in BANDS])
    bar_png(
        module_png,
        "Destination module means",
        "Destination",
        "Mean score",
        DESTS,
        [s["means"][d] for d in DESTS],
        ymax=100,
    )
    build_pdf(pdf, logo, grade_png, module_png)
    write_sources(sources)

    old_gold = ROOT / "certification_results_GOLD.xlsx"
    if old_gold.exists():
        old_gold.unlink()

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
            (work_xlsx, "TravelTraining/certification_working.xlsx"),
            (grade_png, "TravelTraining/grade_distribution.png"),
            (module_png, "TravelTraining/module_means.png"),
            (pdf, "TravelTraining/Destination_Specialist_Results.pdf"),
        ],
    )
    print(s["counts"], s["means"], s["cohort"], s["cert"], s["hi"], s["lo"], s["spread"])


if __name__ == "__main__":
    main()

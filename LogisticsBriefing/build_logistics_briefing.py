#!/usr/bin/env python3
"""Build Q3 logistics source pack and gold Excel / PNG / notes PDF."""

from __future__ import annotations

import shutil
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from PIL import Image, ImageEnhance
from reportlab.lib.colors import HexColor, Color, white
from reportlab.lib.units import cm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Table as RLTable
from reportlab.platypus import TableStyle

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
FONT_DIR = Path("/tmp/roboto")
BLUE = HexColor("#2E86C1")
RED = HexColor("#C0392B")
CARD = HexColor("#F4F7FA")
BODY = HexColor("#202124")
MUTED = HexColor("#5F6368")
SLIDE_W = 13.333 * inch
SLIDE_H = 7.5 * inch
NOTES_H = 1.85 * inch
PAGE = (SLIDE_W, SLIDE_H + NOTES_H)

HUBS = [
    ("North Crossdock", "North", 0.942, 18.40),
    ("Harbor Gate", "Coast", 0.881, 21.10),
    ("Prairie Hub", "Central", 0.910, 16.75),
    ("Ridge Terminal", "West", 0.864, 22.05),
    ("Canal Yard", "East", 0.925, 17.20),
    ("Mesa Sort", "West", 0.798, 24.60),
    ("Lakeside DC", "North", 0.900, 19.15),
    ("Metro East", "East", 0.872, 20.40),
]

POLICY = (
    "A hub is treated as below the bar when its on-time delivery rate is below 90 percent. "
    "Rates at exactly 90 percent remain on the bar. Cost per stop is reported for context and does not change the flag."
)

PARAS = {
    "North Crossdock": "North Crossdock posted a 94.2 percent on-time rate with a cost per stop of $18.40. Inbound appointments held, and the dock turned trailers without a same-week backlog.",
    "Harbor Gate": "Harbor Gate posted an 88.1 percent on-time rate with a cost per stop of $21.10. Late ocean dray and a short night crew pushed appointments past the delivery window on three weekdays.",
    "Prairie Hub": "Prairie Hub posted a 91.0 percent on-time rate with a cost per stop of $16.75. Mix stayed inside the plan, and exception tickets closed inside one shift.",
    "Ridge Terminal": "Ridge Terminal posted an 86.4 percent on-time rate with a cost per stop of $22.05. A closed weigh-station lane and two broken dock plates stacked outbound freight into the next morning.",
    "Canal Yard": "Canal Yard posted a 92.5 percent on-time rate with a cost per stop of $17.20. Live unloads stayed inside the booked hour, and the yard hostler pool covered peak gates.",
    "Mesa Sort": "Mesa Sort posted a 79.8 percent on-time rate with a cost per stop of $24.60. Sort-belt stops and a late inbound wave left parcels short of the last mile cutoff.",
    "Lakeside DC": "Lakeside DC posted a 90.0 percent on-time rate with a cost per stop of $19.15. The site sits on the bar. Weekend volume was covered with overtime rather than missed windows.",
    "Metro East": "Metro East posted an 87.2 percent on-time rate with a cost per stop of $20.40. City congestion and a closed left-turn onto the ring road added 20 minutes to the last ten stops on four routes.",
}

FLAGGED = [h for h in HUBS if h[2] < 0.9]


def register_fonts() -> None:
    pdfmetrics.registerFont(TTFont("Roboto", str(FONT_DIR / "Roboto-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("Roboto-Bold", str(FONT_DIR / "Roboto-Bold.ttf")))


def build_docx(path: Path) -> None:
    doc = Document()
    t = doc.add_paragraph("Q3 logistics report")
    t.runs[0].bold = True
    t.runs[0].font.size = Pt(20)
    doc.add_heading("On-time policy", level=1)
    doc.add_paragraph(POLICY)
    doc.add_heading("Hub table", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for i, h in enumerate(["Hub", "Region", "OnTime", "CostPerStop"]):
        table.rows[0].cells[i].text = h
    for name, region, ot, cost in HUBS:
        row = table.add_row().cells
        row[0].text = name
        row[1].text = region
        row[2].text = f"{ot * 100:.1f}%"
        row[3].text = f"{cost:.2f}"
    for name, *_rest in HUBS:
        doc.add_heading(name, level=1)
        doc.add_paragraph(PARAS[name])
    doc.save(path)


def build_xlsx(path: Path) -> None:
    fill = PatternFill("solid", fgColor="2E86C1")
    hf = Font(name="Calibri", bold=True, color="FFFFFF")
    wb = Workbook()
    ws = wb.active
    ws.title = "Hubs"
    ws.append(["Hub", "Region", "OnTime", "CostPerStop", "Flag", "Threshold"])
    for i, (name, region, ot, cost) in enumerate(HUBS, start=2):
        ws.append([name, region, ot, cost])
        ws.cell(i, 3).number_format = "0.0%"
        ws.cell(i, 4).number_format = "0.00"
        ws.cell(i, 5).value = '=IF([@OnTime]<0.9,"Flagged","Clear")'
        ws.cell(i, 6).value = 0.9
        ws.cell(i, 6).number_format = "0%"
    for c in range(1, 7):
        ws.cell(1, c).fill = fill
        ws.cell(1, c).font = hf
    tab = Table(displayName="Hubs", ref="A1:F9")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    for col, w in zip("ABCDEF", [18, 12, 12, 14, 12, 12]):
        ws.column_dimensions[col].width = w

    fl = wb.create_sheet("Flags")
    fl["A1"] = "Metric"
    fl["B1"] = "Value"
    fl["A1"].fill = fill
    fl["B1"].fill = fill
    fl["A1"].font = hf
    fl["B1"].font = hf
    fl["A2"] = "FlaggedCount"
    fl["B2"] = '=COUNTIF(Hubs[Flag],"Flagged")'
    fl["A3"] = "MeanOnTime"
    fl["B3"] = "=AVERAGE(Hubs[OnTime])"
    fl["B3"].number_format = "0.0%"
    fl["A4"] = "ResponseCount"
    fl["B4"] = "=COUNTA(Hubs[Hub])"

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "On-time rate by hub"
    chart.y_axis.title = "Percent"
    chart.x_axis.title = "Hub"
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1
    data = Reference(ws, min_col=3, min_row=1, max_col=3, max_row=9)
    thresh = Reference(ws, min_col=6, min_row=1, max_row=9)
    cats = Reference(ws, min_col=1, min_row=2, max_row=9)
    chart.add_data(data, titles_from_data=True)
    chart.add_data(thresh, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = "2E86C1"
    chart.series[1].graphicalProperties.solidFill = "C0392B"
    chart.series[0].dLbls = DataLabelList()
    chart.series[0].dLbls.showVal = True
    ws.add_chart(chart, "H2")
    wb.save(path)


def build_chart_png(path: Path) -> None:
    names = [h[0] for h in HUBS]
    ots = [h[2] * 100 for h in HUBS]
    fig, ax = plt.subplots(figsize=(9, 4.8), dpi=140)
    x = range(len(names))
    ax.bar(list(x), ots, color="#2E86C1", width=0.55, label="OnTime")
    ax.plot(list(x), [90] * len(x), color="#C0392B", linewidth=2, label="Threshold")
    ax.set_xticks(list(x), names, rotation=25, ha="right")
    ax.set_ylim(0, 100)
    ax.set_title("On-time rate by hub", color="#2E86C1", fontweight="bold")
    ax.set_xlabel("Hub")
    ax.set_ylabel("Percent")
    ax.legend(loc="lower center", ncol=2, bbox_to_anchor=(0.5, -0.32))
    for i, v in enumerate(ots):
        ax.text(i, v + 1.2, f"{v:.1f}", ha="center", fontsize=8, color="#2E86C1")
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def dim_photo(src: Path, dest: Path) -> None:
    im = Image.open(src).convert("RGB")
    im = ImageEnhance.Brightness(im).enhance(0.45)
    im.save(dest, quality=85)


def notes_band(c: canvas.Canvas, text: str, page: int, total: int = 5) -> None:
    w, h = PAGE
    c.setFillColor(HexColor("#EEF2F6"))
    c.rect(0, 0, w, NOTES_H, fill=1, stroke=0)
    c.setFillColor(BLUE)
    c.rect(0, NOTES_H, w, 1.5, fill=1, stroke=0)
    c.setFillColor(MUTED)
    c.setFont("Roboto-Bold", 9)
    c.drawString(0.5 * inch, NOTES_H - 0.28 * inch, "Speaker notes")
    c.setFont("Roboto", 9)
    t = c.beginText(0.5 * inch, NOTES_H - 0.48 * inch)
    t.setFillColor(BODY)
    t.setFont("Roboto", 9)
    t.setLeading(12)
    for line in wrap(text, 145):
        t.textLine(line)
    c.drawText(t)
    c.setFillColor(MUTED)
    c.setFont("Roboto", 9)
    c.drawRightString(w - 0.45 * inch, 0.18 * inch, f"{page} / {total}")


def chrome(c: canvas.Canvas) -> None:
    w = SLIDE_W
    h = NOTES_H
    top = PAGE[1]
    c.setFillColor(white)
    c.rect(0, h, w, SLIDE_H, fill=1, stroke=0)
    c.setFillColor(BLUE)
    c.rect(0, top - 0.12 * cm, w, 0.12 * cm, fill=1, stroke=0)
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.55 * inch, h + 0.22 * inch, "Confidential Logistics Operations")
    c.drawRightString(w - 0.55 * inch, h + 0.22 * inch, str(c.getPageNumber()))


def rounded(c, x, y, cw, ch):
    c.setFillColor(CARD)
    c.setStrokeColor(BLUE)
    c.setLineWidth(1)
    c.roundRect(x, y, cw, ch, 8, fill=1, stroke=1)


def wrap(text: str, width: int) -> list[str]:
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


def build_pdf(path: Path, photo: Path, chart: Path) -> None:
    w, page_h = PAGE
    base = NOTES_H
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle("Q3 logistics briefing")

    # Slide 1
    c.drawImage(str(photo), 0, base, width=SLIDE_W, height=SLIDE_H, preserveAspectRatio=False, mask="auto")
    c.setFillColor(Color(1, 1, 1, alpha=0.18))
    c.rect(0, base, SLIDE_W, SLIDE_H, fill=1, stroke=0)
    c.setFillColor(BLUE)
    c.rect(0, page_h - 0.12 * cm, w, 0.12 * cm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, base + 4.1 * inch, "Q3 logistics briefing")
    c.setFillColor(HexColor("#E8EEF4"))
    c.setFont("Roboto", 18)
    c.drawString(0.7 * inch, base + 3.55 * inch, "Hubs below the 90 percent on-time bar")
    c.setFillColor(white)
    c.setFont("Roboto", 12)
    c.drawString(0.7 * inch, base + 0.35 * inch, "Confidential Logistics Operations")
    notes_band(c, "Title slide. Warehouse exterior cropped 16:9 behind the title.", 1)
    c.showPage()

    # Slide 2
    chrome(c)
    y0 = base
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, y0 + SLIDE_H - 0.95 * inch, "On-time rule")
    card_w, card_h = 5.6 * inch, 1.15 * inch
    gap = 0.35 * inch
    x0 = 0.7 * inch
    yc = y0 + SLIDE_H - 2.4 * inch
    rounded(c, x0, yc, card_w, card_h)
    rounded(c, x0 + card_w + gap, yc, card_w, card_h)
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(x0 + 0.2 * inch, yc + 0.75 * inch, "FlaggedCount")
    c.drawString(x0 + card_w + gap + 0.2 * inch, yc + 0.75 * inch, "MeanOnTime")
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(x0 + 0.2 * inch, yc + 0.28 * inch, "4")
    c.drawString(x0 + card_w + gap + 0.2 * inch, yc + 0.28 * inch, "88.7%")
    data = [
        ["Measure", "Threshold", "Flag rule"],
        ["OnTime", "90%", "Flagged when OnTime is below 90 percent"],
    ]
    table = RLTable(data, colWidths=[2.4 * inch, 2.2 * inch, 6.4 * inch], rowHeights=[24, 28])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), BLUE),
                ("TEXTCOLOR", (0, 0), (-1, 0), white),
                ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Roboto"),
                ("FONTSIZE", (0, 0), (-1, -1), 14),
                ("TEXTCOLOR", (0, 1), (-1, -1), BODY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, BLUE),
            ]
        )
    )
    tw, th = table.wrap(w, SLIDE_H)
    table.drawOn(c, 0.7 * inch, yc - 0.35 * inch - th)
    notes_band(c, POLICY, 2)
    c.showPage()

    # Slide 3
    chrome(c)
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, y0 + SLIDE_H - 0.95 * inch, "On-time rate by hub")
    img_w, img_h = 7.15 * inch, 4.35 * inch
    c.setStrokeColor(BLUE)
    c.setLineWidth(1)
    c.rect(0.7 * inch, y0 + 1.35 * inch, img_w, img_h, fill=0, stroke=1)
    c.drawImage(str(chart), 0.75 * inch, y0 + 1.4 * inch, width=img_w - 0.1 * inch, height=img_h - 0.1 * inch, preserveAspectRatio=True, mask="auto")
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.7 * inch, y0 + 1.12 * inch, "On-time percent versus 90 percent")
    bx, by, bw, bh = 8.1 * inch, y0 + 1.35 * inch, 4.45 * inch, 4.35 * inch
    rounded(c, bx, by, bw, bh)
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 14)
    c.drawString(bx + 0.2 * inch, by + bh - 0.4 * inch, "Key findings")
    c.setFillColor(BODY)
    c.setFont("Roboto", 14)
    fy = by + bh - 1.0 * inch
    for line in [
        "Lowest OnTime: Mesa Sort (79.8%)",
        "Flagged hubs: 4",
        "Mean OnTime: 88.7%",
    ]:
        t = c.beginText(bx + 0.2 * inch, fy)
        t.setFont("Roboto", 14)
        t.setFillColor(BODY)
        t.setLeading(18)
        for part in wrap(line, 34):
            t.textLine(part)
        c.drawText(t)
        fy -= 0.95 * inch
    notes_band(c, "Chart of OnTime versus the 90 percent Threshold series for all eight hubs.", 3)
    c.showPage()

    # Slide 4
    chrome(c)
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, y0 + SLIDE_H - 0.95 * inch, "Flagged hubs")
    rows = [["Hub", "Region", "OnTime"]]
    for name, region, ot, _cost in FLAGGED:
        rows.append([name, region, f"{ot * 100:.1f}%"])
    table = RLTable(rows, colWidths=[4.4 * inch, 3.2 * inch, 3.4 * inch], rowHeights=[24] * 5)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), BLUE),
                ("TEXTCOLOR", (0, 0), (-1, 0), white),
                ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Roboto"),
                ("FONTSIZE", (0, 0), (-1, -1), 14),
                ("TEXTCOLOR", (0, 1), (-1, -1), BODY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (2, 0), (2, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, BLUE),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, CARD]),
            ]
        )
    )
    tw, th = table.wrap(w, SLIDE_H)
    table.drawOn(c, 0.7 * inch, y0 + SLIDE_H - 1.25 * inch - th)
    flagged_notes = " ".join(PARAS[h[0]] for h in FLAGGED)
    notes_band(c, flagged_notes, 4)
    c.showPage()

    # Slide 5
    chrome(c)
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, y0 + SLIDE_H - 0.95 * inch, "Briefing close")
    c.setFillColor(BODY)
    c.setFont("Roboto", 16)
    bullets = [
        "Keep the 90 percent on-time bar.",
        "Assign recovery to Flagged hubs only.",
        "Rerun OnTime after the next cycle.",
    ]
    byb = y0 + SLIDE_H - 1.7 * inch
    for b in bullets:
        c.circle(0.85 * inch, byb + 4, 3, fill=1, stroke=0)
        c.drawString(1.05 * inch, byb, b)
        byb -= 0.45 * inch
    notes_band(c, "Close: keep the bar, recover Flagged hubs only, rerun OnTime next cycle.", 5)
    c.showPage()
    c.save()


def write_sources(path: Path) -> None:
    path.write_text(
        """warehouse_exterior.jpg
  Wikimedia Commons File:Brown Warehouse Ft W IN.jpg
  https://commons.wikimedia.org/wiki/File:Brown_Warehouse_Ft_W_IN.jpg
  Author: Smallbones
  License: CC0

Logistics_Report_Q3.docx
  Original synthetic hub table and Heading 1 sections. Fictional hub names. No PII.

Expected
  Logistics_Report_Q3_working.docx — saved working copy of the report
  hub_rates.xlsx — Table Hubs, Flag, Threshold 90%, Flags KPIs, navy/red chart
  ontime_by_hub.png
  Q3_Logistics_Briefing.pdf — five slides with speaker notes under each slide
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
    photo = ROOT / "warehouse_exterior.jpg"
    dim = ROOT / "warehouse_exterior_dim.jpg"
    docx = ROOT / "Logistics_Report_Q3.docx"
    working = ROOT / "Logistics_Report_Q3_working.docx"
    xlsx = ROOT / "hub_rates.xlsx"
    png = ROOT / "ontime_by_hub.png"
    pdf = ROOT / "Q3_Logistics_Briefing.pdf"
    sources = ROOT / "SOURCES.txt"

    build_docx(docx)
    shutil.copyfile(docx, working)
    build_xlsx(xlsx)
    build_chart_png(png)
    dim_photo(photo, dim)
    build_pdf(pdf, dim, png)
    write_sources(sources)

    zip_dir(
        REPO / "LogisticsBriefing_Assets.zip",
        [
            (docx, "LogisticsBriefing/Logistics_Report_Q3.docx"),
            (photo, "LogisticsBriefing/warehouse_exterior.jpg"),
            (sources, "LogisticsBriefing/SOURCES.txt"),
        ],
    )
    zip_dir(
        REPO / "LogisticsBriefing_Expected.zip",
        [
            (working, "LogisticsBriefing/Logistics_Report_Q3_working.docx"),
            (xlsx, "LogisticsBriefing/hub_rates.xlsx"),
            (png, "LogisticsBriefing/ontime_by_hub.png"),
            (pdf, "LogisticsBriefing/Q3_Logistics_Briefing.pdf"),
        ],
    )
    print("wrote", ROOT)


if __name__ == "__main__":
    main()

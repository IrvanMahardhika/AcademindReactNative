#!/usr/bin/env python3
from __future__ import annotations

import shutil
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill
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
FONTS = Path("/tmp/bbfonts")
NAVY = HexColor("#2C3E50")
ACCENT = HexColor("#E74C3C")
MUTED = HexColor("#7F8C8D")
CARD = HexColor("#F7F9FA")
PAGE = landscape((13.333 * inch, 7.5 * inch))


def fonts():
    pdfmetrics.registerFont(TTFont("Montserrat", str(FONTS / "Montserrat-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("Montserrat-Bold", str(FONTS / "Montserrat-Bold.ttf")))
    pdfmetrics.registerFont(TTFont("OpenSans", str(FONTS / "OpenSans-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("OpenSans-Bold", str(FONTS / "OpenSans-Bold.ttf")))


def build_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "KPIs"
    ws.append(["Metric", "Value"])
    rows = [
        ("Subscribers", 400),
        ("MRR", 18000),
        ("Retention", 0.92),
        ("RegionalGrowth", 0.12),
        ("CafePartners", 15),
        ("RaiseAsk", 75000),
    ]
    for r in rows:
        ws.append(list(r))
    fill = PatternFill("solid", fgColor="2C3E50")
    hf = Font(name="Calibri", bold=True, color="FFFFFF")
    ws["A1"].fill = fill
    ws["B1"].fill = fill
    ws["A1"].font = hf
    ws["B1"].font = hf
    ws["B2"].number_format = "0"
    ws["B3"].number_format = '"$"#,##0'
    ws["B4"].number_format = "0%"
    ws["B5"].number_format = "0%"
    ws["B6"].number_format = "0"
    ws["B7"].number_format = '"$"#,##0'
    ws.add_table(Table(displayName="KPIs", ref="A1:B7", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))
    chart = BarChart()
    chart.type = "col"
    chart.title = "Specialty coffee growth"
    chart.y_axis.title = "Percent"
    chart.x_axis.title = "Metric"
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 0.2
    chart.legend = None
    # chart only RegionalGrowth: put a small range E1:F2
    ws["E1"] = "Metric"
    ws["F1"] = "Percent"
    ws["E2"] = "RegionalGrowth"
    ws["F2"] = "=B5"
    ws["F2"].number_format = "0%"
    data = Reference(ws, min_col=6, min_row=1, max_row=2)
    cats = Reference(ws, min_col=5, min_row=2, max_row=2)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = "E74C3C"
    chart.series[0].dLbls = DataLabelList()
    chart.series[0].dLbls.showVal = True
    ws.add_chart(chart, "D9")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 14
    wb.save(path)


def build_png(path: Path) -> None:
    fig, ax = plt.subplots(figsize=(6.4, 4.2), dpi=140)
    ax.bar(["RegionalGrowth"], [12], color="#E74C3C", width=0.45)
    ax.set_title("Specialty coffee growth", color="#2C3E50", fontweight="bold")
    ax.set_xlabel("Metric")
    ax.set_ylabel("Percent")
    ax.set_ylim(0, 20)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.text(0, 12.4, "12%", ha="center", color="#E74C3C")
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def chrome(c, w, h, logo: Path):
    c.setFillColor(white)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.rect(0, h - 0.12 * cm, w, 0.12 * cm, fill=1, stroke=0)
    size = 2.0 * cm
    c.drawImage(str(logo), 0.4 * cm, h - size - 0.4 * cm, width=size, height=size, mask="auto", preserveAspectRatio=True)
    c.setFillColor(MUTED)
    c.setFont("OpenSans", 12)
    c.drawString(0.7 * inch, 0.28 * inch, "Bean & Brew Roasters")
    c.drawRightString(w - 0.55 * inch, 0.28 * inch, str(c.getPageNumber()))


def card(c, x, y, cw, ch, border):
    c.setFillColor(CARD)
    c.setStrokeColor(border)
    c.setLineWidth(1)
    c.roundRect(x, y, cw, ch, 8, fill=1, stroke=1)


def build_pdf(path: Path, logo: Path, chart: Path) -> None:
    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle("Bean & Brew Pitch Deck")
    chrome(c, w, h, logo)
    c.setFillColor(NAVY)
    c.setFont("Montserrat-Bold", 32)
    c.drawString(0.7 * inch, h / 2 + 0.15 * inch, "Bean & Brew Roasters")
    c.setFillColor(MUTED)
    c.setFont("OpenSans", 18)
    c.drawString(0.7 * inch, h / 2 - 0.3 * inch, "Small-batch coffee, roasted fresh weekly")
    c.showPage()

    chrome(c, w, h, logo)
    c.setFillColor(NAVY)
    c.setFont("Montserrat-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "The roast gap")
    cw, ch = 5.6 * inch, 3.6 * inch
    card(c, 0.7 * inch, 1.5 * inch, cw, ch, NAVY)
    card(c, 0.7 * inch + cw + 0.35 * inch, 1.5 * inch, cw, ch, NAVY)
    c.setFillColor(ACCENT)
    c.setFont("Montserrat-Bold", 18)
    c.drawString(0.95 * inch, 4.7 * inch, "Problem")
    c.drawString(0.95 * inch + cw + 0.35 * inch, 4.7 * inch, "Solution")
    c.setFillColor(NAVY)
    c.setFont("OpenSans", 16)
    t = c.beginText(0.95 * inch, 4.25 * inch)
    t.setLeading(22)
    t.textLines("Most coffee sits in warehouses\nfor months before reaching cups")
    c.drawText(t)
    t = c.beginText(0.95 * inch + cw + 0.35 * inch, 4.25 * inch)
    t.setLeading(22)
    t.textLines("We roast to order and deliver\nwithin 48 hours")
    c.drawText(t)
    c.showPage()

    chrome(c, w, h, logo)
    c.setFillColor(NAVY)
    c.setFont("Montserrat-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Market and model")
    c.setStrokeColor(NAVY)
    c.rect(0.7 * inch, 1.55 * inch, 6.4 * inch, 4.2 * inch, fill=0, stroke=1)
    c.drawImage(str(chart), 0.8 * inch, 1.7 * inch, width=6.2 * inch, height=3.8 * inch, preserveAspectRatio=True, mask="auto")
    c.setFillColor(MUTED)
    c.setFont("OpenSans", 12)
    c.drawString(0.7 * inch, 1.28 * inch, "Specialty coffee market growing 12 percent annually in our region")
    data = [
        ["Channel", "Detail"],
        ["Direct-to-consumer", "Subscriptions"],
        ["Wholesale", "15 local cafes"],
    ]
    table = RLTable(data, colWidths=[2.4 * inch, 2.6 * inch], rowHeights=[24, 28, 28])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), white),
                ("FONTNAME", (0, 0), (-1, 0), "Montserrat-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "OpenSans"),
                ("FONTSIZE", (0, 0), (-1, -1), 14),
                ("TEXTCOLOR", (0, 1), (-1, -1), NAVY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, NAVY),
            ]
        )
    )
    tw, th = table.wrap(w, h)
    table.drawOn(c, 7.5 * inch, 3.2 * inch)
    c.showPage()

    chrome(c, w, h, logo)
    c.setFillColor(NAVY)
    c.setFont("Montserrat-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Traction")
    cards = [
        (0.7 * inch, "Subscribers", "400"),
        (4.85 * inch, "MRR", "$18K"),
        (9.0 * inch, "Retention", "92%"),
    ]
    for x, lab, val in cards:
        card(c, x, 2.6 * inch, 3.7 * inch, 2.2 * inch, ACCENT)
        c.setFillColor(MUTED)
        c.setFont("OpenSans", 12)
        c.drawString(x + 0.25 * inch, 4.35 * inch, lab)
        c.setFillColor(NAVY)
        c.setFont("Montserrat-Bold", 28)
        c.drawString(x + 0.25 * inch, 3.55 * inch, val)
    c.setFillColor(MUTED)
    c.setFont("OpenSans", 14)
    c.drawString(0.7 * inch, 2.15 * inch, "92 percent retention after 6 months")
    c.showPage()

    chrome(c, w, h, logo)
    c.setFillColor(NAVY)
    c.setFont("Montserrat-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "The team")
    data = [
        ["Role", "Profile"],
        ["Founder", "Ex-barista champion"],
        ["Head Roaster", "SCA certified"],
        ["Ops Lead", "Operations"],
    ]
    table = RLTable(data, colWidths=[3.4 * inch, 7.4 * inch], rowHeights=[24, 26, 26, 26])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), white),
                ("FONTNAME", (0, 0), (-1, 0), "Montserrat-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "OpenSans"),
                ("FONTSIZE", (0, 0), (-1, -1), 14),
                ("TEXTCOLOR", (0, 1), (-1, -1), NAVY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, NAVY),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, CARD]),
            ]
        )
    )
    tw, th = table.wrap(w, h)
    table.drawOn(c, 0.7 * inch, h - 1.2 * inch - th)
    c.setFillColor(ACCENT)
    c.setFont("Montserrat-Bold", 18)
    c.drawString(0.7 * inch, 2.4 * inch, "Seeking $75K for second roaster and delivery van")
    c.setFillColor(NAVY)
    c.setFont("OpenSans", 16)
    c.drawString(0.7 * inch, 1.85 * inch, "Join us in bringing fresh coffee to every kitchen")
    c.showPage()
    c.save()


def main():
    fonts()
    ROOT.mkdir(exist_ok=True)
    logo = ROOT / "coffee_bean_logo.png"
    shutil.copyfile("/tmp/coffee_bean_logo.png", logo)
    shutil.copyfile("/tmp/bean.svg", ROOT / "Coffee_bean_Pinhead_icon.svg")
    xlsx = ROOT / "traction_kpis.xlsx"
    png = ROOT / "market_growth.png"
    pdf = ROOT / "Bean_Brew_Pitch_Deck.pdf"
    sources = ROOT / "SOURCES.txt"
    build_xlsx(xlsx)
    build_png(png)
    build_pdf(pdf, logo, png)
    sources.write_text(
        """coffee_bean_logo.png
  Wikimedia Commons File:Coffee bean Pinhead icon.svg
  https://commons.wikimedia.org/wiki/File:Coffee_bean_Pinhead_icon.svg
  License: CC0

traction_kpis.xlsx — synthetic Bean & Brew metrics. No PII.
""",
        encoding="utf-8",
    )
    with ZipFile(REPO / "BeanBrew_Assets.zip", "w", ZIP_DEFLATED) as z:
        z.write(logo, "BeanBrew/coffee_bean_logo.png")
        z.write(xlsx, "BeanBrew/traction_kpis.xlsx")
        z.write(sources, "BeanBrew/SOURCES.txt")
        z.write(ROOT / "Coffee_bean_Pinhead_icon.svg", "BeanBrew/Coffee_bean_Pinhead_icon.svg")
    with ZipFile(REPO / "BeanBrew_Expected.zip", "w", ZIP_DEFLATED) as z:
        z.write(png, "BeanBrew/market_growth.png")
        z.write(pdf, "BeanBrew/Bean_Brew_Pitch_Deck.pdf")
    print("ok")


if __name__ == "__main__":
    main()

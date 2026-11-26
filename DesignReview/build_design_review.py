#!/usr/bin/env python3
from __future__ import annotations

from datetime import date
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
FONTS = Path("/tmp/roboto")
BLUE = HexColor("#1A73E8")
HEAD = HexColor("#E8F0FE")
BODY = HexColor("#202124")
MUTED = HexColor("#5F6368")
PAGE = landscape((13.333 * inch, 7.5 * inch))
CUTOFF = date(2024, 8, 1)

ATTENDEES = [
    ("Kim Hale", "Product design"),
    ("Jordan Voss", "API"),
    ("Riley Chen", "Auth and ops"),
    ("Sam Ortiz", "Facilitator"),
]
AGENDA = [
    ("Visual system", "Kim Hale"),
    ("Checkout API", "Jordan Voss"),
    ("Auth hardening", "Riley Chen"),
]
DECISIONS = [
    ("Contrast", "Ship the navy-on-white header on the storefront this cycle."),
    ("Checkout", "Keep a single cart service; no split for guest checkout."),
    ("Auth", "Turn on token redaction in application logs before the next release."),
]
# ID, Description, Owner, Priority, Due, Status, Subsystem
ACTIONS = [
    ("A-01", "Fix storefront nav contrast", "Kim Hale", "High", date(2024, 7, 22), "Open", "UI"),
    ("A-02", "Spec checkout API error codes", "Jordan Voss", "High", date(2024, 8, 5), "Open", "API"),
    ("A-03", "Audit refresh-token lifetime", "Riley Chen", "Medium", date(2024, 7, 28), "Open", "Auth"),
    ("A-04", "Rewrite footer legal copy", "Kim Hale", "Low", date(2024, 7, 20), "Open", "UI"),
    ("A-05", "Ship dark-mode tokens", "Kim Hale", "High", date(2024, 7, 18), "Done", "UI"),
    ("A-06", "Load-test cart at 2x peak", "Jordan Voss", "Medium", date(2024, 7, 25), "Open", "API"),
    ("A-07", "Confirm vendor SLA hours", "Riley Chen", "Medium", date(2024, 8, 12), "Open", "Ops"),
    ("A-08", "Redact tokens in app logs", "Riley Chen", "High", date(2024, 7, 30), "Open", "Auth"),
]


def flagged(row):
    return row[5] == "Open" and row[3] != "Low"


def next_cycle(row):
    return flagged(row) and row[4] < CUTOFF


def fonts():
    pdfmetrics.registerFont(TTFont("Roboto", str(FONTS / "Roboto-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("Roboto-Bold", str(FONTS / "Roboto-Bold.ttf")))


def build_docx(path: Path) -> None:
    doc = Document()
    t = doc.add_paragraph("Q3 Design Review minutes")
    t.runs[0].bold = True
    t.runs[0].font.size = Pt(20)
    doc.add_heading("Attendees", 1)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Name"
    table.rows[0].cells[1].text = "Role"
    for name, role in ATTENDEES:
        r = table.add_row().cells
        r[0].text, r[1].text = name, role
    doc.add_heading("Agenda", 1)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Topic"
    table.rows[0].cells[1].text = "Owner"
    for topic, owner in AGENDA:
        r = table.add_row().cells
        r[0].text, r[1].text = topic, owner
    doc.add_heading("Decisions", 1)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Decision"
    table.rows[0].cells[1].text = "Detail"
    for d, detail in DECISIONS:
        r = table.add_row().cells
        r[0].text, r[1].text = d, detail
    doc.add_heading("Action items", 1)
    table = doc.add_table(rows=1, cols=7)
    table.style = "Table Grid"
    for i, h in enumerate(["ID", "Description", "Owner", "Priority", "Due Date", "Status", "Subsystem"]):
        table.rows[0].cells[i].text = h
    for row in ACTIONS:
        cells = table.add_row().cells
        cells[0].text = row[0]
        cells[1].text = row[1]
        cells[2].text = row[2]
        cells[3].text = row[3]
        cells[4].text = row[4].strftime("%Y-%m-%d")
        cells[5].text = row[5]
        cells[6].text = row[6]
    doc.add_heading("Cutoff", 1)
    doc.add_paragraph("Next-cycle cutoff date: 2024-08-01. Flagged items due before this date are Due; other Flagged items are Hold.")
    doc.save(path)


def header_row(ws, n):
    fill = PatternFill("solid", fgColor="1A73E8")
    font = Font(name="Calibri", bold=True, color="FFFFFF")
    for i in range(1, n + 1):
        ws.cell(1, i).fill = fill
        ws.cell(1, i).font = font


def build_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Actions"
    ws.append(["ID", "Description", "Owner", "Priority", "Due Date", "Status", "Subsystem", "Flag", "NextCycle"])
    for i, row in enumerate(ACTIONS, start=2):
        ws.append([row[0], row[1], row[2], row[3], row[4], row[5], row[6]])
        ws.cell(i, 5).number_format = "YYYY-MM-DD"
        ws.cell(i, 8).value = '=IF(AND([@Status]="Open",[@Priority]<>"Low"),"Flagged","Clear")'
        ws.cell(i, 9).value = '=IF(AND([@Flag]="Flagged",[@Due Date]<Control!$B$1),"Due","Hold")'
    header_row(ws, 9)
    ws.add_table(Table(displayName="Actions", ref="A1:I9", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))

    ctrl = wb.create_sheet("Control")
    ctrl["A1"] = "Cutoff"
    ctrl["B1"] = CUTOFF
    ctrl["B1"].number_format = "YYYY-MM-DD"

    flagged_rows = sorted([r for r in ACTIONS if flagged(r)], key=lambda r: (r[2], r[0]))
    rv = wb.create_sheet("Review")
    rv.append(["ID", "Description", "Owner", "Due Date", "Subsystem"])
    for i, r in enumerate(flagged_rows, start=2):
        rv.append([r[0], r[1], r[2], r[4], r[6]])
        rv.cell(i, 4).number_format = "YYYY-MM-DD"
    last = 1 + len(flagged_rows)
    rv.cell(last + 1, 1).value = "Total"
    rv.cell(last + 1, 2).value = f"=COUNTA(A2:A{last})"
    header_row(rv, 5)
    rv.add_table(Table(displayName="Review", ref=f"A1:E{last}", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))

    owners = {}
    for r in flagged_rows:
        owners[r[2]] = owners.get(r[2], 0) + 1
    sm = wb.create_sheet("Summary")
    sm.append(["Owner", "FlaggedCount"])
    for i, (o, n) in enumerate(sorted(owners.items()), start=2):
        sm.append([o, n])
    header_row(sm, 2)
    sm["A6"] = "FlaggedCount"
    sm["B6"] = '=COUNTIF(Actions[Flag],"Flagged")'
    sm["A7"] = "NextCycleCount"
    sm["B7"] = '=COUNTIF(Actions[NextCycle],"Due")'
    chart = BarChart()
    chart.type = "col"
    chart.title = "Open actions by owner"
    chart.y_axis.title = "Items"
    chart.x_axis.title = "Owner"
    chart.y_axis.scaling.min = 0
    chart.legend = None
    data = Reference(sm, min_col=2, min_row=1, max_row=1 + len(owners))
    cats = Reference(sm, min_col=1, min_row=2, max_row=1 + len(owners))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = "1A73E8"
    chart.series[0].dLbls = DataLabelList()
    chart.series[0].dLbls.showVal = True
    sm.add_chart(chart, "D2")
    wb.save(path)


def build_png(path: Path) -> None:
    flagged_rows = [r for r in ACTIONS if flagged(r)]
    owners = {}
    for r in flagged_rows:
        owners[r[2]] = owners.get(r[2], 0) + 1
    names = list(sorted(owners))
    vals = [owners[n] for n in names]
    fig, ax = plt.subplots(figsize=(7.2, 4.2), dpi=140)
    bars = ax.bar(names, vals, color="#1A73E8", width=0.55)
    ax.set_title("Open actions by owner", color="#1A73E8", fontweight="bold")
    ax.set_xlabel("Owner")
    ax.set_ylabel("Items")
    ax.set_ylim(0, max(vals) + 1.2)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.08, str(v), ha="center", color="#1A73E8")
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def chrome(c, w, h):
    c.setFillColor(white)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    c.setFillColor(BLUE)
    c.rect(0, h - 0.12 * cm, w, 0.12 * cm, fill=1, stroke=0)
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.55 * inch, 0.28 * inch, "Confidential Design Operations")
    c.drawRightString(w - 0.55 * inch, 0.28 * inch, str(c.getPageNumber()))


def style_table(data, widths):
    table = RLTable(data, colWidths=widths, rowHeights=[22] * len(data))
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), HEAD),
                ("TEXTCOLOR", (0, 0), (-1, 0), BLUE),
                ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Roboto"),
                ("FONTSIZE", (0, 0), (-1, -1), 11),
                ("TEXTCOLOR", (0, 1), (-1, -1), BODY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("GRID", (0, 0), (-1, -1), 0.4, BLUE),
            ]
        )
    )
    return table


def card(c, x, y, cw, ch):
    c.setFillColor(HEAD)
    c.setStrokeColor(BLUE)
    c.setLineWidth(1)
    c.roundRect(x, y, cw, ch, 8, fill=1, stroke=1)


def build_pdf(path: Path, chart: Path) -> None:
    fonts()
    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle("Q3 Design Review")
    chrome(c, w, h)
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 32)
    c.drawString(0.7 * inch, h / 2 + 0.15 * inch, "Q3 Design Review")
    c.setFillColor(MUTED)
    c.setFont("Roboto", 18)
    c.drawString(0.7 * inch, h / 2 - 0.3 * inch, "Open High and Medium actions")
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Overview")
    t1 = style_table([["Name", "Role"]] + [list(a) for a in ATTENDEES], [2.6 * inch, 3.2 * inch])
    tw, th = t1.wrap(w, h)
    t1.drawOn(c, 0.7 * inch, h - 1.2 * inch - th)
    t2 = style_table([["Topic", "Owner"]] + [list(a) for a in AGENDA], [3.4 * inch, 2.4 * inch])
    tw2, th2 = t2.wrap(w, h)
    t2.drawOn(c, 7.0 * inch, h - 1.2 * inch - th2)
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Decisions")
    t = style_table([["Decision", "Detail"]] + [list(d) for d in DECISIONS], [2.2 * inch, 9.4 * inch])
    tw, th = t.wrap(w, h)
    t.drawOn(c, 0.7 * inch, h - 1.25 * inch - th)
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Action items")
    rows = [["ID", "Description", "Due Date", "Subsystem"]]
    flagged_rows = sorted([r for r in ACTIONS if flagged(r)], key=lambda r: (r[2], r[0]))
    for r in flagged_rows:
        rows.append([r[0], r[1], r[4].isoformat(), r[6]])
    rows.append(["Total", str(len(flagged_rows)), "", ""])
    t = style_table(rows, [1.1 * inch, 4.6 * inch, 1.6 * inch, 1.8 * inch])
    tw, th = t.wrap(w, h)
    t.drawOn(c, 0.7 * inch, h - 1.2 * inch - th)
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Next cycle")
    card(c, 0.7 * inch, h - 2.35 * inch, 3.4 * inch, 1.05 * inch)
    card(c, 4.3 * inch, h - 2.35 * inch, 3.4 * inch, 1.05 * inch)
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.9 * inch, h - 1.55 * inch, "FlaggedCount")
    c.drawString(4.5 * inch, h - 1.55 * inch, "NextCycleCount")
    c.setFillColor(BLUE)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.9 * inch, h - 2.05 * inch, str(sum(1 for r in ACTIONS if flagged(r))))
    c.drawString(4.5 * inch, h - 2.05 * inch, str(sum(1 for r in ACTIONS if next_cycle(r))))
    c.setStrokeColor(BLUE)
    c.rect(8.0 * inch, 1.3 * inch, 4.5 * inch, 4.4 * inch, fill=0, stroke=1)
    c.drawImage(str(chart), 8.05 * inch, 1.35 * inch, width=4.4 * inch, height=4.3 * inch, preserveAspectRatio=True, mask="auto")
    due = [r for r in ACTIONS if next_cycle(r)]
    due.sort(key=lambda r: (r[4], r[0]))
    t = style_table([["ID", "Owner", "Due Date"]] + [[r[0], r[2], r[4].isoformat()] for r in due], [1.3 * inch, 2.4 * inch, 1.8 * inch])
    tw, th = t.wrap(w, h)
    t.drawOn(c, 0.7 * inch, 1.4 * inch)
    c.showPage()
    c.save()


def main():
    ROOT.mkdir(exist_ok=True)
    docx = ROOT / "Design_Review_Meeting_Minutes.docx"
    working = ROOT / "Design_Review_Meeting_Minutes_working.docx"
    xlsx = ROOT / "action_items.xlsx"
    png = ROOT / "open_by_owner.png"
    pdf = ROOT / "Design_Review_Recap.pdf"
    sources = ROOT / "SOURCES.txt"
    build_docx(docx)
    import shutil

    shutil.copyfile(docx, working)
    build_xlsx(xlsx)
    build_png(png)
    build_pdf(pdf, png)
    sources.write_text(
        "Design_Review_Meeting_Minutes.docx — synthetic minutes. Fictional names. No PII.\n",
        encoding="utf-8",
    )
    with ZipFile(REPO / "DesignReview_Assets.zip", "w", ZIP_DEFLATED) as z:
        z.write(docx, "DesignReview/Design_Review_Meeting_Minutes.docx")
        z.write(sources, "DesignReview/SOURCES.txt")
    with ZipFile(REPO / "DesignReview_Expected.zip", "w", ZIP_DEFLATED) as z:
        z.write(working, "DesignReview/Design_Review_Meeting_Minutes_working.docx")
        z.write(xlsx, "DesignReview/action_items.xlsx")
        z.write(png, "DesignReview/open_by_owner.png")
        z.write(pdf, "DesignReview/Design_Review_Recap.pdf")
    print("flagged", [r[0] for r in ACTIONS if flagged(r)])
    print("due", [r[0] for r in ACTIONS if next_cycle(r)])


if __name__ == "__main__":
    main()

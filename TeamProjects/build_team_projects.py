#!/usr/bin/env python3
"""Build Team Projects priorities workbook, gold scan/report, chart PNG, and briefing PDF."""

from __future__ import annotations

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
FONT_DIR = Path("/tmp/roboto")
NAVY = "1F4E79"
NAVY_HEX = HexColor("#1F4E79")
CARD = HexColor("#F4F7FA")
BODY = HexColor("#202124")
MUTED = HexColor("#5F6368")
PAGE = landscape((13.333 * inch, 7.5 * inch))

# Size MB using 1 MiB = 1048576 bytes (TreeSize MB)
FOLDERS = [
    # Folder, OwnerGroup, Review_Required, Note, SizeMB
    ("AtlasCore", "Platform", "Yes", "Core services drop", 48.0),
    ("BeaconDocs", "Docs", "No", "Published manuals", 12.5),
    ("CedarMedia", "Design", "Yes", "Campaign renders", 36.0),
    ("DriftLogs", "Ops", "No", "Rotated logs", 8.0),
    ("EmberData", "Analytics", "Yes", "Warehouse extracts", 64.0),
    ("FalconLegal", "Legal", "No", "Template pack", 15.0),
    ("HarborML", "Research", "Yes", "Checkpoint dumps", 22.0),
    ("IrisWeb", "Web", "No", "Static assets", 10.5),
]
TOTAL = sum(r[4] for r in FOLDERS)


def pct(mb: float) -> float:
    return round(100 * mb / TOTAL, 1)


def register_fonts() -> None:
    pdfmetrics.registerFont(TTFont("Roboto", str(FONT_DIR / "Roboto-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("Roboto-Bold", str(FONT_DIR / "Roboto-Bold.ttf")))


def header(ws, n):
    fill = PatternFill("solid", fgColor=NAVY)
    font = Font(name="Calibri", bold=True, color="FFFFFF")
    for i in range(1, n + 1):
        ws.cell(1, i).fill = fill
        ws.cell(1, i).font = font


def build_priorities(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Priorities"
    ws.append(["Folder", "OwnerGroup", "Review_Required", "Note"])
    for name, owner, req, note, _mb in FOLDERS:
        ws.append([name, owner, req, note])
    header(ws, 4)
    tab = Table(displayName="Priorities", ref="A1:D9")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    for col, w in zip("ABCD", [16, 14, 18, 22]):
        ws.column_dimensions[col].width = w
    wb.save(path)


def build_scan(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Scan"
    ws.append(["Folder", "SizeMB", "PercentOfParent"])
    for name, _o, _r, _n, mb in FOLDERS:
        ws.append([name, mb, pct(mb) / 100.0])
    header(ws, 3)
    for i in range(2, 10):
        ws.cell(i, 2).number_format = "0.0"
        ws.cell(i, 3).number_format = "0.0%"
    tab = Table(displayName="Scan", ref="A1:C9")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    for col, w in zip("ABC", [16, 12, 18]):
        ws.column_dimensions[col].width = w
    wb.save(path)


def build_report(path: Path) -> None:
    wb = Workbook()
    scan = wb.active
    scan.title = "Scan"
    scan.append(["Folder", "SizeMB", "PercentOfParent"])
    for name, _o, _r, _n, mb in FOLDERS:
        scan.append([name, mb, pct(mb) / 100.0])
        r = scan.max_row
        scan.cell(r, 2).number_format = "0.0"
        scan.cell(r, 3).number_format = "0.0%"
    header(scan, 3)
    scan.add_table(Table(displayName="Scan", ref="A1:C9", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))

    pr = wb.create_sheet("Priorities")
    pr.append(["Folder", "OwnerGroup", "Review_Required", "Note", "SizeMB", "PercentOfParent", "Flag"])
    for i, (name, owner, req, note, mb) in enumerate(FOLDERS, start=2):
        pr.append([name, owner, req, note])
        pr.cell(i, 5).value = f"=XLOOKUP(A{i},Scan[Folder],Scan[SizeMB])"
        pr.cell(i, 5).number_format = "0.0"
        pr.cell(i, 6).value = f"=XLOOKUP(A{i},Scan[Folder],Scan[PercentOfParent])"
        pr.cell(i, 6).number_format = "0.0%"
        pr.cell(i, 7).value = f'=IF(C{i}="Yes","Flagged","Clear")'
    header(pr, 7)
    pr.add_table(Table(displayName="Priorities", ref="A1:G9", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))

    flagged = [f for f in FOLDERS if f[2] == "Yes"]
    rv = wb.create_sheet("Review")
    rv.append(["Folder", "SizeMB", "PercentOfParent"])
    for i, (name, _o, _r, _n, mb) in enumerate(flagged, start=2):
        rv.append([name, mb, pct(mb) / 100.0])
        rv.cell(i, 2).number_format = "0.0"
        rv.cell(i, 3).number_format = "0.0%"
    last = 1 + len(flagged)
    rv.cell(last + 1, 1).value = "Total"
    rv.cell(last + 1, 2).value = f"=SUM(B2:B{last})"
    rv.cell(last + 1, 2).number_format = "0.0"
    header(rv, 3)
    rv.add_table(Table(displayName="Review", ref=f"A1:C{last}", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))

    sm = wb.create_sheet("Summary")
    sm.append(["Metric", "Value"])
    sm["A2"] = "FlaggedCount"
    sm["B2"] = '=COUNTIF(Priorities[Flag],"Flagged")'
    sm["A3"] = "TotalFlaggedMB"
    sm["B3"] = f"=SUM(Review[SizeMB])"
    sm["B3"].number_format = "0.0"
    sm["A4"] = "ScanRootMB"
    sm["B4"] = "=SUM(Scan[SizeMB])"
    sm["B4"].number_format = "0.0"
    header(sm, 2)

    chart = BarChart()
    chart.type = "col"
    chart.title = "Folder size (MB)"
    chart.y_axis.title = "Size (MB)"
    chart.x_axis.title = "Folder"
    chart.y_axis.scaling.min = 0
    chart.legend = None
    data = Reference(pr, min_col=5, min_row=1, max_row=9)
    cats = Reference(pr, min_col=1, min_row=2, max_row=9)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = NAVY
    chart.series[0].dLbls = DataLabelList()
    chart.series[0].dLbls.showVal = True
    pr.add_chart(chart, "I2")
    for ws in (scan, pr, rv, sm):
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18
    wb.save(path)


def build_png(path: Path) -> None:
    names = [f[0] for f in FOLDERS]
    vals = [f[4] for f in FOLDERS]
    fig, ax = plt.subplots(figsize=(9.2, 4.6), dpi=140)
    bars = ax.bar(names, vals, color="#1F4E79", width=0.62)
    ax.set_title("Folder size (MB)", color="#1F4E79", fontweight="bold")
    ax.set_xlabel("Folder")
    ax.set_ylabel("Size (MB)")
    ax.set_ylim(0, max(vals) + 12)
    ax.tick_params(axis="x", rotation=20)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + 1.2, f"{v:.1f}", ha="center", color="#1F4E79", fontsize=9)
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def rounded(c, x, y, cw, ch):
    c.setFillColor(CARD)
    c.setStrokeColor(NAVY_HEX)
    c.setLineWidth(1)
    c.roundRect(x, y, cw, ch, 8, fill=1, stroke=1)


def chrome(c, w, h):
    c.setFillColor(white)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    c.setFillColor(NAVY_HEX)
    c.rect(0, h - 0.12 * cm, w, 0.12 * cm, fill=1, stroke=0)
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.55 * inch, 0.28 * inch, "Confidential Workspace Operations")
    c.drawRightString(w - 0.55 * inch, 0.28 * inch, str(c.getPageNumber()))


def build_pdf(path: Path, chart: Path) -> None:
    flagged = [f for f in FOLDERS if f[2] == "Yes"]
    flagged_sum = sum(f[4] for f in flagged)
    largest = max(flagged, key=lambda f: f[4])
    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle("Folder size review")

    chrome(c, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h / 2 + 0.15 * inch, "Folder size review")
    c.setFillColor(MUTED)
    c.setFont("Roboto", 18)
    c.drawString(0.7 * inch, h / 2 - 0.3 * inch, "Top-level folders marked Review_Required")
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Review rule")
    cw, ch = 5.6 * inch, 1.15 * inch
    rounded(c, 0.7 * inch, h - 2.4 * inch, cw, ch)
    rounded(c, 0.7 * inch + cw + 0.35 * inch, h - 2.4 * inch, cw, ch)
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.9 * inch, h - 1.7 * inch, "FlaggedCount")
    c.drawString(0.9 * inch + cw + 0.35 * inch, h - 1.7 * inch, "TotalFlaggedMB")
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.9 * inch, h - 2.15 * inch, str(len(flagged)))
    c.drawString(0.9 * inch + cw + 0.35 * inch, h - 2.15 * inch, f"{flagged_sum:.1f}")
    data = [
        ["Field", "Source", "Rule"],
        ["Flag", "Priorities Review_Required", "Flagged when Review_Required is Yes"],
        ["SizeMB", "TreeSize scan via XLOOKUP", "Size of the top-level folder in MB"],
        ["Total", "Review SizeMB", "SUM of Flagged SizeMB"],
    ]
    table = RLTable(data, colWidths=[2.0 * inch, 3.6 * inch, 5.4 * inch], rowHeights=[24] * 4)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY_HEX),
                ("TEXTCOLOR", (0, 0), (-1, 0), white),
                ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Roboto"),
                ("FONTSIZE", (0, 0), (-1, -1), 12),
                ("TEXTCOLOR", (0, 1), (-1, -1), BODY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, NAVY_HEX),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, CARD]),
            ]
        )
    )
    tw, th = table.wrap(w, h)
    table.drawOn(c, 0.7 * inch, h - 2.65 * inch - th)
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Folder size (MB)")
    img_w, img_h = 7.2 * inch, 4.4 * inch
    c.setStrokeColor(NAVY_HEX)
    c.rect(0.7 * inch, 1.3 * inch, img_w, img_h, fill=0, stroke=1)
    c.drawImage(str(chart), 0.75 * inch, 1.35 * inch, width=img_w - 0.1 * inch, height=img_h - 0.1 * inch, preserveAspectRatio=True, mask="auto")
    c.setFillColor(MUTED)
    c.setFont("Roboto", 12)
    c.drawString(0.7 * inch, 1.08 * inch, "Size in MB for every folder in Priorities")
    bx, by, bw, bh = 8.15 * inch, 1.3 * inch, 4.4 * inch, 4.4 * inch
    rounded(c, bx, by, bw, bh)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 14)
    c.drawString(bx + 0.2 * inch, by + bh - 0.4 * inch, "Key findings")
    c.setFillColor(BODY)
    c.setFont("Roboto", 14)
    lines = [
        f"Largest Flagged: {largest[0]} ({largest[4]:.1f} MB)",
        f"FlaggedCount: {len(flagged)}",
        f"TotalFlaggedMB: {flagged_sum:.1f}",
    ]
    fy = by + bh - 1.1 * inch
    for line in lines:
        c.drawString(bx + 0.2 * inch, fy, line)
        fy -= 0.85 * inch
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Flagged folders")
    rows = [["Folder", "SizeMB", "PercentOfParent"]]
    for name, _o, _r, _n, mb in flagged:
        rows.append([name, f"{mb:.1f}", f"{pct(mb):.1f}%"])
    rows.append(["Total", f"{flagged_sum:.1f}", ""])
    table = RLTable(rows, colWidths=[4.2 * inch, 3.0 * inch, 3.8 * inch], rowHeights=[24] * len(rows))
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
                ("ALIGN", (1, 0), (2, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, NAVY_HEX),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, CARD]),
                ("FONTNAME", (0, -1), (-1, -1), "Roboto-Bold"),
            ]
        )
    )
    tw, th = table.wrap(w, h)
    table.drawOn(c, 0.7 * inch, h - 1.25 * inch - th)
    c.showPage()
    c.save()


def write_sources(path: Path) -> None:
    path.write_text(
        """Team_Projects folder tree is created by seed_team_projects.ps1 (payload.bin sizes in MiB).
folder_priorities.xlsx lists Folder, OwnerGroup, Review_Required, Note. Fictional names. No PII.

Expected
  treesize_scan.xlsx — Folder, SizeMB, PercentOfParent after a TreeSize MB scan
  Large_Folders_Report.xlsx — Scan, Priorities with XLOOKUP and Flag, Review FILTER-equivalent, Summary, chart
  folder_sizes.png
  Large_Folders_Briefing.pdf — four 16:9 slides
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
    (ROOT / "Team_Projects").mkdir(exist_ok=True)
    for name, *_rest in FOLDERS:
        d = ROOT / "Team_Projects" / name
        d.mkdir(exist_ok=True)
        (d / ".gitkeep").write_text("", encoding="utf-8")

    pri = ROOT / "folder_priorities.xlsx"
    scan = ROOT / "treesize_scan.xlsx"
    report = ROOT / "Large_Folders_Report.xlsx"
    png = ROOT / "folder_sizes.png"
    pdf = ROOT / "Large_Folders_Briefing.pdf"
    sources = ROOT / "SOURCES.txt"

    build_priorities(pri)
    build_scan(scan)
    build_report(report)
    build_png(png)
    build_pdf(pdf, png)
    write_sources(sources)

    zip_dir(
        REPO / "TeamProjects_Assets.zip",
        [
            (pri, "TeamProjects/folder_priorities.xlsx"),
            (ROOT / "seed_team_projects.ps1", "TeamProjects/seed_team_projects.ps1"),
            (sources, "TeamProjects/SOURCES.txt"),
        ]
        + [(ROOT / "Team_Projects" / f[0] / ".gitkeep", f"TeamProjects/Team_Projects/{f[0]}/.gitkeep") for f in FOLDERS],
    )
    zip_dir(
        REPO / "TeamProjects_Expected.zip",
        [
            (scan, "TeamProjects/treesize_scan.xlsx"),
            (report, "TeamProjects/Large_Folders_Report.xlsx"),
            (png, "TeamProjects/folder_sizes.png"),
            (pdf, "TeamProjects/Large_Folders_Briefing.pdf"),
        ],
    )
    print("total", TOTAL, "flagged", [f[0] for f in FOLDERS if f[2] == "Yes"])


if __name__ == "__main__":
    main()

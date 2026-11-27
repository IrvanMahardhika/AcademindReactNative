#!/usr/bin/env python3
"""Build Team Projects source tree, messy TreeSize export, quota report, and briefing PDF."""

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
RED = HexColor("#C0392B")
CARD = HexColor("#F4F7FA")
BODY = HexColor("#202124")
MUTED = HexColor("#5F6368")
PAGE = landscape((13.333 * inch, 7.5 * inch))
QUOTA = 25.0
SCAN_ROOT = r"D:\Projects\Workspace\TeamProjects\Team_Projects"

FOLDERS = [
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


def action_for(req: str, mb: float) -> str:
    if req == "Yes" and mb > QUOTA:
        return "Archive"
    if req == "Yes" and mb <= QUOTA:
        return "Hold"
    return "Clear"


def register_fonts() -> None:
    pdfmetrics.registerFont(TTFont("Roboto", str(FONT_DIR / "Roboto-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("Roboto-Bold", str(FONT_DIR / "Roboto-Bold.ttf")))


def header(ws, n):
    fill = PatternFill("solid", fgColor=NAVY)
    font = Font(name="Calibri", bold=True, color="FFFFFF")
    for i in range(1, n + 1):
        ws.cell(1, i).fill = fill
        ws.cell(1, i).font = font


def ensure_payloads() -> None:
    mib = 1024 * 1024
    chunk = b"\0" * mib
    for name, *_rest, mb in FOLDERS:
        n = int(mb * mib)
        p = ROOT / "Team_Projects" / name / "payload.bin"
        p.parent.mkdir(parents=True, exist_ok=True)
        (p.parent / ".gitkeep").write_text("", encoding="utf-8")
        if p.exists() and p.stat().st_size == n:
            continue
        remaining = n
        with open(p, "wb") as fh:
            while remaining:
                w = min(len(chunk), remaining)
                fh.write(chunk[:w])
                remaining -= w


def build_priorities(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Priorities"
    ws.append(["Folder", "OwnerGroup", "Review_Required", "Note"])
    for name, owner, req, note, _mb in FOLDERS:
        ws.append([name, owner, req, note])
    header(ws, 4)
    ws.add_table(Table(displayName="PrioritiesSrc", ref="A1:D9", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))
    for col, w in zip("ABCD", [16, 14, 18, 22]):
        ws.column_dimensions[col].width = w
    wb.save(path)


def build_raw_scan(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Scan"
    ws.append(["Path", "Size", "Percent of Parent", "Files"])
    for name, _o, _r, _n, mb in FOLDERS:
        ws.append([rf"{SCAN_ROOT}\{name}", f"{mb:.1f} MB", pct(mb) / 100.0, 1])
        r = ws.max_row
        ws.cell(r, 3).number_format = "0.0%"
    header(ws, 4)
    for col, w in zip("ABCD", [62, 12, 20, 10]):
        ws.column_dimensions[col].width = w
    wb.save(path)


def build_report(path: Path) -> None:
    wb = Workbook()
    ctrl = wb.active
    ctrl.title = "Control"
    ctrl["A1"] = "Quota"
    ctrl["B1"] = QUOTA
    ctrl["B1"].number_format = "0.0"
    header(ctrl, 2)

    scan = wb.create_sheet("Scan")
    scan.append(["Folder", "SizeMB", "PercentOfParent"])
    for name, _o, _r, _n, mb in FOLDERS:
        scan.append([name, mb, pct(mb) / 100.0])
        r = scan.max_row
        scan.cell(r, 2).number_format = "0.0"
        scan.cell(r, 3).number_format = "0.0%"
    header(scan, 3)
    scan.add_table(Table(displayName="Scan", ref="A1:C9", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))

    pr = wb.create_sheet("Priorities")
    pr.append(["Folder", "OwnerGroup", "Review_Required", "Note", "SizeMB", "PercentOfParent", "Action", "Quota"])
    for i, (name, owner, req, note, mb) in enumerate(FOLDERS, start=2):
        pr.append([name, owner, req, note])
        pr.cell(i, 5).value = f"=XLOOKUP(A{i},Scan[Folder],Scan[SizeMB])"
        pr.cell(i, 5).number_format = "0.0"
        pr.cell(i, 6).value = f"=XLOOKUP(A{i},Scan[Folder],Scan[PercentOfParent])"
        pr.cell(i, 6).number_format = "0.0%"
        pr.cell(i, 7).value = (
            f'=IF(AND(C{i}="Yes",E{i}>Control!$B$1),"Archive",'
            f'IF(AND(C{i}="Yes",E{i}<=Control!$B$1),"Hold","Clear"))'
        )
        pr.cell(i, 8).value = "=Control!$B$1"
        pr.cell(i, 8).number_format = "0.0"
    header(pr, 8)
    pr.add_table(Table(displayName="Priorities", ref="A1:H9", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))

    archive = [f for f in FOLDERS if action_for(f[2], f[4]) == "Archive"]
    hold = [f for f in FOLDERS if action_for(f[2], f[4]) == "Hold"]

    rv = wb.create_sheet("Review")
    rv.append(["Folder", "OwnerGroup", "SizeMB", "PercentOfParent"])
    for i, (name, owner, _r, _n, mb) in enumerate(archive, start=2):
        rv.append([name, owner, mb, pct(mb) / 100.0])
        rv.cell(i, 3).number_format = "0.0"
        rv.cell(i, 4).number_format = "0.0%"
    last = 1 + len(archive)
    rv.cell(last + 1, 1).value = "Total"
    rv.cell(last + 1, 3).value = f"=SUM(C2:C{last})"
    rv.cell(last + 1, 3).number_format = "0.0"
    header(rv, 4)
    rv.add_table(Table(displayName="Review", ref=f"A1:D{last}", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))

    hd = wb.create_sheet("Hold")
    hd.append(["Folder", "OwnerGroup", "SizeMB", "PercentOfParent"])
    for i, (name, owner, _r, _n, mb) in enumerate(hold, start=2):
        hd.append([name, owner, mb, pct(mb) / 100.0])
        hd.cell(i, 3).number_format = "0.0"
        hd.cell(i, 4).number_format = "0.0%"
    header(hd, 4)
    hd.add_table(Table(displayName="HoldList", ref=f"A1:D{1 + len(hold)}", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))

    sm = wb.create_sheet("Summary")
    sm.append(["Metric", "Value"])
    sm["A2"] = "ArchiveCount"
    sm["B2"] = '=COUNTIF(Priorities[Action],"Archive")'
    sm["A3"] = "HoldCount"
    sm["B3"] = '=COUNTIF(Priorities[Action],"Hold")'
    sm["A4"] = "ArchiveMB"
    sm["B4"] = f"=SUM(Review[SizeMB])"
    sm["B4"].number_format = "0.0"
    sm["A5"] = "ScanRootMB"
    sm["B5"] = "=SUM(Scan[SizeMB])"
    sm["B5"].number_format = "0.0"
    header(sm, 2)

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Folder size vs quota (MB)"
    chart.y_axis.title = "Size (MB)"
    chart.x_axis.title = "Folder"
    chart.y_axis.scaling.min = 0
    data = Reference(pr, min_col=5, min_row=1, max_row=9)
    quota_ref = Reference(pr, min_col=8, min_row=1, max_row=9)
    cats = Reference(pr, min_col=1, min_row=2, max_row=9)
    chart.add_data(data, titles_from_data=True)
    chart.add_data(quota_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = NAVY
    chart.series[1].graphicalProperties.solidFill = "C0392B"
    chart.series[0].dLbls = DataLabelList()
    chart.series[0].dLbls.showVal = True
    pr.add_chart(chart, "J2")
    for ws in (scan, pr, rv, hd, sm):
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18
    wb.save(path)


def build_png(path: Path) -> None:
    names = [f[0] for f in FOLDERS]
    vals = [f[4] for f in FOLDERS]
    fig, ax = plt.subplots(figsize=(9.2, 4.6), dpi=140)
    x = range(len(names))
    ax.bar(list(x), vals, color="#1F4E79", width=0.55, label="SizeMB")
    ax.plot(list(x), [QUOTA] * len(names), color="#C0392B", linewidth=2, label="Quota")
    ax.set_xticks(list(x), names, rotation=20, ha="right")
    ax.set_title("Folder size vs quota (MB)", color="#1F4E79", fontweight="bold")
    ax.set_xlabel("Folder")
    ax.set_ylabel("Size (MB)")
    ax.set_ylim(0, max(vals) + 12)
    ax.legend(loc="lower center", ncol=2, bbox_to_anchor=(0.5, -0.38))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for i, v in enumerate(vals):
        ax.text(i, v + 1.2, f"{v:.1f}", ha="center", color="#1F4E79", fontsize=9)
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


def tbl(data, widths):
    table = RLTable(data, colWidths=widths, rowHeights=[22] * len(data))
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY_HEX),
                ("TEXTCOLOR", (0, 0), (-1, 0), white),
                ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Roboto"),
                ("FONTSIZE", (0, 0), (-1, -1), 11),
                ("TEXTCOLOR", (0, 1), (-1, -1), BODY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("GRID", (0, 0), (-1, -1), 0.4, NAVY_HEX),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, CARD]),
            ]
        )
    )
    return table


def build_pdf(path: Path, chart: Path) -> None:
    archive = [f for f in FOLDERS if action_for(f[2], f[4]) == "Archive"]
    hold = [f for f in FOLDERS if action_for(f[2], f[4]) == "Hold"]
    archive_mb = sum(f[4] for f in archive)
    largest = max(archive, key=lambda f: f[4])
    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle("Folder size review")

    chrome(c, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h / 2 + 0.2 * inch, "Folder size review")
    c.setFillColor(MUTED)
    c.setFont("Roboto", 18)
    c.drawString(0.7 * inch, h / 2 - 0.25 * inch, "Archive Review_Required folders over the 25 MB quota")
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Quota rule")
    cards = [
        (0.7 * inch, "ArchiveCount", str(len(archive))),
        (4.85 * inch, "HoldCount", str(len(hold))),
        (9.0 * inch, "ArchiveMB", f"{archive_mb:.1f}"),
    ]
    for x, lab, val in cards:
        rounded(c, x, h - 2.45 * inch, 3.7 * inch, 1.2 * inch)
        c.setFillColor(MUTED)
        c.setFont("Roboto", 12)
        c.drawString(x + 0.2 * inch, h - 1.55 * inch, lab)
        c.setFillColor(NAVY_HEX)
        c.setFont("Roboto-Bold", 28)
        c.drawString(x + 0.2 * inch, h - 2.15 * inch, val)
    data = [
        ["Field", "Source", "Rule"],
        ["Archive", "Review_Required and SizeMB", "Yes and SizeMB greater than Quota"],
        ["Hold", "Review_Required and SizeMB", "Yes and SizeMB at or under Quota"],
        ["ArchiveMB", "Review SizeMB", "SUM of Archive SizeMB"],
    ]
    table = tbl(data, [2.0 * inch, 3.8 * inch, 5.4 * inch])
    tw, th = table.wrap(w, h)
    table.drawOn(c, 0.7 * inch, h - 2.75 * inch - th)
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Folder size vs quota (MB)")
    c.setStrokeColor(NAVY_HEX)
    c.rect(0.7 * inch, 1.25 * inch, 7.2 * inch, 4.4 * inch, fill=0, stroke=1)
    c.drawImage(str(chart), 0.75 * inch, 1.3 * inch, width=7.1 * inch, height=4.3 * inch, preserveAspectRatio=True, mask="auto")
    rounded(c, 8.15 * inch, 1.25 * inch, 4.4 * inch, 4.4 * inch)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 14)
    c.drawString(8.35 * inch, 5.3 * inch, "Key findings")
    c.setFillColor(BODY)
    c.setFont("Roboto", 14)
    fy = 4.7 * inch
    for line in [
        f"Largest Archive: {largest[0]} ({largest[4]:.1f} MB)",
        f"ArchiveCount: {len(archive)}",
        f"HoldCount: {len(hold)}",
        f"ArchiveMB: {archive_mb:.1f}",
    ]:
        c.drawString(8.35 * inch, fy, line)
        fy -= 0.7 * inch
    c.showPage()

    chrome(c, w, h)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 28)
    c.drawString(0.7 * inch, h - 0.95 * inch, "Archive list")
    rows = [["Folder", "OwnerGroup", "SizeMB", "PercentOfParent"]]
    for name, owner, _r, _n, mb in archive:
        rows.append([name, owner, f"{mb:.1f}", f"{pct(mb):.1f}%"])
    rows.append(["Total", "", f"{archive_mb:.1f}", ""])
    table = tbl(rows, [2.4 * inch, 2.0 * inch, 1.5 * inch, 2.0 * inch])
    tw, th = table.wrap(w, h)
    table.drawOn(c, 0.7 * inch, h - 1.2 * inch - th)
    c.setFillColor(NAVY_HEX)
    c.setFont("Roboto-Bold", 16)
    c.drawString(8.0 * inch, h - 1.15 * inch, "Hold (under quota)")
    hrows = [["Folder", "OwnerGroup", "SizeMB", "PercentOfParent"]]
    for name, owner, _r, _n, mb in hold:
        hrows.append([name, owner, f"{mb:.1f}", f"{pct(mb):.1f}%"])
    ht = tbl(hrows, [1.7 * inch, 1.5 * inch, 1.1 * inch, 1.6 * inch])
    tw, th = ht.wrap(w, h)
    ht.drawOn(c, 8.0 * inch, h - 1.45 * inch - th)
    c.showPage()
    c.save()


def write_sources(path: Path) -> None:
    path.write_text(
        """Team_Projects/<name>/payload.bin sizes (MiB): AtlasCore 48, BeaconDocs 12.5, CedarMedia 36,
DriftLogs 8, EmberData 64, FalconLegal 15, HarborML 22, IrisWeb 10.5.
folder_priorities.xlsx — Folder, OwnerGroup, Review_Required, Note. Fictional names. No PII.

Quota is 25 MB. Archive = Review_Required Yes and SizeMB > 25. Hold = Yes and SizeMB <= 25.

Expected
  treesize_scan.xlsx — raw TreeSize-style export (full path, Size as text with MB)
  Large_Folders_Report.xlsx — Control, cleaned Scan, XLOOKUP, Action, Review, Hold, Summary, chart
  folder_sizes.png — SizeMB vs Quota
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
    ensure_payloads()
    pri = ROOT / "folder_priorities.xlsx"
    scan = ROOT / "treesize_scan.xlsx"
    report = ROOT / "Large_Folders_Report.xlsx"
    png = ROOT / "folder_sizes.png"
    pdf = ROOT / "Large_Folders_Briefing.pdf"
    sources = ROOT / "SOURCES.txt"
    build_priorities(pri)
    build_raw_scan(scan)
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
        + [(ROOT / "Team_Projects" / f[0] / ".gitkeep", f"TeamProjects/Team_Projects/{f[0]}/.gitkeep") for f in FOLDERS]
        + [(ROOT / "Team_Projects" / f[0] / "payload.bin", f"TeamProjects/Team_Projects/{f[0]}/payload.bin") for f in FOLDERS],
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
    print("archive", [f[0] for f in FOLDERS if action_for(f[2], f[4]) == "Archive"])
    print("hold", [f[0] for f in FOLDERS if action_for(f[2], f[4]) == "Hold"])


if __name__ == "__main__":
    main()

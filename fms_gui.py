import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from datetime import datetime, timedelta
import subprocess
from io import BytesIO
import random


def decode_xml_bytes(data):
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def safe_int(value, default=0):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def is_cancelled(status):
    return "avmeld" in (status or "").strip().lower()


def parse_competition(xml_text, log):
    rows = []
    root = ET.fromstring(xml_text)
    comp = root.find("Competition")
    if comp is None:
        return rows

    for participant in comp.findall("Participant"):
        given = (participant.attrib.get("GivenName") or "").strip()
        family = (participant.attrib.get("FamilyName") or "").strip()
        print_name = (participant.attrib.get("PrintName") or "").strip()
        gender = (participant.attrib.get("Gender") or "").strip()
        org = (participant.attrib.get("Organisation") or "").strip()
        participant_code = (participant.attrib.get("Code") or "").strip()

        for discipline in participant.findall("Discipline"):
            for reg in discipline.findall("RegisteredEvent"):
                event_code = (reg.attrib.get("Event") or "").strip()
                entry_order = ""
                music = {}
                clubs = {}
                elements_free = []
                elements_short = []

                for entry in reg.findall("EventEntry"):
                    code = (entry.attrib.get("Code") or "").strip()
                    pos = safe_int(entry.attrib.get("Pos"))
                    val = (entry.attrib.get("Value") or "").strip()

                    if code == "ENTRY_ORDER":
                        entry_order = val
                    elif code == "MUSIC":
                        if pos:
                            music[pos] = val
                    elif code == "CLUB":
                        if pos:
                            clubs[pos] = val
                    elif code == "ELEMENT_CODE_FREE":
                        elements_free.append((pos, val))
                    elif code == "ELEMENT_CODE_SHORT":
                        elements_short.append((pos, val))

                elements_free = [v for _, v in sorted(elements_free) if v]
                elements_short = [v for _, v in sorted(elements_short) if v]

                log(f"Leser deltager: {print_name} (Event: {event_code})")
                rows.append(
                    {
                        "PrintName": print_name,
                        "GivenName": given,
                        "FamilyName": family,
                        "Gender": gender,
                        "Organisation": org,
                        "ParticipantCode": participant_code,
                        "Event": event_code,
                        "EntryOrder": entry_order,
                        "Music1": music.get(1, ""),
                        "Music2": music.get(2, ""),
                        "Club1": clubs.get(1, ""),
                        "Club2": clubs.get(2, ""),
                        "ElementsFree": ", ".join(elements_free),
                        "ElementsShort": ", ".join(elements_short),
                    }
                )

    return rows


def normalize_name(value):
    return " ".join((value or "").strip().lower().split())


def normalize_text(value):
    return " ".join((value or "").strip().lower().split())


def name_matches_filename(given, family, filename):
    target1 = normalize_text(f"{given} {family}")
    target2 = normalize_text(f"{family} {given}")
    hay = normalize_text(filename)
    return (target1 and target1 in hay) or (target2 and target2 in hay)


def format_duration(seconds):
    if seconds is None:
        return ""
    total = int(round(seconds))
    mins = total // 60
    secs = total % 60
    return f"{mins}:{secs:02d}"


def parse_time_hhmm(value):
    if not value:
        return None
    parts = [p.strip() for p in value.split(":")]
    try:
        if len(parts) == 2:
            h, m = int(parts[0]), int(parts[1])
            return datetime(2000, 1, 1, h, m, 0)
        if len(parts) == 3:
            h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
            return datetime(2000, 1, 1, h, m, s)
    except ValueError:
        return None
    return None


def parse_duration_mmss(value):
    if not value:
        return None
    parts = [p.strip() for p in value.split(":")]
    try:
        if len(parts) == 2:
            m, s = int(parts[0]), int(parts[1])
            return m * 60 + s
        if len(parts) == 3:
            h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
            return h * 3600 + m * 60 + s
    except ValueError:
        return None
    return None


def is_registered(status):
    if is_cancelled(status):
        return False
    if not status:
        return False
    text = str(status).strip().lower()
    if "ikke sjekket inn" in text:
        return True
    return "påmeld" in text or "registr" in text or "bekreftet" in text


def load_participants_from_excel(excel_path, log):
    try:
        import openpyxl
    except Exception:
        log("Mangler openpyxl. Installer med: pip install openpyxl")
        return []

    if not Path(excel_path).exists():
        log(f"Fant ikke excel: {excel_path}")
        return []

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
    except Exception as exc:
        log(f"Kunne ikke lese excel: {excel_path} ({exc})")
        return []

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        log("Excel er tom.")
        return []

    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    header_map = {h: idx for idx, h in enumerate(headers)}

    def idx(*names):
        for name in names:
            if name in header_map:
                return header_map[name]
        return None

    i_given = idx("Fornavn")
    i_family = idx("Etternavn")
    i_gender = idx("Kjønn")
    i_club = idx("Klubb")
    i_status = idx("Påmelding", "Påmeldingsstatus", "Deltakerstatus")

    if i_given is None or i_family is None:
        log("Finner ikke nødvendige kolonner i excel (Fornavn/Etternavn).")
        return []

    out = []
    for row in rows[1:]:
        given = row[i_given] if i_given < len(row) else ""
        family = row[i_family] if i_family < len(row) else ""
        gender = row[i_gender] if i_gender is not None and i_gender < len(row) else ""
        club = row[i_club] if i_club is not None and i_club < len(row) else ""
        status = row[i_status] if i_status is not None and i_status < len(row) else ""

        if not (given or family):
            continue

        print_name = f"{str(family).strip()}, {str(given).strip()}".strip(", ")
        out.append(
            {
                "PrintName": print_name,
                "GivenName": (str(given).strip() if given is not None else ""),
                "FamilyName": (str(family).strip() if family is not None else ""),
                "Gender": (str(gender).strip() if gender is not None else ""),
                "Organisation": (str(club).strip() if club is not None else ""),
                "ParticipantCode": "",
                "Event": "",
                "EntryOrder": "",
                "Påmelding": (str(status).strip() if status is not None else ""),
                "Music1": "",
                "Music2": "",
                "Club1": "",
                "Club2": "",
                "ElementsFree": "",
                "ElementsShort": "",
                "Manglende i zip": "",
            }
        )

    log(f"Lest excel: {excel_path} ({len(out)} rader)")
    return out


def parse_officials(xml_text, log):
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return 0
    comp = root.find("Competition")
    if comp is None:
        log("Ingen officials funnet i judges-filen.")
        return 0
    officials = comp.findall("Official")
    log(f"Fant {len(officials)} officials i judges-filen.")
    return len(officials)


def generate_excel(rows, out_path, log):
    try:
        import openpyxl
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Font
    except Exception:
        log("Mangler openpyxl. Installer med: pip install openpyxl")
        return False

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Participants"

    headers = [
        "PrintName",
        "Organisation",
        "ParticipantCode",
        "Event",
        "Påmelding",
        "Manglende i zip",
        "Musikk",
        "MusikkTid",
        "Club1",
        "Club2",
        "ElementsFree",
        "ElementsShort",
    ]
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
        if is_cancelled(row.get("Påmelding", "")):
            from openpyxl.styles import PatternFill

            fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
            for cell in ws[ws.max_row]:
                cell.fill = fill
                cell.font = Font(color="7A0B0B")

    for idx, _ in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = 18

    try:
        wb.save(out_path)
    except PermissionError:
        log(f"Kunne ikke skrive Excel (filen er trolig åpen): {out_path}")
        return False
    except Exception as exc:
        log(f"Kunne ikke skrive Excel: {out_path} ({exc})")
        return False
    log(f"Excel skrevet: {out_path}")
    return True


def generate_html(rows, out_path, title, log):
    headers = [
        "PrintName",
        "Organisation",
        "ParticipantCode",
        "Event",
        "Påmelding",
        "Manglende i zip",
        "Musikk",
        "MusikkTid",
        "Club1",
        "Club2",
        "ElementsFree",
        "ElementsShort",
    ]

    def esc(s):
        return (
            str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    rows_html = []
    for row in rows:
        row_style = ""
        if is_cancelled(row.get("Påmelding", "")):
            row_style = ' style="background:#f8d7da;color:#7a0b0b;"'
        cells = "".join(f"<td>{esc(row.get(h, ''))}</td>" for h in headers)
        rows_html.append(f"<tr{row_style}>{cells}</tr>")

    html = f"""<!doctype html>
<html lang="no">
<head>
  <meta charset="utf-8">
  <title>{esc(title)}</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; }}
    table {{ border-collapse: collapse; width: 100%; }}
    th, td {{ border: 1px solid #ccc; padding: 6px 8px; text-align: left; }}
    th {{ background: #f4f4f4; }}
  </style>
</head>
<body>
  <h2>{esc(title)}</h2>
  <table>
    <thead>
      <tr>
        {''.join(f'<th>{esc(h)}</th>' for h in headers)}
      </tr>
    </thead>
    <tbody>
      {''.join(rows_html)}
    </tbody>
  </table>
</body>
</html>
"""

    try:
        Path(out_path).write_text(html, encoding="utf-8")
    except PermissionError:
        log(f"Kunne ikke skrive HTML (filen er trolig åpen): {out_path}")
        return False
    except Exception as exc:
        log(f"Kunne ikke skrive HTML: {out_path} ({exc})")
        return False
    log(f"HTML skrevet: {out_path}")
    return True


def generate_pdf(rows, out_path, title, log):
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import (
            SimpleDocTemplate,
            Table,
            TableStyle,
            Paragraph,
            Spacer,
            Image,
        )
    except Exception:
        log("Mangler reportlab. Installer med: pip install reportlab")
        return False

    headers = [
        "PrintName",
        "Organisation",
        "Event",
        "Påmelding",
        "Manglende i zip",
        "Musikk",
        "MusikkTid",
        "ElementsFree",
    ]
    data = [headers]
    for row in rows:
        data.append([row.get(h, "") for h in headers])

    doc = SimpleDocTemplate(out_path, pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Loddefjord IL Kunstløp", styles["Title"]),
        Spacer(1, 6),
    ]
    logo_path = Path(__file__).resolve().parent / "Lil Logo.jpg"
    if logo_path.exists():
        try:
            logo = Image(str(logo_path))
            logo.drawHeight = 50
            logo.drawWidth = 50 * (logo.imageWidth / logo.imageHeight)
            story.append(logo)
            story.append(Spacer(1, 6))
        except Exception:
            log("Klarte ikke å lese logo-bildet.")
    story.append(Paragraph(title, styles["Heading2"]))
    story.append(Spacer(1, 12))

    table = Table(data, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    for idx, row in enumerate(rows, start=1):
        if is_cancelled(row.get("Påmelding", "")):
            style_cmds.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#F8D7DA")))
            style_cmds.append(("TEXTCOLOR", (0, idx), (-1, idx), colors.HexColor("#7A0B0B")))
    table.setStyle(TableStyle(style_cmds))
    story.append(table)
    try:
        doc.build(story)
    except PermissionError:
        log(f"Kunne ikke skrive PDF (filen er trolig åpen): {out_path}")
        return False
    except Exception as exc:
        log(f"Kunne ikke skrive PDF: {out_path} ({exc})")
        return False
    log(f"PDF skrevet: {out_path}")
    return True


def build_startliste(rows, group_size, interval_seconds, start_time):
    entries = []
    if not rows:
        return entries
    group_size = max(1, group_size)
    interval_seconds = max(1, interval_seconds)

    index = 0
    group_num = 1
    while index < len(rows):
        group_start_dt = start_time + timedelta(seconds=index * interval_seconds)
        group_label_time = group_start_dt.strftime("%H:%M:%S")
        if group_num > 1:
            group_label_time = f"ca. {group_label_time}"
        entries.append(
            {
                "is_group": True,
                "start": group_label_time,
                "nr": "",
                "navn": f"Oppvarmingsgruppe {group_num}",
                "klubb": "",
            }
        )
        group_rows = rows[index : index + group_size]
        for offset, row in enumerate(group_rows, start=1):
            entries.append(
                {
                    "is_group": False,
                    "start": "",
                    "nr": index + offset,
                    "navn": f"{row.get('GivenName', '')} {row.get('FamilyName', '')}".strip(),
                    "klubb": row.get("Organisation", ""),
                }
            )
        index += group_size
        group_num += 1
    return entries


def generate_startliste_excel(entries, out_path, title, log):
    try:
        import openpyxl
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Font, PatternFill
    except Exception:
        log("Mangler openpyxl. Installer med: pip install openpyxl")
        return False

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Startliste"

    ws.append([title])
    ws.append([])
    headers = ["Start (ca)", "Nr.", "Navn", "Klubb"]
    ws.append(headers)

    header_font = Font(bold=True)
    for cell in ws[3]:
        cell.font = header_font

    group_fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
    group_font = Font(bold=True)

    for entry in entries:
        ws.append([entry["start"], entry["nr"], entry["navn"], entry["klubb"]])
        if entry["is_group"]:
            for cell in ws[ws.max_row]:
                cell.font = group_font
                cell.fill = group_fill

    ws.column_dimensions[get_column_letter(1)].width = 14
    ws.column_dimensions[get_column_letter(2)].width = 6
    ws.column_dimensions[get_column_letter(3)].width = 38
    ws.column_dimensions[get_column_letter(4)].width = 16

    try:
        wb.save(out_path)
    except PermissionError:
        log(f"Kunne ikke skrive Excel (filen er trolig åpen): {out_path}")
        return False
    except Exception as exc:
        log(f"Kunne ikke skrive Excel: {out_path} ({exc})")
        return False

    log(f"Startliste Excel skrevet: {out_path}")
    return True


def generate_startliste_pdf(entries, out_path, title, log):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    except Exception:
        log("Mangler reportlab. Installer med: pip install reportlab")
        return False

    data = [["Start (ca)", "Nr.", "Navn", "Klubb"]]
    for entry in entries:
        data.append([entry["start"], entry["nr"], entry["navn"], entry["klubb"]])

    doc = SimpleDocTemplate(out_path, pagesize=A4)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles["Title"]), Spacer(1, 12)]

    table = Table(data, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    row_idx = 1
    for entry in entries:
        if entry["is_group"]:
            style_cmds.append(
                ("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor("#EEEEEE"))
            )
            style_cmds.append(("FONTNAME", (0, row_idx), (-1, row_idx), "Helvetica-Bold"))
        row_idx += 1
    table.setStyle(TableStyle(style_cmds))
    story.append(table)

    try:
        doc.build(story)
    except PermissionError:
        log(f"Kunne ikke skrive PDF (filen er trolig åpen): {out_path}")
        return False
    except Exception as exc:
        log(f"Kunne ikke skrive PDF: {out_path} ({exc})")
        return False

    log(f"Startliste PDF skrevet: {out_path}")
    return True


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("FSM Data Dekoder")
        self.root.minsize(900, 620)
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 9))
        self.rows = []
        self.zip_path = None

        base_dir = Path(__file__).resolve().parent
        self.folder_var = tk.StringVar(value=str(base_dir))

        container = ttk.Frame(root, padding=10)
        container.pack(fill="both", expand=True)

        header = ttk.Frame(container)
        header.pack(fill="x", pady=(0, 8))
        ttk.Label(header, text="FSM Data Dekoder", style="Title.TLabel").pack(
            side="left"
        )
        ttk.Label(
            header,
            text="Skann zip, sjekk musikkfiler og lag rapporter",
            style="Subtitle.TLabel",
        ).pack(side="left", padx=12)

        folder_frame = ttk.Labelframe(container, text="Kilde")
        folder_frame.pack(fill="x", pady=6)
        ttk.Label(folder_frame, text="Mappe med zip:").pack(side="left", padx=(8, 4))
        ttk.Entry(folder_frame, textvariable=self.folder_var, width=60).pack(
            side="left", padx=6
        )
        ttk.Button(folder_frame, text="Velg mappe", command=self.choose_folder).pack(
            side="left"
        )

        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill="x", pady=4)
        ttk.Button(btn_frame, text="Skann filer", command=self.read_zip).pack(side="left")
        ttk.Button(btn_frame, text="?", width=3, command=self.show_about).pack(
            side="right"
        )

        log_frame = ttk.Labelframe(container, text="Logg")
        log_frame.pack(fill="both", expand=True, pady=6)
        self.log_widget = ScrolledText(log_frame, height=12, wrap="word")
        self.log_widget.pack(fill="both", expand=True, padx=6, pady=6)

        out_frame = ttk.Labelframe(container, text="Rapporter")
        out_frame.pack(fill="x", pady=6)
        self.var_pdf = tk.BooleanVar(value=False)
        self.var_excel = tk.BooleanVar(value=True)
        self.var_html = tk.BooleanVar(value=True)
        self.chk_pdf = ttk.Checkbutton(out_frame, text="PDF", variable=self.var_pdf)
        self.chk_excel = ttk.Checkbutton(out_frame, text="Excel", variable=self.var_excel)
        self.chk_html = ttk.Checkbutton(out_frame, text="HTML", variable=self.var_html)
        self.chk_pdf.pack(side="left", padx=6, pady=6)
        self.chk_excel.pack(side="left", padx=6, pady=6)
        self.chk_html.pack(side="left", padx=6, pady=6)
        self.btn_generate = ttk.Button(
            out_frame, text="Lag filer", command=self.generate_files, state="disabled"
        )
        self.btn_generate.pack(side="right", padx=6, pady=6)

        start_frame = ttk.Labelframe(container, text="Startliste")
        start_frame.pack(fill="x", pady=6)
        today_str = datetime.now().strftime("%d.%m.%y")
        self.start_date_var = tk.StringVar(value=today_str)
        self.start_time_var = tk.StringVar(value="18:00")
        self.interval_var = tk.StringVar(value="3:40")
        self.group_size_var = tk.StringVar(value="8")
        self.location_var = tk.StringVar(value="iskanten")

        ttk.Label(start_frame, text="Dato:").pack(side="left", padx=(6, 2))
        try:
            from tkcalendar import DateEntry

            self.date_widget = DateEntry(
                start_frame,
                textvariable=self.start_date_var,
                date_pattern="dd.mm.yy",
                width=10,
            )
        except Exception:
            self.date_widget = ttk.Entry(
                start_frame, textvariable=self.start_date_var, width=10
            )
        self.date_widget.pack(side="left", padx=4)
        ttk.Label(start_frame, text="Start kl:").pack(side="left")
        ttk.Entry(start_frame, textvariable=self.start_time_var, width=8).pack(
            side="left", padx=4
        )
        ttk.Label(start_frame, text="Intervall:").pack(side="left")
        ttk.Entry(start_frame, textvariable=self.interval_var, width=6).pack(
            side="left", padx=4
        )
        ttk.Label(start_frame, text="Gruppe str:").pack(side="left")
        ttk.Entry(start_frame, textvariable=self.group_size_var, width=4).pack(
            side="left", padx=4
        )
        ttk.Label(start_frame, text="Sted:").pack(side="left")
        ttk.Entry(start_frame, textvariable=self.location_var, width=14).pack(
            side="left", padx=4
        )
        self.shuffle_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            start_frame, text="Tilfeldig rekkefølge", variable=self.shuffle_var
        ).pack(side="left", padx=6)
        self.btn_startliste = ttk.Button(
            start_frame,
            text="Lag startliste (PDF+Excel)",
            command=self.generate_startliste,
            state="disabled",
        )
        self.btn_startliste.pack(side="right", padx=6)

        self.set_output_controls(enabled=False)

    def log(self, msg):
        self.log_widget.insert("end", msg + "\n")
        self.log_widget.see("end")
        self.root.update_idletasks()

    def set_output_controls(self, enabled):
        state = "normal" if enabled else "disabled"
        self.chk_pdf.config(state=state)
        self.chk_excel.config(state=state)
        self.chk_html.config(state=state)
        self.btn_generate.config(state=state)
        self.btn_startliste.config(state=state)

    def choose_folder(self):
        path = filedialog.askdirectory(initialdir=self.folder_var.get())
        if path:
            self.folder_var.set(path)

    def read_zip(self):
        self.log_widget.delete("1.0", "end")
        self.rows = []
        self.zip_path = None

        folder = Path(self.folder_var.get())
        if not folder.exists():
            messagebox.showerror("Feil", "Mappen finnes ikke.")
            return

        zips = sorted(folder.glob("*.zip"))
        if not zips:
            messagebox.showerror("Feil", "Fant ingen zip-filer i mappen.")
            return

        if len(zips) > 1:
            self.log(f"Fant flere zip-filer. Bruker: {zips[0].name}")
        self.zip_path = zips[0]

        excel_name = "Deltakerliste - KUNSTLØP_ Oppvisning Bergen.xlsx"
        excel_path = self.zip_path.parent / excel_name
        self.rows = load_participants_from_excel(excel_path, self.log)
        if not self.rows:
            messagebox.showerror("Feil", "Fant ingen deltakere i excel-filen.")
            self.set_output_controls(enabled=False)
            return

        self.log(f"Leser zip: {self.zip_path.name}")
        zip_rows = []
        with zipfile.ZipFile(self.zip_path, "r") as zf:
            xml_entries = [e for e in zf.infolist() if e.filename.lower().endswith(".xml")]
            if not xml_entries:
                messagebox.showerror("Feil", "Fant ingen xml-filer i zip.")
                return

            for entry in xml_entries:
                self.log(f"Leser fil: {entry.filename}")
                data = zf.read(entry)
                xml_text = decode_xml_bytes(data)
                if "judges" in entry.filename.lower():
                    parse_officials(xml_text, self.log)
                else:
                    zip_rows.extend(parse_competition(xml_text, self.log))

        music_zip = None
        for z in zips:
            if "musikk" in z.name.lower():
                music_zip = z
                break

        music_files = []
        music_durations = {}
        if music_zip:
            self.log(f"Leser musikk-zip: {music_zip.name}")
            try:
                with zipfile.ZipFile(music_zip, "r") as mz:
                    for entry in mz.infolist():
                        if entry.filename.lower().endswith(".mp3"):
                            music_files.append(entry.filename)
                            try:
                                from mutagen.mp3 import MP3

                                data = mz.read(entry)
                                audio = MP3(BytesIO(data))
                                music_durations[entry.filename] = audio.info.length
                            except Exception:
                                music_durations[entry.filename] = None
            except Exception as exc:
                self.log(f"Kunne ikke lese musikk-zip: {music_zip} ({exc})")
        else:
            self.log("Fant ingen musikk-zip (navn med 'MUSIKK').")

        zip_map = {}
        for row in zip_rows:
            key = (
                normalize_name(row.get("GivenName")),
                normalize_name(row.get("FamilyName")),
            )
            if key[0] or key[1]:
                zip_map[key] = row

        excel_keys = set()
        for row in self.rows:
            key = (
                normalize_name(row.get("GivenName")),
                normalize_name(row.get("FamilyName")),
            )
            excel_keys.add(key)
            zip_row = zip_map.get(key)
            if zip_row:
                row["ParticipantCode"] = zip_row.get("ParticipantCode", "")
                row["Event"] = zip_row.get("Event", "")
                row["EntryOrder"] = zip_row.get("EntryOrder", "")
                row["Music1"] = zip_row.get("Music1", "")
                row["Music2"] = zip_row.get("Music2", "")
                row["Club1"] = zip_row.get("Club1", "")
                row["Club2"] = zip_row.get("Club2", "")
                row["ElementsFree"] = zip_row.get("ElementsFree", "")
                row["ElementsShort"] = zip_row.get("ElementsShort", "")
                row["Manglende i zip"] = ""
            else:
                row["Manglende i zip"] = "JA"
                self.log(f"Mangler i zip: {row.get('PrintName')}")

            if music_files:
                matched = None
                for fname in music_files:
                    if name_matches_filename(
                        row.get("GivenName"),
                        row.get("FamilyName"),
                        fname,
                    ):
                        matched = fname
                        break
                row["Musikk"] = "ok" if matched else "mangler"
                row["MusikkTid"] = format_duration(
                    music_durations.get(matched) if matched else None
                )
                if not matched:
                    self.log(f"Mangler musikk: {row.get('PrintName')}")
            else:
                row["Musikk"] = "mangler"
                row["MusikkTid"] = ""

        for key, row in zip_map.items():
            if key not in excel_keys:
                self.log(f"Finnes i zip, men ikke i excel: {row.get('PrintName')}")

        if not self.rows:
            messagebox.showwarning("Info", "Fant ingen deltakere i xml.")
            self.set_output_controls(enabled=False)
            return

        self.log(f"Totalt deltakere: {len(self.rows)}")
        self.set_output_controls(enabled=True)

    def generate_files(self):
        if not self.rows or not self.zip_path:
            messagebox.showerror("Feil", "Ingen data lastet.")
            return

        out_dir = self.zip_path.parent / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        base_name = self.zip_path.stem

        if self.var_excel.get():
            generate_excel(self.rows, str(out_dir / f"{base_name}.xlsx"), self.log)
        if self.var_html.get():
            generate_html(
                self.rows, str(out_dir / f"{base_name}.html"), base_name, self.log
            )
        if self.var_pdf.get():
            generate_pdf(self.rows, str(out_dir / f"{base_name}.pdf"), base_name, self.log)

        self.log("Ferdig.")

    def generate_startliste(self):
        if not self.rows or not self.zip_path:
            messagebox.showerror("Feil", "Ingen data lastet.")
            return

        start_dt = parse_time_hhmm(self.start_time_var.get())
        if not start_dt:
            messagebox.showerror("Feil", "Ugyldig starttid. Bruk HH:MM.")
            return

        interval_seconds = parse_duration_mmss(self.interval_var.get())
        if not interval_seconds:
            messagebox.showerror("Feil", "Ugyldig intervall. Bruk M:SS eller H:MM:SS.")
            return

        try:
            group_size = int(self.group_size_var.get())
        except ValueError:
            messagebox.showerror("Feil", "Ugyldig gruppe-storrelse. Bruk et tall.")
            return
        if group_size <= 0:
            messagebox.showerror("Feil", "Gruppe-storrelse må være > 0.")
            return

        date_text = self.start_date_var.get().strip()
        location = self.location_var.get().strip() or "iskanten"
        start_time_text = start_dt.strftime("%H.%M")
        title = f"Startliste Oppvisning {date_text}, kl. {start_time_text}, {location}"

        filtered = [r for r in self.rows if is_registered(r.get("Påmelding", ""))]
        if not filtered:
            messagebox.showwarning("Info", "Fant ingen påmeldte i listen.")
            return
        if self.shuffle_var.get():
            random.shuffle(filtered)

        entries = build_startliste(filtered, group_size, interval_seconds, start_dt)
        out_dir = self.zip_path.parent / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        base_name = self.zip_path.stem
        generate_startliste_excel(
            entries, str(out_dir / f"Startliste_{base_name}.xlsx"), title, self.log
        )
        generate_startliste_pdf(
            entries, str(out_dir / f"Startliste_{base_name}.pdf"), title, self.log
        )
        self.log("Startliste ferdig.")

    def show_about(self):
        version = "ukjent"
        try:
            result = subprocess.run(
                ["git", "rev-parse", "--short", "HEAD"],
                capture_output=True,
                text=True,
                cwd=Path(__file__).resolve().parent,
            )
            if result.returncode == 0:
                version = result.stdout.strip() or "ukjent"
        except Exception:
            version = "ukjent"

        month_names = [
            "januar",
            "februar",
            "mars",
            "april",
            "mai",
            "juni",
            "juli",
            "august",
            "september",
            "oktober",
            "november",
            "desember",
        ]
        now = datetime.now()
        month_year = f"{month_names[now.month - 1]} {now.year}"

        about = tk.Toplevel(self.root)
        about.title("Om")
        about.resizable(False, False)

        frame = ttk.Frame(about, padding=16)
        frame.pack(fill="both", expand=True)

        logo_path = Path(__file__).resolve().parent / "Lil Logo.jpg"
        logo_label = None
        if logo_path.exists():
            try:
                from PIL import Image, ImageTk

                img = Image.open(logo_path)
                img.thumbnail((180, 180))
                photo = ImageTk.PhotoImage(img)
                logo_label = ttk.Label(frame, image=photo)
                logo_label.image = photo
                logo_label.pack(pady=(0, 8))
            except Exception:
                logo_label = None

        ttk.Label(
            frame,
            text="Loddefjord IL Kunstløp",
            font=("Segoe UI", 12, "bold"),
        ).pack()
        ttk.Label(
            frame,
            text=(
                "Programmet leser XML fra FSM-zip og lager "
                "Excel/HTML/PDF-utskrifter."
            ),
            wraplength=360,
            justify="center",
        ).pack(pady=(6, 2))
        ttk.Label(
            frame,
            text=f"Revisjon: {version} • {month_year}",
        ).pack(pady=(4, 0))


def main():
    root = tk.Tk()
    app = App(root)
    root.geometry("900x600")
    root.mainloop()


if __name__ == "__main__":
    main()

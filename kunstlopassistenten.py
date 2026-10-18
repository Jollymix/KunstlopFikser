import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from tkinter.scrolledtext import ScrolledText
from datetime import datetime, timedelta
import subprocess
import tempfile
from io import BytesIO
import hashlib
import random
import re
import json
import shutil
import wave
from tkinter import font as tkfont

APP_NAME = "Kunstløpassistenten"

# Farger: mørk skifer med isblå aksent (samme oppsett som pc_debugger, annen farge)
BG = "#0F172A"            # vindusbakgrunn
PANEL = "#1B2336"         # paneler
FIELD = "#0B1222"         # felt, tabeller og logg
LINE = "#334155"          # kanter
TEXT = "#F8FAFC"          # hovedtekst
TEXT_DIM = "#94A3B8"      # sekundær tekst
TEXT_DISABLED = "#5B6B82"
BUTTON_HOVER = "#243049"
ACCENT = "#38BDF8"        # hovedknapper og markering
ACCENT_HOVER = "#7DD3FC"
ACCENT_TEXT = "#0F172A"
GROUP_BG = "#26324A"      # oppvarmingsgrupper i tidsplanen
WARN = "#F2B33D"
GOOD = "#86D9A3"
BAD = "#F87171"

SETUP_FILETYPES = [("Oppsett", "*.json"), ("Alle filer", "*.*")]


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
    roots = []
    try:
        roots = [ET.fromstring(xml_text)]
    except ET.ParseError as exc:
        if "junk after document element" in str(exc):
            log("XML inneholder flere dokumenter, prøver å hente alle OdfBody.")
            for match in re.finditer(r"<OdfBody\b", xml_text):
                start = match.start()
                end = xml_text.find("</OdfBody>", start)
                if end != -1:
                    chunk = xml_text[start : end + len("</OdfBody>")]
                    try:
                        roots.append(ET.fromstring(chunk))
                    except ET.ParseError:
                        continue
        else:
            raise

    for idx_root, root in enumerate(roots, start=1):
        log(f"Leser OdfBody {idx_root} av {len(roots)}.")
        comp = root.find("Competition")
        if comp is None:
            continue

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
    text = (value or "").strip().lower()
    text = (
        text.replace("ø", "o")
        .replace("å", "a")
        .replace("æ", "ae")
        .replace("ö", "o")
        .replace("ä", "a")
        .replace("é", "e")
    )
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def tokenize_name(value):
    norm = normalize_text(value)
    return [t for t in norm.split() if t]

def build_name_key(given, family, strict=True):
    given_tokens = tokenize_name(given)
    family_tokens = tokenize_name(family)
    if not given_tokens and not family_tokens:
        return ("", "")
    if strict:
        return (normalize_name(given), normalize_name(family))
    given_primary = given_tokens[0] if given_tokens else ""
    family_norm = " ".join(family_tokens)
    return (given_primary, family_norm)


def sanitize_filename(value):
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", (value or "").strip())
    return cleaned.strip("._") or "fil"


def detect_audio_format(data):
    if data[:4] == b"RIFF" and data[8:12] == b"WAVE":
        return "WAV", ".wav"
    if len(data) >= 12 and data[4:8] == b"ftyp":
        return "M4A", ".m4a"
    if data[:3] == b"ID3":
        return "MP3", ".mp3"
    if len(data) >= 2 and data[0] == 0xFF and (data[1] & 0xE0) == 0xE0:
        version_bits = (data[1] >> 3) & 0x03
        layer_bits = (data[1] >> 1) & 0x03
        if version_bits != 0x01 and layer_bits != 0x00:
            return "MP3", ".mp3"
    return "Ukjent", ""


def get_audio_duration(data, audio_format):
    try:
        if audio_format == "WAV":
            with wave.open(BytesIO(data), "rb") as audio:
                frame_rate = audio.getframerate()
                return audio.getnframes() / frame_rate if frame_rate else None
        if audio_format == "MP3":
            from mutagen.mp3 import MP3

            return MP3(BytesIO(data)).info.length
        if audio_format == "M4A":
            from mutagen.mp4 import MP4

            return MP4(BytesIO(data)).info.length
    except Exception:
        return None
    return None


def find_vlc_executable():
    candidates = [shutil.which("vlc")]
    for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
        base_dir = os.environ.get(env_name)
        if base_dir:
            candidates.append(Path(base_dir) / "VideoLAN" / "VLC" / "vlc.exe")
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(candidate)
    return None


def get_version():
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
    if version == "ukjent":
        try:
            version_file = Path(__file__).resolve().parent / "version.txt"
            if version_file.exists():
                version = version_file.read_text(encoding="utf-8").strip() or "ukjent"
        except Exception:
            pass
    if version == "ukjent":
        try:
            mtime = Path(__file__).resolve().stat().st_mtime
            version = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d")
        except Exception:
            pass
    return version


def format_generated_ts():
    computer_name = os.environ.get("COMPUTERNAME") or os.environ.get("HOSTNAME") or ""
    suffix = f" • {computer_name}" if computer_name else ""
    version = get_version()
    return (
        datetime.now().strftime("Generert %d.%m.%Y %H:%M")
        + suffix
        + f" • Revisjon: {version}"
    )


def name_matches_filename(given, family, filename):
    family_tokens = tokenize_name(family)
    hay = normalize_text(filename)
    if not family_tokens:
        return False
    if all(token in hay for token in family_tokens):
        return True
    return family_tokens[0] in hay


def format_duration(seconds):
    if seconds is None:
        return ""
    total = int(round(seconds))
    mins = total // 60
    secs = total % 60
    return f"{mins}:{secs:02d}"


def format_hms(seconds):
    total = int(seconds)
    return f"{total // 3600:02d}:{total % 3600 // 60:02d}:{total % 60:02d}"


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


def parse_date_ddmmyy(value):
    try:
        return datetime.strptime(value.strip(), "%d.%m.%y").date()
    except Exception:
        return None


def format_date_long(date_obj):
    months = [
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
    return f"{date_obj.day} {months[date_obj.month - 1]} {date_obj.year}"


def is_registered(status):
    if is_cancelled(status):
        return False
    if not status:
        return False
    text = str(status).strip().lower()
    if "ikke sjekket inn" in text:
        return True
    return "påmeld" in text or "registr" in text or "bekreftet" in text


def new_participant(given, family, club="", gender="", status=""):
    given = str(given or "").strip()
    family = str(family or "").strip()
    return {
        "PrintName": f"{family}, {given}".strip(", "),
        "NavnFraIsonen": f"{given} {family}".strip(),
        "NavnFraFsm": "",
        "GivenName": given,
        "FamilyName": family,
        "Gender": str(gender or "").strip(),
        "Organisation": str(club or "").strip(),
        "ParticipantCode": "",
        "Event": "",
        "EntryOrder": "",
        "Påmelding": str(status or "").strip(),
        "Music1": "",
        "Music2": "",
        "Club1": "",
        "Club2": "",
        "ElementsFree": "",
        "ElementsShort": "",
        "Manglende i zip": "",
        "Musikk": "",
        "MusikkFil": "",
        "MusikkFormat": "",
        "MusikkFormatAvvik": False,
        "MusikkSpillbar": False,
        "MusikkTid": "",
        "MusikkSek": "",
        "StartTid": "",
        "SluttTid": "",
    }


def split_name(full_name):
    parts = (full_name or "").split()
    if len(parts) < 2:
        return " ".join(parts), ""
    return " ".join(parts[:-1]), parts[-1]


def display_name(row):
    full = f"{row.get('GivenName', '')} {row.get('FamilyName', '')}".strip()
    return full or row.get("NavnFraIsonen") or row.get("PrintName") or ""


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

        out.append(new_participant(given, family, club, gender, status))

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
        "StartTid",
        "SluttTid",
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
        "StartTid",
        "SluttTid",
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
        "StartTid",
        "SluttTid",
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
    story.append(Spacer(1, 6))
    story.append(Paragraph(format_generated_ts(), styles["Normal"]))
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


def water_pause_positions(group_count, water_count):
    # Fordeler vanningspausene jevnt mellom oppvarmingsgruppene (etter gruppe nr. N).
    positions = {round(group_count * k / (water_count + 1)) for k in range(1, water_count + 1)}
    return sorted(p for p in positions if 0 < p < group_count)


def is_marker(item):
    return item.get("Kind") in ("warmup", "pause")


def place_water_pauses(items, water_count):
    # Fjerner gamle vanningspauser og legger nye jevnt fordelt foran oppvarmingsgruppene.
    items = [item for item in items if item.get("Kind") != "pause"]
    warmups = [k for k, item in enumerate(items) if item.get("Kind") == "warmup"]
    for position in reversed(water_pause_positions(len(warmups), water_count)):
        items.insert(warmups[position], {"Kind": "pause"})
    return items


def layout_items(rows, group_size, water_count):
    # Lager rekkefølgen oppvarmingsgruppe, utøvere, ... Avmeldte legges skjult bakerst.
    group_size = max(1, group_size)
    active = [r for r in rows if not is_cancelled(r.get("Påmelding", ""))]
    items = []
    for start in range(0, len(active), group_size):
        items.append({"Kind": "warmup"})
        items.extend(active[start : start + group_size])
    items.extend(r for r in rows if is_cancelled(r.get("Påmelding", "")))
    return place_water_pauses(items, water_count)


def group_overflows(items, group_size):
    # {id(oppvarmingsrad): antall utøvere} for grupper med flere utøvere enn tillatt.
    counts = {}
    current = None
    for item in items:
        if item.get("Kind") == "warmup":
            current = id(item)
            counts[current] = 0
        elif current is not None and not is_marker(item) and not is_cancelled(item.get("Påmelding", "")):
            counts[current] += 1
    return {key: count for key, count in counts.items() if count > group_size}


def build_startliste(
    items,
    run_seconds,
    start_time,
    ready_seconds=0,
    warmup_seconds=0,
    water_seconds=0,
    water_label="Vanningspause",
):
    entries = []
    current = start_time
    group = 0
    nr = 0

    def add(kind, seconds, nr, navn, klubb, row):
        nonlocal current
        end = current + timedelta(seconds=seconds)
        entries.append(
            {
                "is_group": kind != "run",
                "kind": kind,
                "start": current.strftime("%H:%M:%S"),
                "end": end.strftime("%H:%M:%S"),
                "duration": format_hms(seconds),
                "nr": nr,
                "navn": navn,
                "klubb": klubb,
                "row": row,
            }
        )
        current = end

    for item in items:
        kind = item.get("Kind")
        if kind == "warmup":
            group += 1
            add("warmup", warmup_seconds, "", f"Oppvarmingsgruppe {group}", "", item)
        elif kind == "pause":
            add("pause", water_seconds, "", water_label, "", item)
        elif not is_cancelled(item.get("Påmelding", "")):
            nr += 1
            add("run", run_seconds + ready_seconds, nr, display_name(item), item.get("Organisation", ""), item)
    return entries


def generate_startliste_excel(entries, out_path, title, log):
    try:
        import openpyxl
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Font, PatternFill
    except Exception:
        log("Mangler openpyxl. Installer med: pip install openpyxl")
        return False

    header_fill = PatternFill(start_color="ADADAD", end_color="ADADAD", fill_type="solid")
    group_fill = PatternFill(start_color="D0D0D0", end_color="D0D0D0", fill_type="solid")
    bold = Font(bold=True)
    generated_ts = format_generated_ts()

    def write_sheet(ws, with_duration):
        ws.append([title])
        ws.append([])
        headers = ["Nr.", "Start", "", "Slutt"] + (["Tidsforbruk"] if with_duration else []) + ["Navn", "Klubb"]
        ws.append(headers)
        for cell in ws[3]:
            cell.font = bold
            cell.fill = header_fill

        group = 0
        for entry in entries:
            start = entry["start"]
            if entry["kind"] == "warmup":
                group += 1
                if group > 1:
                    start = f"ca. {start}"
            duration = [entry["duration"]] if with_duration else []
            ws.append([entry["nr"], start, "-", entry["end"]] + duration + [entry["navn"], entry["klubb"]])
            if entry["is_group"]:
                for cell in ws[ws.max_row]:
                    cell.font = bold
                    cell.fill = group_fill

        widths = [6, 10, 3, 10] + ([12] if with_duration else []) + [38, 16]
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = width
        ws.append([])
        ws.append([generated_ts])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Startliste"
    write_sheet(ws, with_duration=False)
    write_sheet(wb.create_sheet("Med tidsforbruk"), with_duration=True)

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
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except Exception:
        log("Mangler reportlab. Installer med: pip install reportlab")
        return False

    font_name = "Helvetica"
    try:
        calibri_path = Path("C:/Windows/Fonts/calibri.ttf")
        if calibri_path.exists():
            pdfmetrics.registerFont(TTFont("Calibri", str(calibri_path)))
            font_name = "Calibri"
    except Exception:
        font_name = "Helvetica"

    data = [["Nr.", "Start", "", "Slutt", "Tidsforbruk", "Navn", "Klubb"]]
    body_style = ParagraphStyle(
        "BodyCell",
        fontName=font_name,
        fontSize=10,
        leading=11,
    )
    for entry in entries:
        name_cell = Paragraph(entry["navn"], body_style)
        club_cell = Paragraph(entry["klubb"], body_style)
        data.append([entry["nr"], entry["start"], "-", entry["end"], entry["duration"], name_cell, club_cell])

    doc = SimpleDocTemplate(out_path, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "StartTitle",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=16,
        leading=18,
    )
    generated_ts = format_generated_ts()
    story = [Paragraph(title, title_style), Spacer(1, 8)]

    col_widths = [26, 56, 10, 56, 62, 190, 123]
    table = Table(data, repeatRows=1, colWidths=col_widths)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ADADAD")),
        ("FONTNAME", (0, 0), (-1, 0), font_name),
        ("FONTNAME", (0, 1), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, 0), 11),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("ALIGN", (0, 0), (4, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LINEABOVE", (0, 0), (-1, 0), 0.5, colors.black),
        ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.black),
        ("LINEBEFORE", (0, 0), (0, -1), 0.5, colors.black),
        ("LINEAFTER", (-1, 0), (-1, -1), 0.5, colors.black),
    ]
    row_idx = 1
    for entry in entries:
        if entry["is_group"]:
            style_cmds.append(
                ("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor("#D0D0D0"))
            )
            style_cmds.append(("FONTNAME", (0, row_idx), (-1, row_idx), "Helvetica-Bold"))
        row_idx += 1
    table.setStyle(TableStyle(style_cmds))
    story.append(table)
    story.append(Spacer(1, 6))
    story.append(Paragraph(generated_ts, ParagraphStyle("Gen", fontName=font_name, fontSize=9)))

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


def generate_vlc_playlist(rows, out_dir, base_name, music_zip, log):
    if not music_zip:
        log("Ingen musikk-zip funnet, kan ikke lage spilleliste.")
        return False

    playlist_path = Path(out_dir) / f"Startliste_{base_name}.m3u"
    music_out = Path(out_dir) / "music"
    music_out.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(music_zip, "r") as mz:
            lines = ["#EXTM3U"]
            for row in rows:
                fname = row.get("MusikkFil", "")
                if not fname:
                    continue
                dest_path = music_out / fname
                dest_path.parent.mkdir(parents=True, exist_ok=True)
                if not dest_path.exists():
                    try:
                        mz.extract(fname, music_out)
                    except Exception:
                        continue
                duration = row.get("MusikkSek", "")
                performer = row.get("PrintName", "")
                song = Path(fname).stem
                title = f"{performer} - {song}".strip(" -")
                extinf = duration if duration != "" else -1
                lines.append(f"#EXTINF:{extinf},{title}")
                lines.append(str(dest_path.resolve()))
    except Exception as exc:
        log(f"Kunne ikke lage spilleliste: {exc}")
        return False

    playlist_path.write_text("\n".join(lines), encoding="utf-8")
    log(f"Spilleliste skrevet: {playlist_path}")
    return True


class App:
    def __init__(self, root):
        self.root = root
        # Hele tidsplanen i rekkefølge: utøvere (deltaker-dict) og markører for
        # oppvarmingsgrupper og vanningspauser ({"Kind": "warmup"/"pause"}).
        self.items = []
        self.entries = []
        self.setup_path = None
        self.dirty = False
        self._loading = False
        self.drag_iid = None
        self.overflow_tips = {}
        self.tip = None
        self.zip_path = None
        self.source_dir = None
        self.music_zip = None
        self.music_cache_dir = None
        self.music_file_info = {}
        self.music_files = []
        self.music_durations = {}
        self.audio_backend = None
        self.audio_ready = False
        self.vlc_process = None
        self.current_track = ""
        self.is_paused = False
        self.current_duration = 0
        self.external_playback = False
        self.play_pause_text = tk.StringVar(value="Spill")
        self.use_external_player_var = tk.BooleanVar(value=False)

        self._init_fonts()
        self._init_style()
        self._init_vars()
        self._build_menu()
        self._build()
        self.refresh_schedule()
        self.dirty = False
        self.update_title()
        self.update_player_ui()
        self.update_clock()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    @property
    def rows(self):
        return [item for item in self.items if not is_marker(item)]

    @rows.setter
    def rows(self, value):
        self.items = list(value)

    # -------------------- Oppsett --------------------
    def _init_fonts(self):
        fams = set(tkfont.families())
        base = "Bahnschrift" if "Bahnschrift" in fams else "Segoe UI"
        cond = "Bahnschrift SemiCondensed" if "Bahnschrift SemiCondensed" in fams else base
        semi = "Bahnschrift SemiBold" if "Bahnschrift SemiBold" in fams else base
        mono = "Consolas" if "Consolas" in fams else "Courier New"
        self.fonts = {
            "body": (base, 10),
            "bold": (semi, 10),
            "small": (cond, 10),
            "head": (semi, 13),
            "title": (semi, 16),
            "clock": (semi, 96),
            "mono": (mono, 10),
        }

    def _init_style(self):
        r = self.root
        f = self.fonts
        r.configure(bg=BG)
        r.option_add("*TCombobox*Listbox.background", FIELD)
        r.option_add("*TCombobox*Listbox.foreground", TEXT)
        r.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
        r.option_add("*TCombobox*Listbox.selectForeground", ACCENT_TEXT)
        r.option_add("*TCombobox*Listbox.font", f["body"])

        s = ttk.Style(r)
        s.theme_use("clam")
        s.configure(".", background=BG, foreground=TEXT, font=f["body"], fieldbackground=FIELD,
                    bordercolor=LINE, lightcolor=PANEL, darkcolor=BG, troughcolor=FIELD,
                    selectbackground=ACCENT, selectforeground=ACCENT_TEXT, insertcolor=TEXT,
                    focuscolor=ACCENT)
        s.configure("TFrame", background=BG)
        s.configure("Panel.TFrame", background=PANEL)
        s.configure("TLabel", background=BG, foreground=TEXT)
        s.configure("Title.TLabel", font=f["title"])
        s.configure("Dim.TLabel", foreground=TEXT_DIM, font=f["small"])
        s.configure("Panel.TLabel", background=PANEL)
        s.configure("PanelDim.TLabel", background=PANEL, foreground=TEXT_DIM, font=f["small"])
        s.configure("PanelHead.TLabel", background=PANEL, font=f["head"])
        s.configure("PanelWarn.TLabel", background=PANEL, foreground=WARN)

        s.configure("TButton", background=FIELD, foreground=TEXT, bordercolor=LINE,
                    padding=(12, 6), relief="flat")
        s.map("TButton", background=[("disabled", PANEL), ("active", BUTTON_HOVER)],
              foreground=[("disabled", TEXT_DISABLED)])
        s.configure("Accent.TButton", background=ACCENT, foreground=ACCENT_TEXT, bordercolor=ACCENT,
                    font=f["head"])
        s.map("Accent.TButton", background=[("disabled", "#3B5568"), ("active", ACCENT_HOVER)],
              foreground=[("disabled", PANEL)])

        s.configure("TEntry", foreground=TEXT, fieldbackground=FIELD, padding=5)
        s.configure("TSpinbox", foreground=TEXT, fieldbackground=FIELD, arrowcolor=TEXT, padding=4)
        s.configure("TCombobox", foreground=TEXT, fieldbackground=FIELD, arrowcolor=TEXT, padding=4)
        s.configure("TCheckbutton", background=PANEL, foreground=TEXT, indicatorbackground=FIELD,
                    indicatorforeground=ACCENT)
        s.map("TCheckbutton", background=[("active", PANEL)], indicatorbackground=[("selected", FIELD)])

        s.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(0, 0, 0, 0))
        s.configure("TNotebook.Tab", background=BG, foreground=TEXT_DIM, padding=(16, 7),
                    bordercolor=BG, font=f["body"])
        s.map("TNotebook.Tab", background=[("selected", PANEL)], foreground=[("selected", TEXT)])

        s.configure("Treeview", background=FIELD, fieldbackground=FIELD, foreground=TEXT,
                    rowheight=26, bordercolor=LINE, font=f["body"])
        s.map("Treeview", background=[("selected", ACCENT)], foreground=[("selected", ACCENT_TEXT)])
        s.configure("Treeview.Heading", background=PANEL, foreground=TEXT_DIM, font=f["small"],
                    relief="flat", bordercolor=LINE, padding=(6, 4))
        s.map("Treeview.Heading", background=[("active", BUTTON_HOVER)])
        s.configure("Horizontal.TProgressbar", background=ACCENT, troughcolor=FIELD, bordercolor=LINE,
                    lightcolor=ACCENT, darkcolor=ACCENT)
        s.configure("Vertical.TScrollbar", background=PANEL, troughcolor=FIELD, bordercolor=LINE,
                    arrowcolor=TEXT_DIM)

    def _init_vars(self):
        self.event_name_var = tk.StringVar(value="Oppvisningsstevne")
        self.location_var = tk.StringVar(value="Iskanten")
        self.start_date_var = tk.StringVar(value=datetime.now().strftime("%d.%m.%y"))
        self.start_time_var = tk.StringVar(value="18:00")
        self.interval_var = tk.StringVar(value="3:10")
        self.ready_var = tk.StringVar(value="0:30")
        self.group_size_var = tk.StringVar(value="8")
        self.warmup_var = tk.StringVar(value="4:00")
        self.water_count_var = tk.StringVar(value="0")
        self.water_duration_var = tk.StringVar(value="15:00")
        # Feltene som lagres i oppsettfilen.
        self.setup_vars = {
            "stevne": self.event_name_var,
            "sted": self.location_var,
            "dato": self.start_date_var,
            "start": self.start_time_var,
            "lengde_lop": self.interval_var,
            "klart_til_neste": self.ready_var,
            "utovere_per_gruppe": self.group_size_var,
            "oppvarming": self.warmup_var,
            "antall_vanningspauser": self.water_count_var,
            "lengde_vanningspause": self.water_duration_var,
        }
        for var in (self.event_name_var, self.location_var, self.start_date_var):
            var.trace_add("write", lambda *_: self.mark_dirty())
        for var in (
            self.start_time_var,
            self.interval_var,
            self.ready_var,
            self.warmup_var,
            self.water_duration_var,
        ):
            var.trace_add("write", lambda *_: self.refresh_schedule())
        self.group_size_var.trace_add("write", lambda *_: self.on_layout_setting_change(regroup=True))
        self.water_count_var.trace_add("write", lambda *_: self.on_layout_setting_change(regroup=False))
        self.new_name_var = tk.StringVar()
        self.new_club_var = tk.StringVar()
        self.var_pdf = tk.BooleanVar(value=False)
        self.var_excel = tk.BooleanVar(value=True)
        self.var_html = tk.BooleanVar(value=True)

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu_file = tk.Menu(menubar, tearoff=0)
        menu_file.add_command(label="Hent deltakerliste fra isonen…", command=self.load_isonen)
        menu_file.add_command(label="Velg musikk-zip…", command=self.choose_music_zip)
        menu_file.add_command(
            label="Skann mappe (isonen, FMS-data og musikk)…", command=self.scan_folder
        )
        menu_file.add_separator()
        menu_file.add_command(label="Nytt oppsett", command=self.clear_rows)
        menu_file.add_command(label="Åpne oppsett…", accelerator="Ctrl+O", command=self.open_setup)
        menu_file.add_command(label="Lagre oppsett", accelerator="Ctrl+S", command=self.save_setup)
        menu_file.add_command(label="Lagre oppsett som…", command=lambda: self.save_setup(ask=True))
        menu_file.add_separator()
        menu_file.add_command(label="Avslutt", command=self.on_close)
        menu_reports = tk.Menu(menubar, tearoff=0)
        menu_reports.add_command(label="Lag deltakerrapport", command=self.generate_files)
        menu_reports.add_checkbutton(label="PDF", variable=self.var_pdf)
        menu_reports.add_checkbutton(label="Excel", variable=self.var_excel)
        menu_reports.add_checkbutton(label="HTML", variable=self.var_html)
        menu_settings = tk.Menu(menubar, tearoff=0)
        self.menu_tabellvisning = tk.Menu(menu_settings, tearoff=0)
        menu_settings.add_cascade(label="Kolonner i stevnetabellen", menu=self.menu_tabellvisning)
        menu_help = tk.Menu(menubar, tearoff=0)
        menu_help.add_command(label="Om", command=self.show_about)
        menubar.add_cascade(label="Fil", menu=menu_file)
        menubar.add_cascade(label="Rapporter", menu=menu_reports)
        menubar.add_cascade(label="Innstillinger", menu=menu_settings)
        menubar.add_cascade(label="Hjelp", menu=menu_help)
        self.root.bind_all("<Control-s>", lambda e: self.save_setup())
        self.root.bind_all("<Control-o>", lambda e: self.open_setup())

    def _build(self):
        r = self.root
        r.title(APP_NAME)
        r.minsize(1100, 700)

        top = ttk.Frame(r, padding=(18, 12, 18, 6))
        top.pack(fill="x")
        title_box = ttk.Frame(top)
        title_box.pack(side="left", anchor="w")
        ttk.Label(title_box, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        self.status_var = tk.StringVar(
            value="Hent en deltakerliste fra isonen, eller legg inn utøvere manuelt."
        )
        ttk.Label(title_box, textvariable=self.status_var, style="Dim.TLabel").pack(anchor="w")

        log_panel = ttk.Frame(r, style="Panel.TFrame", padding=(14, 8))
        log_panel.pack(side="bottom", fill="x", padx=18, pady=(0, 18))
        ttk.Label(log_panel, text="Logg", style="PanelHead.TLabel").pack(anchor="w")
        log_frame = ttk.Frame(log_panel, style="Panel.TFrame")
        log_frame.pack(fill="x", pady=(4, 0))
        self.log_widget = tk.Text(log_frame, height=5, bg=FIELD, fg=TEXT_DIM, insertbackground=TEXT,
                                  font=self.fonts["mono"], relief="flat", wrap="word", padx=10,
                                  pady=6, highlightthickness=0)
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_widget.yview)
        self.log_widget.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side="right", fill="y")
        self.log_widget.pack(side="left", fill="both", expand=True)

        body = ttk.Frame(r, padding=(18, 4, 18, 10))
        body.pack(fill="both", expand=True)
        nb = ttk.Notebook(body)
        nb.pack(fill="both", expand=True)
        self._tab_schedule(nb)
        self._tab_music(nb)

    def _panel(self, nb, title):
        frame = ttk.Frame(nb, style="Panel.TFrame", padding=20)
        nb.add(frame, text=title)
        return frame

    # -------------------- Fane: Tidsplan --------------------
    def _tab_schedule(self, nb):
        p = self._panel(nb, "Tidsplan/planlegging")
        p.columnconfigure(1, weight=1)
        p.rowconfigure(0, weight=1)

        form = ttk.Frame(p, style="Panel.TFrame")
        form.grid(row=0, column=0, sticky="ns", padx=(0, 24))
        ttk.Label(form, text="Oppsett", style="PanelHead.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        fields = [
            ("Stevne", self.event_name_var, 22, None),
            ("Sted", self.location_var, 22, None),
            ("Dato", self.start_date_var, 10, "date"),
            ("Start kl.", self.start_time_var, 8, None),
            ("Lengde på løp", self.interval_var, 8, None),
            ("Klart til neste løp", self.ready_var, 8, None),
            ("Utøvere per gruppe", self.group_size_var, 6, (1, 40)),
            ("Oppvarming per gruppe", self.warmup_var, 8, None),
            ("Antall vanningspauser", self.water_count_var, 6, (0, 20)),
            ("Lengde på vanningspause", self.water_duration_var, 8, None),
        ]
        for i, (label, var, width, kind) in enumerate(fields, start=1):
            ttk.Label(form, text=label, style="Panel.TLabel").grid(
                row=i, column=0, sticky="w", pady=3, padx=(0, 12))
            widget = None
            if kind == "date":
                try:
                    from tkcalendar import DateEntry

                    widget = DateEntry(form, textvariable=var, date_pattern="dd.mm.yy", width=width)
                except Exception:
                    widget = None
            elif kind:
                widget = ttk.Spinbox(form, from_=kind[0], to=kind[1], textvariable=var, width=width,
                                     font=self.fonts["body"])
            if widget is None:
                widget = ttk.Entry(form, textvariable=var, width=width, font=self.fonts["body"])
            widget.grid(row=i, column=1, sticky="w", pady=3)
        row = len(fields) + 1
        ttk.Label(
            form,
            text="Tider skrives som M:SS, starttid som TT:MM.\n"
                 "Lengde på løp + klart til neste løp = tid per utøver.",
            style="PanelDim.TLabel",
            justify="left",
        ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(4, 0))

        ttk.Label(form, text="Legg til utøver", style="PanelHead.TLabel").grid(
            row=row + 1, column=0, columnspan=2, sticky="w", pady=(22, 8))
        ttk.Label(form, text="Navn", style="Panel.TLabel").grid(row=row + 2, column=0, sticky="w", pady=3)
        name_entry = ttk.Entry(form, textvariable=self.new_name_var, width=22, font=self.fonts["body"])
        name_entry.grid(row=row + 2, column=1, sticky="w", pady=3)
        ttk.Label(form, text="Klubb", style="Panel.TLabel").grid(row=row + 3, column=0, sticky="w", pady=3)
        club_entry = ttk.Entry(form, textvariable=self.new_club_var, width=22, font=self.fonts["body"])
        club_entry.grid(row=row + 3, column=1, sticky="w", pady=3)
        for entry in (name_entry, club_entry):
            entry.bind("<Return>", lambda e: self.add_runner())
        self.name_entry = name_entry
        ttk.Button(form, text="Legg til", command=self.add_runner).grid(
            row=row + 4, column=1, sticky="w", pady=(6, 0))

        actions = ttk.Frame(form, style="Panel.TFrame")
        actions.grid(row=row + 5, column=0, columnspan=2, sticky="we", pady=(24, 0))
        ttk.Button(actions, text="Hent fra isonen…", command=self.load_isonen).pack(fill="x")
        ttk.Button(actions, text="Åpne oppsett…", command=self.open_setup).pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="Lagre oppsett", command=self.save_setup).pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="Lagre startliste (Excel + PDF)", style="Accent.TButton",
                   command=self.generate_startliste).pack(fill="x", pady=(16, 0))

        table = ttk.Frame(p, style="Panel.TFrame")
        table.grid(row=0, column=1, sticky="nsew")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        columns = {
            "nr": ("Nr.", 50, "center"),
            "start": ("Start", 90, "center"),
            "dash": ("", 24, "center"),
            "slutt": ("Slutt", 90, "center"),
            "tid": ("Tidsforbruk", 100, "center"),
            "navn": ("Navn", 260, "w"),
            "klubb": ("Klubb", 200, "w"),
        }
        self.schedule_tree = ttk.Treeview(table, columns=tuple(columns), show="headings",
                                          selectmode="browse")
        for col, (text, width, anchor) in columns.items():
            self.schedule_tree.heading(col, text=text, anchor=anchor)
            self.schedule_tree.column(col, width=width, anchor=anchor, stretch=col in ("navn", "klubb"))
        self.schedule_tree.tag_configure("warmup", background=GROUP_BG, font=self.fonts["bold"])
        self.schedule_tree.tag_configure("pause", background=GROUP_BG, foreground=WARN,
                                         font=self.fonts["bold"])
        self.schedule_tree.tag_configure("overfull", background=WARN, foreground=ACCENT_TEXT,
                                         font=self.fonts["bold"])
        self.schedule_tree.bind("<Motion>", self.on_schedule_hover)
        self.schedule_tree.bind("<Leave>", lambda e: self.hide_tip())
        self.schedule_tree.bind("<Double-1>", lambda e: self.edit_selected())
        self.schedule_tree.bind("<Delete>", lambda e: self.delete_selected())
        self.schedule_tree.bind("<Alt-Up>", lambda e: self.move_selected(-1) or "break")
        self.schedule_tree.bind("<Alt-Down>", lambda e: self.move_selected(1) or "break")
        self.schedule_tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.schedule_tree.bind("<B1-Motion>", self.on_drag_motion)
        self.schedule_tree.bind("<ButtonRelease-1>", self.on_drag_drop)
        self.drop_line = tk.Frame(self.schedule_tree, bg=ACCENT, height=3)
        scroll = ttk.Scrollbar(table, orient="vertical", command=self.schedule_tree.yview)
        self.schedule_tree.configure(yscrollcommand=scroll.set)
        self.schedule_tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        self.summary_var = tk.StringVar()
        self.summary_label = ttk.Label(table, textvariable=self.summary_var, style="PanelDim.TLabel")
        self.summary_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        side = ttk.Frame(table, style="Panel.TFrame")
        side.grid(row=0, column=2, sticky="ns", padx=(12, 0))
        for text, command, gap in (
            ("Flytt opp", lambda: self.move_selected(-1), 0),
            ("Flytt ned", lambda: self.move_selected(1), 4),
            ("Rediger", self.edit_selected, 16),
            ("Slett", self.delete_selected, 4),
            ("Randomiser", self.shuffle_rows, 16),
            ("Sorter fornavn", self.sort_by_given, 4),
            ("Sorter etternavn", self.sort_by_family, 4),
            ("Eksporter til Excel", self.export_excel, 16),
        ):
            ttk.Button(side, text=text, command=command, width=16).pack(fill="x", pady=(gap, 0))
        ttk.Label(side, text="Dra rader med musen,\neller bruk Alt+↑/↓.\nDobbeltklikk redigerer.",
                  style="PanelDim.TLabel", justify="left").pack(anchor="w", pady=(16, 0))

    # -------------------- Fane: Stevne --------------------
    def _tab_music(self, nb):
        p = self._panel(nb, "Stevne")
        p.columnconfigure(0, weight=1)
        p.rowconfigure(2, weight=1)

        # Bahnschrift har ulik bredde på sifrene, så hvert tegn får en fast plass for å unngå at klokka hopper.
        clock_font = tkfont.Font(font=self.fonts["clock"])
        digit_w = max(clock_font.measure(d) for d in "0123456789")
        colon_w = clock_font.measure(":")
        widths = [digit_w, digit_w, colon_w, digit_w, digit_w, colon_w, digit_w, digit_w]
        self.clock_canvas = tk.Canvas(p, width=sum(widths), height=clock_font.metrics("linespace"),
                                      bg=PANEL, highlightthickness=0)
        self.clock_canvas.grid(row=0, column=0, pady=(0, 12))
        self.clock_items = []
        x = 0
        for width in widths:
            self.clock_items.append(self.clock_canvas.create_text(
                x + width / 2, 0, anchor="n", text="0", fill=TEXT, font=clock_font))
            x += width

        bar = ttk.Frame(p, style="Panel.TFrame")
        bar.grid(row=1, column=0, sticky="we", pady=(0, 12))
        ttk.Button(bar, text="Velg musikk-zip…", style="Accent.TButton",
                   command=self.choose_music_zip).pack(side="left")
        ttk.Button(bar, text="Lag spilleliste", command=self.generate_playlist_only).pack(
            side="left", padx=(8, 0))
        ttk.Checkbutton(bar, text="Bruk VLC direkte", variable=self.use_external_player_var,
                        command=self.on_use_external_player_toggle).pack(side="left", padx=(16, 0))
        self.music_info_var = tk.StringVar(value="Ingen musikk-zip valgt.")
        ttk.Label(bar, textvariable=self.music_info_var, style="PanelDim.TLabel").pack(
            side="left", padx=(16, 0))

        table = ttk.Frame(p, style="Panel.TFrame")
        table.grid(row=2, column=0, sticky="nsew")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        columns = {
            "startnummer": "Nr.",
            "lop_gjennomfort": "Løp gjennomført",
            "starttid": "Start",
            "navn": "Navn",
            "klubb": "Klubb",
            "påmelding": "Påmelding",
            "musikknavn": "Lydfil",
            "musikkformat": "Format",
            "musikktid": "Lengde",
        }
        self.tree = ttk.Treeview(table, columns=tuple(columns), show="headings", selectmode="browse")
        for col, text in columns.items():
            anchor = "w" if col in ("navn", "klubb", "påmelding", "musikknavn") else "center"
            self.tree.heading(col, text=text, anchor=anchor)
            self.tree.column(col, width=100, anchor=anchor)
        self.tree.tag_configure("missing_music", foreground=BAD)
        self.tree.tag_configure("audio_warning", foreground=WARN)
        self.tree.tag_configure("audio_error", foreground=BAD, background="#3B1D24")
        self.tree.tag_configure("cancelled", foreground=TEXT_DISABLED)
        self.tree.tag_configure("completed", foreground=GOOD, background="#16302A")
        self.tree.bind("<Button-1>", self.on_tree_checkbox_click)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Return>", lambda e: self.toggle_play_pause())
        scroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        ttk.Label(p, text="Dobbeltklikk eller Enter spiller lydfilen. Klikk i «Løp gjennomført» for å krysse av.",
                  style="PanelDim.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 0))

        self.table_columns = tuple(columns)
        self.table_column_vars = {}
        for column, text in columns.items():
            variable = tk.BooleanVar(value=True)
            self.table_column_vars[column] = variable
            self.menu_tabellvisning.add_checkbutton(
                label=text,
                variable=variable,
                command=lambda changed=column: self.update_visible_table_columns(changed),
            )
        self.menu_tabellvisning.add_separator()
        self.menu_tabellvisning.add_command(label="Vis alle kolonner", command=self.show_all_table_columns)

        player = ttk.Frame(p, style="Panel.TFrame")
        player.grid(row=4, column=0, sticky="we", pady=(16, 0))
        self.btn_player_play_pause = ttk.Button(player, textvariable=self.play_pause_text,
                                                style="Accent.TButton", command=self.toggle_play_pause,
                                                width=8)
        self.btn_player_play_pause.pack(side="left")
        self.btn_player_stop = ttk.Button(player, text="Stopp", command=self.stop_playback, width=8)
        self.btn_player_stop.pack(side="left", padx=(8, 0))
        info = ttk.Frame(player, style="Panel.TFrame")
        info.pack(side="left", fill="x", expand=True, padx=(18, 0))
        self.player_track_var = tk.StringVar(value="Ingen spor valgt")
        self.player_time_var = tk.StringVar(value="0:00 / 0:00")
        ttk.Label(info, textvariable=self.player_track_var, style="PanelHead.TLabel").pack(anchor="w")
        self.player_progress = ttk.Progressbar(info, mode="determinate")
        self.player_progress.pack(fill="x", pady=4)
        ttk.Label(info, textvariable=self.player_time_var, style="PanelDim.TLabel").pack(anchor="w")

    # -------------------- Generelt --------------------
    def log(self, msg):
        self.log_widget.insert("end", msg + "\n")
        self.log_widget.see("end")
        self.root.update_idletasks()

    def update_clock(self):
        for item, char in zip(self.clock_items, datetime.now().strftime("%H:%M:%S")):
            self.clock_canvas.itemconfigure(item, text=char)
        self.root.after(500, self.update_clock)

    def output_dir(self):
        out_dir = (self.source_dir or Path(__file__).resolve().parent) / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        return out_dir

    def active_rows(self):
        return [r for r in self.rows if not is_cancelled(r.get("Påmelding", ""))]

    def visible_items(self):
        return [i for i in self.items if is_marker(i) or not is_cancelled(i.get("Påmelding", ""))]

    def index_of(self, item):
        return next((i for i, x in enumerate(self.items) if x is item), None)

    def confirm_replace(self):
        if self.dirty:
            return self.confirm_discard()
        if not self.items:
            return True
        return messagebox.askyesno(
            "Erstatte listen?",
            "Dette erstatter den nåværende listen. Vil du fortsette?",
            parent=self.root,
        )

    def confirm_discard(self):
        if not self.dirty or not self.items:
            return True
        answer = messagebox.askyesnocancel(
            "Ulagrede endringer",
            "Oppsettet har endringer som ikke er lagret. Vil du lagre først?",
            parent=self.root,
        )
        if answer is None:
            return False
        return self.save_setup() if answer else True

    def mark_dirty(self):
        if self._loading:
            return
        self.dirty = True
        self.update_title()

    def update_title(self):
        name = self.setup_path.name if self.setup_path else "Nytt oppsett"
        self.root.title(f"{APP_NAME} – {name}{' *' if self.dirty else ''}")

    def update_status(self):
        parts = []
        if self.setup_path:
            parts.append(f"Oppsett: {self.setup_path}")
        elif self.source_dir:
            parts.append(f"Kilde: {self.source_dir}")
        parts.append(f"{len(self.active_rows())} utøvere")
        if self.music_zip:
            parts.append(f"Musikk: {self.music_zip.name}")
        self.status_var.set(" · ".join(parts))

    # -------------------- Tidsplan --------------------
    @staticmethod
    def parse_optional_duration(value):
        value = (value or "").strip()
        if value in ("", "0"):
            return 0
        return parse_duration_mmss(value)

    @staticmethod
    def read_int(var, default):
        try:
            return int(var.get())
        except (ValueError, tk.TclError):
            return default

    def read_schedule_settings(self):
        start = parse_time_hhmm(self.start_time_var.get())
        if not start:
            return "Ugyldig starttid. Bruk TT:MM."
        run_seconds = parse_duration_mmss(self.interval_var.get())
        if not run_seconds:
            return "Ugyldig lengde på løp. Bruk M:SS."
        ready = self.parse_optional_duration(self.ready_var.get())
        if ready is None:
            return "Ugyldig «klart til neste løp». Bruk M:SS."
        warmup = self.parse_optional_duration(self.warmup_var.get())
        if warmup is None:
            return "Ugyldig oppvarming. Bruk M:SS."
        water_seconds = self.parse_optional_duration(self.water_duration_var.get())
        if water_seconds is None:
            return "Ugyldig lengde på vanningspause. Bruk M:SS."
        return {
            "run_seconds": run_seconds,
            "start_time": start,
            "ready_seconds": ready,
            "warmup_seconds": warmup,
            "water_seconds": water_seconds,
        }

    def relayout(self):
        self.items = layout_items(
            self.rows, self.read_int(self.group_size_var, 8), self.read_int(self.water_count_var, 0)
        )

    def on_layout_setting_change(self, regroup):
        if self._loading:
            return
        if regroup:
            self.relayout()
        else:
            self.items = place_water_pauses(self.items, self.read_int(self.water_count_var, 0))
        self.refresh_schedule()

    def sync_water_count(self):
        self._loading = True
        try:
            self.water_count_var.set(str(sum(1 for i in self.items if i.get("Kind") == "pause")))
        finally:
            self._loading = False

    def selected_item(self):
        selected = self.schedule_tree.selection()
        if not selected:
            return None
        return self.items[int(selected[0][1:])]

    def refresh_schedule(self, select=None):
        if not hasattr(self, "schedule_tree"):
            return
        if select is None:
            select = self.selected_item()
        self.mark_dirty()
        settings = self.read_schedule_settings()
        if isinstance(settings, str):
            self.summary_label.configure(style="PanelWarn.TLabel")
            self.summary_var.set(settings)
            return
        self.summary_label.configure(style="PanelDim.TLabel")

        self.entries = build_startliste(self.items, **settings)
        for row in self.rows:
            row["StartTid"] = ""
            row["SluttTid"] = ""
        tree = self.schedule_tree
        tree.delete(*tree.get_children())
        positions = {id(item): i for i, item in enumerate(self.items)}
        group_size = self.read_int(self.group_size_var, 8)
        overflows = group_overflows(self.items, group_size)
        self.overflow_tips = {}
        select_iid = None
        for entry in self.entries:
            item = entry["row"]
            iid = f"i{positions[id(item)]}"
            if item is select:
                select_iid = iid
            if entry["kind"] == "run":
                item["StartTid"] = entry["start"]
                item["SluttTid"] = entry["end"]
            tags = () if entry["kind"] == "run" else (entry["kind"],)
            count = overflows.get(id(item))
            if count:
                tags = ("overfull",)
                extra = count - group_size
                self.overflow_tips[iid] = (
                    f"{entry['navn']} har {count} utøvere, men maks er {group_size}.\n"
                    f"Flytt {extra} utøver{'e' if extra > 1 else ''} til en annen gruppe, "
                    "eller øk «Utøvere per gruppe»."
                )
            tree.insert("", "end", iid=iid, tags=tags, values=(
                entry["nr"], entry["start"], "-", entry["end"], entry["duration"],
                entry["navn"], entry["klubb"],
            ))
        if select_iid:
            tree.selection_set(select_iid)
            tree.focus(select_iid)
            tree.see(select_iid)

        runners = sum(1 for e in self.entries if e["kind"] == "run")
        if not runners:
            self.summary_var.set("Ingen utøvere ennå. Hent fra isonen eller legg til manuelt.")
        else:
            groups = sum(1 for e in self.entries if e["kind"] == "warmup")
            pauses = sum(1 for e in self.entries if e["kind"] == "pause")
            text = (
                f"{runners} utøvere · {groups} oppvarmingsgrupper · "
                f"{pauses} vanningspauser · ferdig kl. {self.entries[-1]['end']}"
            )
            wanted = self.read_int(self.water_count_var, 0)
            if pauses < wanted:
                text += f" (bare {pauses} vanningspauser får plass mellom gruppene)"
            if overflows:
                text += f" · {len(overflows)} gruppe{'r' if len(overflows) > 1 else ''} har for mange utøvere"
                self.summary_label.configure(style="PanelWarn.TLabel")
            self.summary_var.set(text)
        self.update_status()
        self.refresh_table()

    def add_runner(self):
        given, family = split_name(self.new_name_var.get().strip())
        if not given:
            self.name_entry.focus_set()
            return
        row = new_participant(given, family, self.new_club_var.get())
        # Ny oppvarmingsgruppe når den siste gruppen er full (eller det ikke finnes noen).
        in_last_group = 0
        has_group = False
        for item in reversed(self.visible_items()):
            if item.get("Kind") == "warmup":
                has_group = True
                break
            if not is_marker(item):
                in_last_group += 1
        if not has_group or in_last_group >= self.read_int(self.group_size_var, 8):
            self.items.append({"Kind": "warmup"})
        self.items.append(row)
        self.match_missing_music()
        self.new_name_var.set("")
        self.name_entry.focus_set()
        self.log(f"La til utøver: {display_name(row)}")
        self.refresh_schedule(select=row)

    def edit_selected(self):
        row = self.selected_item()
        if row is None or is_marker(row):
            return
        name = simpledialog.askstring("Rediger utøver", "Navn:", initialvalue=display_name(row),
                                      parent=self.root)
        if name is None or not name.strip():
            return
        club = simpledialog.askstring("Rediger utøver", "Klubb:",
                                      initialvalue=row.get("Organisation", ""), parent=self.root)
        if club is None:
            return
        given, family = split_name(name.strip())
        if (given, family) != (row.get("GivenName"), row.get("FamilyName")):
            row["MusikkFil"] = ""
        row["GivenName"], row["FamilyName"] = given, family
        row["NavnFraIsonen"] = f"{given} {family}".strip()
        row["PrintName"] = f"{family}, {given}".strip(", ")
        row["Organisation"] = club.strip()
        self.match_missing_music()
        self.refresh_schedule(select=row)

    def move_item(self, src, dst):
        # Etter pop havner raden under målet når den flyttes ned, og over når den flyttes opp.
        group_size = self.read_int(self.group_size_var, 8)
        before = group_overflows(self.items, group_size)
        item = self.items.pop(src)
        self.items.insert(dst, item)
        self.refresh_schedule(select=item)
        after = group_overflows(self.items, group_size)
        grown = [key for key, count in after.items() if count > before.get(key, 0)]
        if grown:
            iids = {f"i{k}": id(x) for k, x in enumerate(self.items)}
            tips = [tip for iid, tip in self.overflow_tips.items() if iids.get(iid) in grown]
            messagebox.showwarning("For mange utøvere i gruppen", "\n\n".join(tips), parent=self.root)

    def on_schedule_hover(self, event):
        tip = self.overflow_tips.get(self.schedule_tree.identify_row(event.y))
        if not tip:
            self.hide_tip()
            return
        if self.tip is None:
            self.tip = tk.Toplevel(self.root)
            self.tip.overrideredirect(True)
            self.tip.attributes("-topmost", True)
            self.tip_label = tk.Label(self.tip, bg=PANEL, fg=TEXT, font=self.fonts["body"], justify="left",
                                      wraplength=380, padx=10, pady=6, highlightthickness=2,
                                      highlightbackground=WARN)
            self.tip_label.pack()
        self.tip_label.configure(text=tip)
        self.tip.geometry(f"+{event.x_root + 14}+{event.y_root + 18}")
        self.tip.deiconify()

    def hide_tip(self):
        if self.tip is not None:
            self.tip.withdraw()

    def move_selected(self, delta):
        item = self.selected_item()
        if item is None:
            messagebox.showinfo("Info", "Velg en rad i tabellen først.")
            return
        visible = self.visible_items()
        pos = next(i for i, x in enumerate(visible) if x is item)
        target = pos + delta
        if 0 <= target < len(visible):
            self.move_item(self.index_of(item), self.index_of(visible[target]))

    def on_drag_start(self, event):
        self.hide_tip()
        self.drag_iid = self.schedule_tree.identify_row(event.y) or None

    def on_drag_motion(self, event):
        tree = self.schedule_tree
        if not self.drag_iid:
            return
        if event.y < 30:
            tree.yview_scroll(-1, "units")
        elif event.y > tree.winfo_height() - 10:
            tree.yview_scroll(1, "units")
        target = tree.identify_row(event.y)
        bbox = tree.bbox(target) if target and target != self.drag_iid else None
        if not bbox:
            self.drop_line.place_forget()
            return
        _, y, _, height = bbox
        below = tree.index(target) > tree.index(self.drag_iid)
        self.drop_line.place(x=0, y=y + height - 2 if below else y - 1, relwidth=1)
        tree.configure(cursor="sb_v_double_arrow")

    def on_drag_drop(self, event):
        self.drop_line.place_forget()
        self.schedule_tree.configure(cursor="")
        source, self.drag_iid = self.drag_iid, None
        target = self.schedule_tree.identify_row(event.y)
        if source and target and source != target:
            self.move_item(int(source[1:]), int(target[1:]))

    def delete_selected(self):
        item = self.selected_item()
        if item is None:
            messagebox.showinfo("Info", "Velg en rad i tabellen først.")
            return
        if item.get("Kind") == "pause":
            navn, question = "vanningspause", "Slette valgt vanningspause?"
        elif item.get("Kind") == "warmup":
            navn, question = "oppvarmingsgruppe", "Slette valgt oppvarmingsgruppe?\n\nUtøverne blir stående."
        else:
            navn = display_name(item)
            question = f"Slette valgt utøver?\n\n{navn}"
        if not messagebox.askyesno("Bekreft sletting", question):
            return
        visible = self.visible_items()
        pos = next(i for i, x in enumerate(visible) if x is item)
        del self.items[self.index_of(item)]
        self.sync_water_count()
        self.log(f"Slettet {navn}.")
        remaining = self.visible_items()
        self.refresh_schedule(select=remaining[min(pos, len(remaining) - 1)] if remaining else None)

    def set_runner_order(self, runners):
        # Bytter rekkefølgen på utøverne, men lar grupper og pauser stå på sine plasser.
        slots = [k for k, item in enumerate(self.items) if not is_marker(item)]
        for slot, row in zip(slots, runners):
            self.items[slot] = row
        self.refresh_schedule()

    def shuffle_rows(self):
        runners = self.rows
        random.shuffle(runners)
        self.set_runner_order(runners)

    def sort_by_given(self):
        self.set_runner_order(sorted(
            self.rows, key=lambda r: (normalize_text(r.get("GivenName")), normalize_text(r.get("FamilyName")))
        ))

    def sort_by_family(self):
        self.set_runner_order(sorted(
            self.rows, key=lambda r: (normalize_text(r.get("FamilyName")), normalize_text(r.get("GivenName")))
        ))

    def clear_rows(self):
        if not self.confirm_discard():
            return
        self.items = []
        self.zip_path = None
        self.setup_path = None
        self.log("Startet et nytt oppsett.")
        self.refresh_schedule()
        self.dirty = False
        self.update_title()

    # -------------------- Oppsettfil --------------------
    def save_setup(self, ask=False):
        path = self.setup_path
        if ask or not path:
            date_obj = parse_date_ddmmyy(self.start_date_var.get())
            date_part = date_obj.strftime("%Y-%m-%d") if date_obj else "dato"
            location = self.location_var.get().strip() or "sted"
            path = filedialog.asksaveasfilename(
                title="Lagre oppsett",
                defaultextension=".json",
                filetypes=SETUP_FILETYPES,
                initialfile=sanitize_filename(f"Oppsett_{location}_{date_part}") + ".json",
                initialdir=str(path.parent if path else self.output_dir()),
            )
            if not path:
                return False
            path = Path(path)
        data = {
            "type": "kunstlopassistenten-oppsett",
            "version": 1,
            "saved": datetime.now().isoformat(timespec="seconds"),
            "settings": {key: var.get() for key, var in self.setup_vars.items()},
            "source_dir": str(self.source_dir) if self.source_dir else "",
            "music_zip": str(self.music_zip) if self.music_zip else "",
            "items": self.items,
        }
        try:
            path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke lagre oppsettet:\n{exc}")
            return False
        self.setup_path = path
        self.dirty = False
        self.update_title()
        self.update_status()
        self.log(f"Lagret oppsett: {path}")
        return True

    def open_setup(self):
        if not self.confirm_discard():
            return
        path = filedialog.askopenfilename(
            title="Åpne oppsett",
            filetypes=SETUP_FILETYPES,
            initialdir=str(self.setup_path.parent if self.setup_path else self.output_dir()),
        )
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke lese oppsettet:\n{exc}")
            return
        if not isinstance(data, dict) or not isinstance(data.get("items"), list):
            if isinstance(data, dict) and isinstance(data.get("order"), list):
                self.load_legacy_order(data["order"], path)
            else:
                messagebox.showerror("Feil", "Filen er ikke et oppsett fra Kunstløpassistenten.")
            return

        self._loading = True
        try:
            for key, value in (data.get("settings") or {}).items():
                if key in self.setup_vars:
                    self.setup_vars[key].set(value)
        finally:
            self._loading = False
        self.items = [item for item in data["items"] if isinstance(item, dict)]
        self.source_dir = Path(data["source_dir"]) if data.get("source_dir") else None
        self.zip_path = None
        music = data.get("music_zip")
        if music and Path(music).is_file():
            self.load_music_zip(music, match=False)
            self.match_missing_music()
        elif music:
            self.log(f"Fant ikke musikk-zip fra oppsettet: {music}")
        self.setup_path = Path(path)
        self.refresh_schedule()
        self.dirty = False
        self.update_title()
        self.log(f"Åpnet oppsett: {path}")

    # -------------------- Innlesing --------------------
    def load_isonen(self):
        initial = self.source_dir or (Path.home() / "Downloads")
        path = filedialog.askopenfilename(
            title="Velg deltakerliste fra isonen",
            initialdir=str(initial),
            filetypes=[("Excel", "*.xlsx"), ("Alle filer", "*.*")],
        )
        if not path or not self.confirm_replace():
            return
        rows = load_participants_from_excel(path, self.log)
        if not rows:
            messagebox.showerror("Feil", "Fant ingen deltakere i excel-filen.")
            return
        self.rows = rows
        self.zip_path = None
        self.setup_path = None
        self.source_dir = Path(path).parent
        cancelled = len(rows) - len(self.active_rows())
        if cancelled:
            self.log(f"{cancelled} avmeldte er tatt ut av tidsplanen.")
        if self.music_files:
            self.match_all_music()
        self.relayout()
        self.refresh_schedule()

    def choose_music_zip(self):
        initial = self.source_dir or (Path.home() / "Downloads")
        path = filedialog.askopenfilename(
            title="Velg musikk-zip fra stevnedatabasen",
            initialdir=str(initial),
            filetypes=[("Zip", "*.zip"), ("Alle filer", "*.*")],
        )
        if path and self.load_music_zip(path):
            self.refresh_schedule()

    def scan_folder(self):
        path = filedialog.askdirectory(
            title="Velg mappe med isonen-liste, FMS-data og musikk",
            initialdir=str(self.source_dir or Path.home()),
        )
        if path and self.confirm_replace():
            self.read_zip(Path(path))

    def read_zip(self, folder):
        self.log_widget.delete("1.0", "end")
        excel_files = sorted(folder.glob("Deltakerliste*.xlsx"))
        if not excel_files:
            messagebox.showerror("Feil", "Fant ingen Deltakerliste*.xlsx i mappen.")
            return
        zips = sorted(folder.glob("*.zip"))
        data_zips = [z for z in zips if z.name.lower().startswith(("fmsdata", "fsmdata"))]
        music_zips = [z for z in zips if z.name.lower().startswith("musikk")]

        if len(excel_files) > 1:
            self.log(f"Fant flere deltakerlister. Bruker: {excel_files[0].name}")
        rows = load_participants_from_excel(excel_files[0], self.log)
        if not rows:
            messagebox.showerror("Feil", "Fant ingen deltakere i excel-filen.")
            return
        self.rows = rows
        self.setup_path = None
        self.source_dir = folder
        self.zip_path = data_zips[0] if data_zips else None
        if self.zip_path:
            if len(data_zips) > 1:
                self.log(f"Fant flere FMSData-zip. Bruker: {self.zip_path.name}")
            self.match_fms_data(self.zip_path)
        else:
            self.log("Fant ingen FMSData*.zip. Hopper over FMS-data.")
        if music_zips:
            if len(music_zips) > 1:
                self.log(f"Fant flere musikk-zip. Bruker: {music_zips[0].name}")
            self.load_music_zip(music_zips[0])
        elif self.music_files:
            self.match_all_music()
        self.log(f"Totalt deltakere: {len(self.rows)}")
        self.relayout()
        self.refresh_schedule()

    def match_fms_data(self, zip_path):
        self.log(f"Leser zip: {zip_path.name}")
        zip_rows = []
        with zipfile.ZipFile(zip_path, "r") as zf:
            xml_entries = [e for e in zf.infolist() if e.filename.lower().endswith(".xml")]
            if not xml_entries:
                self.log("Fant ingen xml-filer i FMS-zip.")
                return
            for entry in xml_entries:
                self.log(f"Leser fil: {entry.filename}")
                xml_text = decode_xml_bytes(zf.read(entry))
                if "judges" in entry.filename.lower():
                    parse_officials(xml_text, self.log)
                else:
                    zip_rows.extend(parse_competition(xml_text, self.log))

        zip_people_strict = {}
        zip_people_loose = {}
        for row in zip_rows:
            strict_key = build_name_key(row.get("GivenName"), row.get("FamilyName"), strict=True)
            loose_key = build_name_key(row.get("GivenName"), row.get("FamilyName"), strict=False)
            if strict_key[0] or strict_key[1]:
                zip_people_strict.setdefault(strict_key, row)
            if loose_key[0] or loose_key[1]:
                zip_people_loose.setdefault(loose_key, []).append(strict_key)

        used_zip_keys = set()
        used_loose_keys = set()
        fms_fields = ("ParticipantCode", "Event", "EntryOrder", "Music1", "Music2", "Club1", "Club2",
                      "ElementsFree", "ElementsShort")
        for row in self.rows:
            strict_key = build_name_key(row.get("GivenName"), row.get("FamilyName"), strict=True)
            loose_key = build_name_key(row.get("GivenName"), row.get("FamilyName"), strict=False)
            zip_row = None
            match_type = ""
            if strict_key in zip_people_strict and strict_key not in used_zip_keys:
                zip_row = zip_people_strict.get(strict_key)
                used_zip_keys.add(strict_key)
                match_type = "eksakt"
            else:
                for candidate_key in zip_people_loose.get(loose_key, []):
                    if candidate_key in used_zip_keys:
                        continue
                    zip_row = zip_people_strict.get(candidate_key)
                    if zip_row:
                        used_zip_keys.add(candidate_key)
                        match_type = "loose"
                        break
            if zip_row:
                for field in fms_fields:
                    row[field] = zip_row.get(field, "")
                zip_print = zip_row.get("PrintName") or f"{zip_row.get('GivenName', '')} {zip_row.get('FamilyName', '')}".strip()
                row["NavnFraFsm"] = zip_print
                row["Manglende i zip"] = ""
                if loose_key[0] or loose_key[1]:
                    used_loose_keys.add(loose_key)
                if match_type == "loose":
                    self.log(f"Matcher (loose): {row.get('NavnFraIsonen', '')} -> {zip_print}")
                else:
                    self.log(f"Matcher: {row.get('NavnFraIsonen', '')} -> {zip_print}")
            else:
                row["Manglende i zip"] = "JA"
                row["NavnFraFsm"] = ""
                self.log(f"Mangler i FSM: {row.get('NavnFraIsonen', '')}")

        for key, zip_row in zip_people_strict.items():
            if key in used_zip_keys:
                continue
            loose_key = build_name_key(zip_row.get("GivenName"), zip_row.get("FamilyName"), strict=False)
            if loose_key in used_loose_keys:
                continue
            row = new_participant(zip_row.get("GivenName"), zip_row.get("FamilyName"),
                                  zip_row.get("Organisation"), zip_row.get("Gender"))
            for field in fms_fields:
                row[field] = zip_row.get(field, "")
            row["NavnFraIsonen"] = ""
            row["NavnFraFsm"] = zip_row.get("PrintName") or display_name(row)
            row["PrintName"] = row["NavnFraFsm"]
            self.log(f"Ekstra i FSM (ny rad): {row['NavnFraFsm']}")
            self.items.append(row)

    def load_music_zip(self, music_zip, match=True):
        self.music_zip = Path(music_zip)
        self.music_cache_dir = None
        self.music_file_info = {}
        self.music_files = []
        self.music_durations = {}
        self.log(f"Leser musikk-zip: {self.music_zip.name}")
        try:
            with zipfile.ZipFile(self.music_zip, "r") as mz:
                for entry in mz.infolist():
                    declared_suffix = Path(entry.filename).suffix.lower()
                    if entry.is_dir() or declared_suffix not in {".mp3", ".wav", ".m4a", ".mp4", ".aac"}:
                        continue
                    data = mz.read(entry)
                    audio_format, detected_suffix = detect_audio_format(data)
                    format_mismatch = bool(detected_suffix and declared_suffix != detected_suffix)
                    self.music_files.append(entry.filename)
                    self.music_durations[entry.filename] = get_audio_duration(data, audio_format)
                    self.music_file_info[entry.filename] = {
                        "format": audio_format,
                        "suffix": detected_suffix or declared_suffix,
                        "declared_suffix": declared_suffix,
                        "format_mismatch": format_mismatch,
                        "playable": audio_format != "Ukjent",
                    }
                    if audio_format == "Ukjent":
                        self.log(f"ADVARSEL: Ukjent lydformat: {entry.filename}")
                    elif format_mismatch:
                        self.log(
                            f"ADVARSEL: {entry.filename} har filendelsen "
                            f"{declared_suffix}, men er egentlig {audio_format}."
                        )
        except Exception as exc:
            self.log(f"Kunne ikke lese musikk-zip: {self.music_zip} ({exc})")
            messagebox.showerror("Feil", f"Kunne ikke lese musikk-zip:\n{exc}")
            self.music_zip = None
            self.music_info_var.set("Ingen musikk-zip valgt.")
            return False
        warnings = sum(1 for info in self.music_file_info.values() if info.get("format_mismatch"))
        errors = sum(1 for info in self.music_file_info.values() if not info.get("playable"))
        self.log(
            f"Lydkontroll: {len(self.music_files)} filer, {warnings} med feil filendelse, "
            f"{errors} med ukjent format."
        )
        self.music_info_var.set(f"{self.music_zip.name} · {len(self.music_files)} lydfiler")
        if match:
            self.match_all_music()
        return True

    def apply_music_match(self, row, used_files):
        matched = self.match_music_file(row, self.music_files, used_files) if self.music_files else ""
        info = self.music_file_info.get(matched, {}) if matched else {}
        seconds = self.music_durations.get(matched) if matched else None
        row["Musikk"] = "ok" if matched else "mangler"
        row["MusikkFil"] = matched
        row["MusikkFormat"] = info.get("format", "")
        row["MusikkFormatAvvik"] = info.get("format_mismatch", False)
        row["MusikkSpillbar"] = info.get("playable", bool(matched))
        if matched and seconds is None:
            row["MusikkTid"] = "Klarer ikke å hente tid"
            row["MusikkSek"] = ""
        else:
            row["MusikkTid"] = format_duration(seconds)
            row["MusikkSek"] = int(round(seconds)) if seconds else ""

    def match_all_music(self):
        used = set()
        for row in self.rows:
            self.apply_music_match(row, used)
        missing = sum(1 for r in self.active_rows() if not r.get("MusikkFil"))
        if missing:
            self.log(f"{missing} utøvere mangler lydfil.")

    def match_missing_music(self):
        if not self.music_files:
            return
        used = {r.get("MusikkFil") for r in self.rows if r.get("MusikkFil")}
        for row in self.rows:
            if not row.get("MusikkFil"):
                self.apply_music_match(row, used)

    def match_music_file(self, row, music_files, used_files):
        given = row.get("GivenName")
        family = row.get("FamilyName")
        given_tokens = tokenize_name(given)
        family_tokens = tokenize_name(family)
        if not family_tokens:
            return ""

        def family_match(fname):
            hay = normalize_text(fname)
            return all(token in hay for token in family_tokens) or family_tokens[0] in hay

        def given_match(fname):
            if not given_tokens:
                return False
            hay = normalize_text(fname)
            return given_tokens[0] in hay

        # Pass 1: require both family + given (if given exists), prefer unused.
        if given_tokens:
            for fname in music_files:
                if fname in used_files:
                    continue
                if family_match(fname) and given_match(fname):
                    used_files.add(fname)
                    return fname

        # Pass 2: family-only fallback, prefer unused.
        for fname in music_files:
            if fname in used_files:
                continue
            if family_match(fname):
                used_files.add(fname)
                return fname
        return ""

    # -------------------- Musikktabell --------------------
    def refresh_table(self):
        selected = self.tree.selection()
        selected_index = self.tree.index(selected[0]) if selected else None
        self.tree.delete(*self.tree.get_children())
        has_music = bool(self.music_files)
        rows_values = []
        nr = 1
        for row in self.rows:
            cancelled = is_cancelled(row.get("Påmelding", ""))
            missing = has_music and not row.get("MusikkFil")
            if cancelled:
                tags = ("cancelled",)
            elif missing:
                tags = ("missing_music",)
            elif has_music and not row.get("MusikkSpillbar", True):
                tags = ("audio_error",)
            elif row.get("MusikkFormatAvvik"):
                tags = ("audio_warning",)
            else:
                tags = ()
            if row.get("LopGjennomfort"):
                tags = tags + ("completed",)
            values = (
                "" if cancelled else nr,
                "☑" if row.get("LopGjennomfort") else "☐",
                row.get("StartTid", ""),
                display_name(row),
                row.get("Organisation", ""),
                row.get("Påmelding", ""),
                "mangler musikk" if missing else row.get("MusikkFil", ""),
                row.get("MusikkFormat", ""),
                row.get("MusikkTid", ""),
            )
            if not cancelled:
                nr += 1
            rows_values.append(values)
            self.tree.insert("", "end", values=values, tags=tags)
        self.autosize_columns(rows_values)
        items = self.tree.get_children()
        if selected_index is not None and selected_index < len(items):
            self.tree.selection_set(items[selected_index])

    def autosize_columns(self, rows_values):
        font = tkfont.nametofont("TkDefaultFont")
        columns = list(self.tree["columns"])
        padding = 24
        max_widths = {}
        for col in columns:
            heading = self.tree.heading(col).get("text", "")
            max_widths[col] = font.measure(str(heading)) + padding
        for row in rows_values:
            for col, value in zip(columns, row):
                width = font.measure(str(value)) + padding
                if width > max_widths.get(col, 0):
                    max_widths[col] = width
        for col in columns:
            self.tree.column(col, width=max_widths.get(col, 80), stretch=col == "musikknavn")

    def update_visible_table_columns(self, changed_column=None):
        visible_columns = tuple(
            column
            for column in self.table_columns
            if self.table_column_vars[column].get()
        )
        if not visible_columns:
            if changed_column:
                self.table_column_vars[changed_column].set(True)
            messagebox.showinfo(
                "Tabellvisning",
                "Minst én kolonne må være synlig.",
            )
            return
        self.tree.configure(displaycolumns=visible_columns)

    def show_all_table_columns(self):
        for variable in self.table_column_vars.values():
            variable.set(True)
        self.tree.configure(displaycolumns=self.table_columns)

    def get_tree_column_name(self, x_position):
        display_column = self.tree.identify_column(x_position)
        if not display_column:
            return None
        display_index = int(display_column[1:]) - 1
        columns = list(self.tree["columns"])
        display_columns = list(self.tree["displaycolumns"])
        if display_columns == ["#all"]:
            display_columns = columns
        if display_index < 0 or display_index >= len(display_columns):
            return None
        return display_columns[display_index]

    def on_tree_checkbox_click(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return None
        if self.get_tree_column_name(event.x) != "lop_gjennomfort":
            return None

        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return "break"
        row_index = self.tree.index(item_id)
        if row_index < 0 or row_index >= len(self.rows):
            return "break"

        row = self.rows[row_index]
        self.tree.selection_set(item_id)
        self.tree.focus(item_id)
        completed = bool(row.get("LopGjennomfort"))
        name = (
            row.get("NavnFraIsonen")
            or row.get("PrintName")
            or row.get("NavnFraFsm")
            or "denne deltakeren"
        )
        if completed:
            question = f"Vil du fjerne markeringen som gjennomført for {name}?"
        else:
            question = f"Vil du markere løpet som gjennomført for {name}?"
        if not messagebox.askyesno(
            "Bekreft løpsstatus",
            question,
            parent=self.root,
        ):
            return "break"

        row["LopGjennomfort"] = not completed
        self.dirty = True
        status = "gjennomført" if not completed else "ikke gjennomført"
        self.log(f"Løpsstatus for {name}: {status}.")
        self.refresh_table()
        items = self.tree.get_children()
        if row_index < len(items):
            self.tree.selection_set(items[row_index])
            self.tree.focus(items[row_index])
        return "break"

    def on_tree_double_click(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        if self.get_tree_column_name(event.x) == "lop_gjennomfort":
            return
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        self.play_mp3_file(self.rows[self.tree.index(item_id)].get("MusikkFil", ""))

    # -------------------- Avspilling --------------------
    def get_cached_mp3_path(self, filename):
        if not filename:
            messagebox.showinfo("Info", "Ingen lydfil registrert på denne raden.")
            return None
        if not self.music_zip:
            messagebox.showerror("Feil", "Fant ingen musikk-zip.")
            return None
        try:
            with zipfile.ZipFile(self.music_zip, "r") as mz:
                try:
                    data = mz.read(filename)
                except KeyError:
                    messagebox.showerror("Feil", f"Fant ikke lydfil i zip: {filename}")
                    return None
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke lese musikk-zip: {exc}")
            return None

        if not self.music_cache_dir:
            base = Path(tempfile.gettempdir()) / "fms_gui_music_cache"
            base.mkdir(parents=True, exist_ok=True)
            safe_zip = sanitize_filename(self.music_zip.stem)
            self.music_cache_dir = base / safe_zip
            self.music_cache_dir.mkdir(parents=True, exist_ok=True)

        audio_info = self.music_file_info.get(filename)
        if not audio_info:
            audio_format, detected_suffix = detect_audio_format(data)
            audio_info = {
                "format": audio_format,
                "suffix": detected_suffix or Path(filename).suffix.lower(),
            }
            self.music_file_info[filename] = audio_info

        safe_stem = sanitize_filename(Path(filename).stem)
        content_hash = hashlib.sha256(data).hexdigest()[:12]
        safe_name = f"{safe_stem}_{content_hash}{audio_info.get('suffix', '')}"
        out_path = self.music_cache_dir / safe_name
        try:
            if out_path.is_file() and out_path.stat().st_size == len(data):
                return out_path
        except OSError:
            pass
        try:
            out_path.write_bytes(data)
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke skrive lydfil: {exc}")
            return None
        return out_path

    def ensure_audio_backend(self):
        if self.audio_ready:
            return True
        try:
            import pygame
        except Exception as exc:
            self.log(f"Intern lydmotor (pygame) er ikke tilgjengelig: {exc}")
            return False
        try:
            pygame.mixer.init()
        except Exception as exc:
            self.log(f"Kunne ikke starte intern lydmotor: {exc}")
            return False
        self.audio_backend = pygame
        self.audio_ready = True
        return True

    def release_internal_audio(self):
        if not self.audio_ready:
            return
        try:
            self.audio_backend.mixer.music.stop()
        except Exception:
            pass
        try:
            self.audio_backend.mixer.music.unload()
        except Exception:
            pass

    def set_external_playback_status(self, filename, player_name):
        self.current_track = filename
        self.external_playback = True
        self.is_paused = False
        self.play_pause_text.set("Spill")
        row = self.find_row_by_mp3(filename)
        if row:
            title = row.get("NavnFraIsonen") or row.get("PrintName") or ""
            self.player_track_var.set(f"{title} - {filename}".strip(" -"))
        else:
            self.player_track_var.set(filename)
        self.player_time_var.set(player_name)

    def start_vlc_playback(self, path, filename, fallback_to_default=True):
        vlc_path = find_vlc_executable()
        if vlc_path:
            try:
                self.vlc_process = subprocess.Popen(
                    [vlc_path, "--play-and-exit", "--no-video", str(path)]
                )
                self.set_external_playback_status(filename, "VLC")
                self.log(f"Spiller med VLC: {filename}")
                return True
            except Exception as exc:
                self.log(f"Kunne ikke starte VLC: {exc}")

        if fallback_to_default:
            try:
                os.startfile(str(path))
                self.set_external_playback_status(filename, "Ekstern avspiller")
                self.log(f"Spiller med Windows-standardspiller: {filename}")
                return True
            except Exception as exc:
                self.log(f"Kunne ikke starte ekstern avspiller: {exc}")

        messagebox.showerror(
            "Feil",
            "Kunne ikke spille av lydfilen. VLC ble ikke funnet eller kunne ikke startes.",
        )
        return False

    def start_playback(self, path, filename):
        if not self.ensure_audio_backend():
            self.log("Intern lydmotor er ikke tilgjengelig. Prøver VLC.")
            return self.start_vlc_playback(path, filename)
        row = self.find_row_by_mp3(filename)
        self.current_duration = safe_int(row.get("MusikkSek")) if row else 0
        try:
            self.release_internal_audio()
            self.audio_backend.mixer.music.load(str(path))
            self.audio_backend.mixer.music.play()
        except Exception as exc:
            self.log(f"Kunne ikke spille av med intern avspiller: {exc}")
            self.release_internal_audio()
            self.log("Prøver VLC som reserveavspiller.")
            return self.start_vlc_playback(path, filename)
        self.current_track = filename
        self.is_paused = False
        self.external_playback = False
        self.play_pause_text.set("Pause")
        if row:
            title = row.get("NavnFraIsonen") or row.get("PrintName") or ""
            self.player_track_var.set(f"{title} - {filename}".strip(" -"))
        else:
            self.player_track_var.set(filename)
        self.log(f"Spiller: {filename}")
        return True

    def play_mp3_file(self, filename):
        path = self.get_cached_mp3_path(filename)
        if not path:
            return
        if self.use_external_player_var.get():
            self.start_vlc_playback(path, filename)
            return
        self.start_playback(path, filename)

    def get_selected_mp3_filename(self):
        selected = self.tree.selection()
        if not selected:
            return ""
        return self.rows[self.tree.index(selected[0])].get("MusikkFil", "")

    def toggle_play_pause(self):
        if self.use_external_player_var.get():
            messagebox.showinfo(
                "Info",
                "Bruker VLC direkte. Pause/fortsett styres i VLC.",
            )
            return
        if self.external_playback:
            messagebox.showinfo(
                "Info",
                "Spiller med reserveavspiller. Pause/fortsett styres der.",
            )
            return
        if self.audio_ready and self.is_paused:
            self.audio_backend.mixer.music.unpause()
            self.is_paused = False
            self.play_pause_text.set("Pause")
            self.log("Fortsetter avspilling.")
            return
        if self.audio_ready and self.audio_backend.mixer.music.get_busy():
            self.audio_backend.mixer.music.pause()
            self.is_paused = True
            self.play_pause_text.set("Spill")
            self.log("Pause.")
            return
        filename = self.get_selected_mp3_filename()
        if not filename:
            messagebox.showinfo("Info", "Velg en rad med lydfil først.")
            return
        path = self.get_cached_mp3_path(filename)
        if not path:
            return
        self.start_playback(path, filename)

    def stop_playback(self):
        if self.use_external_player_var.get():
            if self.vlc_process and self.vlc_process.poll() is None:
                self.vlc_process.terminate()
                self.vlc_process = None
                self.external_playback = False
                self.play_pause_text.set("Spill")
                self.log("Stoppet VLC-avspilling.")
            else:
                messagebox.showinfo("Info", "Stopp avspillingen i VLC.")
            return
        if self.external_playback:
            if self.vlc_process and self.vlc_process.poll() is None:
                self.vlc_process.terminate()
                self.vlc_process = None
                self.external_playback = False
                self.play_pause_text.set("Spill")
                self.log("Stoppet VLC-avspilling.")
            else:
                messagebox.showinfo("Info", "Stopp avspillingen i den eksterne spilleren.")
            return
        if not self.ensure_audio_backend():
            return
        self.release_internal_audio()
        self.current_track = ""
        self.is_paused = False
        self.external_playback = False
        self.play_pause_text.set("Spill")
        self.log("Stoppet avspilling.")

    def find_row_by_mp3(self, filename):
        for row in self.rows:
            if row.get("MusikkFil") == filename:
                return row
        return None

    def on_close(self):
        if not self.confirm_discard():
            return
        if self.vlc_process and self.vlc_process.poll() is None:
            try:
                self.vlc_process.terminate()
            except Exception:
                pass
        if self.audio_ready:
            self.release_internal_audio()
            try:
                self.audio_backend.mixer.quit()
            except Exception:
                pass
        self.root.destroy()

    def update_player_ui(self):
        if self.external_playback:
            if self.vlc_process and self.vlc_process.poll() is not None:
                self.vlc_process = None
                self.current_track = ""
                self.external_playback = False
                self.play_pause_text.set("Spill")
                self.player_time_var.set("0:00 / 0:00")
            self.player_progress.config(maximum=1, value=0)
            self.root.after(500, self.update_player_ui)
            return
        elapsed = 0
        if self.audio_ready:
            pos_ms = self.audio_backend.mixer.music.get_pos()
            if pos_ms and pos_ms > 0:
                elapsed = pos_ms / 1000.0
            if (
                self.current_track
                and not self.is_paused
                and not self.audio_backend.mixer.music.get_busy()
            ):
                self.release_internal_audio()
                self.current_track = ""
                self.play_pause_text.set("Spill")
        duration = self.current_duration or 0
        if duration > 0:
            self.player_progress.config(maximum=duration, value=min(elapsed, duration))
            elapsed_text = format_duration(elapsed)
            total_text = format_duration(duration)
            self.player_time_var.set(f"{elapsed_text} / {total_text}")
        else:
            self.player_progress.config(maximum=1, value=0)
            self.player_time_var.set("0:00 / 0:00")
        self.root.after(500, self.update_player_ui)

    def on_use_external_player_toggle(self):
        if self.use_external_player_var.get():
            self.player_time_var.set("VLC direkte")
            self.btn_player_play_pause.config(state="disabled")
            self.btn_player_stop.config(state="disabled")
        else:
            if not self.external_playback:
                self.player_time_var.set("0:00 / 0:00")
            self.btn_player_play_pause.config(state="normal")
            self.btn_player_stop.config(state="normal")

    # -------------------- Lagring og eksport --------------------
    def row_key(self, row):
        code = (row.get("ParticipantCode") or "").strip()
        if code:
            return f"code:{code}"
        given = (row.get("GivenName") or "").strip()
        family = (row.get("FamilyName") or "").strip()
        event = (row.get("Event") or "").strip()
        return f"name:{given}|{family}|{event}"

    def load_legacy_order(self, order, path):
        # Gamle rekkefølge-filer: sorterer utøverne som allerede er lastet inn.
        buckets = {}
        for row in self.rows:
            buckets.setdefault(self.row_key(row), []).append(row)
        new_rows = []
        for item in order:
            if isinstance(item, dict):
                if item.get("type") == "pause":
                    continue
                key = item.get("key")
            else:
                key = item
            bucket = buckets.get(key)
            if bucket:
                new_rows.append(bucket.pop(0))
            elif isinstance(item, dict) and (item.get("given") or item.get("family")):
                new_rows.append(new_participant(item.get("given"), item.get("family"), item.get("club")))
        for bucket in buckets.values():
            new_rows.extend(bucket)
        self.rows = new_rows
        self.match_missing_music()
        self.relayout()
        self.refresh_schedule()
        self.log(f"Lastet gammel rekkefølge-fil: {path}")

    def generate_files(self):
        if not self.rows:
            messagebox.showerror("Feil", "Ingen data lastet.")
            return
        out_dir = self.output_dir()
        base_name = self.zip_path.stem if self.zip_path else "Deltakerliste"
        if self.var_excel.get():
            generate_excel(self.rows, str(out_dir / f"{base_name}.xlsx"), self.log)
        if self.var_html.get():
            generate_html(self.rows, str(out_dir / f"{base_name}.html"), base_name, self.log)
        if self.var_pdf.get():
            generate_pdf(self.rows, str(out_dir / f"{base_name}.pdf"), base_name, self.log)
        self.show_folder_link(out_dir, "Rapport ferdig", "Rapporten er lagret i:")

    def startliste_title(self):
        # Returnerer (tittel, filnavn uten endelse), eller None etter å ha vist feilen.
        settings = self.read_schedule_settings()
        if isinstance(settings, str):
            messagebox.showerror("Feil", settings)
            return None
        if not self.entries:
            messagebox.showwarning("Info", "Ingen utøvere i tidsplanen.")
            return None
        date_obj = parse_date_ddmmyy(self.start_date_var.get())
        if not date_obj:
            messagebox.showerror("Feil", "Ugyldig dato. Bruk DD.MM.ÅÅ.")
            return None
        location = self.location_var.get().strip()
        title = " ".join(
            part for part in (self.event_name_var.get().strip(), location, format_date_long(date_obj)) if part
        )
        return title, sanitize_filename(f"Startliste_{location}_{date_obj:%Y-%m-%d}")

    def export_excel(self):
        result = self.startliste_title()
        if not result:
            return
        title, base_name = result
        path = filedialog.asksaveasfilename(
            title="Eksporter startliste til Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"{base_name}.xlsx",
            initialdir=str(self.output_dir()),
        )
        if not path:
            return
        if generate_startliste_excel(self.entries, path, title, self.log):
            self.show_folder_link(Path(path).parent, "Startliste eksportert", "Excel-fila er lagret i:")
        else:
            messagebox.showerror("Feil", "Kunne ikke skrive Excel-fila. Er den åpen i Excel?")

    def generate_startliste(self):
        result = self.startliste_title()
        if not result:
            return
        title, base_name = result
        out_dir = self.output_dir()
        ok_excel = generate_startliste_excel(self.entries, str(out_dir / f"{base_name}.xlsx"), title, self.log)
        ok_pdf = generate_startliste_pdf(self.entries, str(out_dir / f"{base_name}.pdf"), title, self.log)
        if ok_excel or ok_pdf:
            self.show_folder_link(out_dir, "Startliste ferdig", "Startlisten er lagret i:")

    def generate_playlist_only(self):
        if not self.music_zip:
            messagebox.showerror("Feil", "Velg en musikk-zip først.")
            return
        active = self.active_rows()
        if not active:
            messagebox.showwarning("Info", "Ingen utøvere i listen.")
            return
        out_dir = self.output_dir()
        generate_vlc_playlist(active, out_dir, self.music_zip.stem, self.music_zip, self.log)
        self.show_folder_link(out_dir, "Spilleliste ferdig", "Spillelisten er lagret i:")

    def show_folder_link(self, folder_path, title, text):
        win = tk.Toplevel(self.root, bg=BG)
        win.title(title)
        win.resizable(False, False)
        frame = ttk.Frame(win, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=text).pack(anchor="w")
        path_text = str(folder_path)
        link = ttk.Label(frame, text=path_text, foreground=ACCENT, cursor="hand2")
        link.pack(anchor="w", pady=(4, 12))

        def open_folder(event=None):
            try:
                os.startfile(path_text)
            except Exception as exc:
                messagebox.showerror("Feil", f"Kunne ikke åpne mappe: {exc}")

        link.bind("<Button-1>", open_folder)
        btns = ttk.Frame(frame)
        btns.pack(fill="x")
        ttk.Button(btns, text="Åpne mappe", style="Accent.TButton", command=open_folder).pack(side="right")
        ttk.Button(btns, text="Lukk", command=win.destroy).pack(side="right", padx=(0, 8))

    def show_about(self):
        version = get_version()
        now = datetime.now()
        month_year = format_date_long(now.date()).split(" ", 1)[1]

        about = tk.Toplevel(self.root, bg=BG)
        about.title("Om")
        about.resizable(False, False)
        frame = ttk.Frame(about, padding=20)
        frame.pack(fill="both", expand=True)

        logo_path = Path(__file__).resolve().parent / "Lil Logo.jpg"
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
                pass

        ttk.Label(frame, text=APP_NAME, style="Title.TLabel").pack()
        ttk.Label(frame, text="Loddefjord IL Kunstløp", style="Dim.TLabel").pack()
        ttk.Label(
            frame,
            text=(
                "Setter opp tidsplan for stevnet og spiller av musikken underveis. "
                "Leser deltakerlister fra isonen, eller lister du lager selv."
            ),
            wraplength=360,
            justify="center",
        ).pack(pady=(10, 2))
        ttk.Label(frame, text=f"Revisjon: {version} • {month_year}", style="Dim.TLabel").pack(pady=(6, 0))


def main():
    root = tk.Tk()
    App(root)
    root.update_idletasks()
    width = int(root.winfo_screenwidth() * 0.85)
    height = int(root.winfo_screenheight() * 0.85)
    root.geometry(f"{width}x{height}")
    root.mainloop()


if __name__ == "__main__":
    main()

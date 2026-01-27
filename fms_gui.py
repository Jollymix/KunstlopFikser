import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
import subprocess


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
    except Exception:
        log("Mangler openpyxl. Installer med: pip install openpyxl")
        return False

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Participants"

    headers = [
        "PrintName",
        "GivenName",
        "FamilyName",
        "Gender",
        "Organisation",
        "ParticipantCode",
        "Event",
        "EntryOrder",
        "Music1",
        "Music2",
        "Club1",
        "Club2",
        "ElementsFree",
        "ElementsShort",
    ]
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])

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
        "GivenName",
        "FamilyName",
        "Gender",
        "Organisation",
        "ParticipantCode",
        "Event",
        "EntryOrder",
        "Music1",
        "Music2",
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
        cells = "".join(f"<td>{esc(row.get(h, ''))}</td>" for h in headers)
        rows_html.append(f"<tr>{cells}</tr>")

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
        "GivenName",
        "FamilyName",
        "Gender",
        "Organisation",
        "Event",
        "EntryOrder",
        "Music1",
        "Music2",
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
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
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


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("FSM Data Dekoder")
        self.rows = []
        self.zip_path = None

        base_dir = Path(__file__).resolve().parent
        self.folder_var = tk.StringVar(value=str(base_dir))

        folder_frame = tk.Frame(root)
        folder_frame.pack(fill="x", padx=10, pady=6)
        tk.Label(folder_frame, text="Mappe med zip:").pack(side="left")
        tk.Entry(folder_frame, textvariable=self.folder_var, width=60).pack(
            side="left", padx=6
        )
        tk.Button(folder_frame, text="Velg mappe", command=self.choose_folder).pack(
            side="left"
        )

        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", padx=10, pady=4)
        tk.Button(btn_frame, text="Les zip", command=self.read_zip).pack(side="left")
        tk.Button(btn_frame, text="?", width=3, command=self.show_about).pack(
            side="right"
        )

        self.log_widget = ScrolledText(root, height=12, wrap="word")
        self.log_widget.pack(fill="both", expand=True, padx=10, pady=6)

        out_frame = tk.LabelFrame(root, text="Lag filer")
        out_frame.pack(fill="x", padx=10, pady=6)
        self.var_pdf = tk.BooleanVar(value=False)
        self.var_excel = tk.BooleanVar(value=True)
        self.var_html = tk.BooleanVar(value=True)
        self.chk_pdf = tk.Checkbutton(out_frame, text="PDF", variable=self.var_pdf)
        self.chk_excel = tk.Checkbutton(out_frame, text="Excel", variable=self.var_excel)
        self.chk_html = tk.Checkbutton(out_frame, text="HTML", variable=self.var_html)
        self.chk_pdf.pack(side="left", padx=6)
        self.chk_excel.pack(side="left", padx=6)
        self.chk_html.pack(side="left", padx=6)
        self.btn_generate = tk.Button(
            out_frame, text="Lag filer", command=self.generate_files, state="disabled"
        )
        self.btn_generate.pack(side="right", padx=6)

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

        self.log(f"Leser zip: {self.zip_path.name}")
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
                    self.rows.extend(parse_competition(xml_text, self.log))

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

        frame = tk.Frame(about, padx=16, pady=12)
        frame.pack(fill="both", expand=True)

        logo_path = Path(__file__).resolve().parent / "Lil Logo.jpg"
        logo_label = None
        if logo_path.exists():
            try:
                from PIL import Image, ImageTk

                img = Image.open(logo_path)
                img.thumbnail((180, 180))
                photo = ImageTk.PhotoImage(img)
                logo_label = tk.Label(frame, image=photo)
                logo_label.image = photo
                logo_label.pack(pady=(0, 8))
            except Exception:
                logo_label = None

        tk.Label(
            frame,
            text="Loddefjord IL Kunstløp",
            font=("Segoe UI", 12, "bold"),
        ).pack()
        tk.Label(
            frame,
            text=(
                "Programmet leser XML fra FSM-zip og lager "
                "Excel/HTML/PDF-utskrifter."
            ),
            wraplength=360,
            justify="center",
        ).pack(pady=(6, 2))
        tk.Label(
            frame,
            text=f"Revisjon: {version} • {month_year}",
            fg="#555555",
        ).pack(pady=(4, 0))


def main():
    root = tk.Tk()
    app = App(root)
    root.geometry("900x600")
    root.mainloop()


if __name__ == "__main__":
    main()

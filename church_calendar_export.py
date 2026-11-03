#!/usr/bin/env python3
"""Download a ChurchSuite iCal feed for a chosen date range and export it to Excel.

Keep config.json private: it contains the ChurchSuite calendar feed URL.
"""

import json
import threading
from datetime import date, datetime, time, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import requests
from icalendar import Calendar
import recurring_ical_events
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "config.json"


def load_calendar_url():
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"Could not find {CONFIG_FILE.name} next to this program."
        )
    with CONFIG_FILE.open("r", encoding="utf-8") as f:
        config = json.load(f)

    url = config.get("calendar_url", "").strip()
    if not url:
        raise ValueError(
            'config.json must contain a "calendar_url" value.'
        )
    if not url.lower().startswith("https://"):
        raise ValueError("The calendar URL should start with https://.")
    return url


def as_excel_datetime(value):
    """Convert an iCalendar date or datetime to an Excel-friendly value."""
    if isinstance(value, datetime):
        # Excel does not support timezone-aware datetimes.
        if value.tzinfo is not None:
            value = value.replace(tzinfo=None)
        return value
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    return None


def event_sort_key(event):
    start = event.get("DTSTART")
    if start is None:
        return datetime.max
    return as_excel_datetime(start.dt) or datetime.max


def download_events(start_date, end_date, status_callback):
    url = load_calendar_url()
    status_callback("Downloading calendar…")

    response = requests.get(url, timeout=(15, 60))
    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "")
    if not response.content:
        raise ValueError("ChurchSuite returned an empty calendar.")

    status_callback("Reading calendar events…")
    calendar = Calendar.from_ical(response.content)

    # Use an exclusive upper boundary so events on the selected end date
    # are included, including all-day events.
    range_start = datetime.combine(start_date, time.min)
    range_end = datetime.combine(end_date + timedelta(days=1), time.min)

    expanded = recurring_ical_events.of(calendar).between(range_start, range_end)

    results = []
    for component in expanded:
        if component.name != "VEVENT":
            continue

        start_prop = component.get("DTSTART")
        if start_prop is None:
            continue

        start_value = start_prop.dt
        start_dt = as_excel_datetime(start_value)

        # Filter using the event's local calendar date.
        event_date = start_value.date() if isinstance(start_value, datetime) else start_value
        if event_date < start_date or event_date > end_date:
            continue

        end_prop = component.get("DTEND")
        end_dt = as_excel_datetime(end_prop.dt) if end_prop else None

        summary = str(component.get("SUMMARY", ""))
        location = str(component.get("LOCATION", ""))
        description = str(component.get("DESCRIPTION", ""))
        uid = str(component.get("UID", ""))

        all_day = not isinstance(start_value, datetime)

        results.append({
            "date": start_dt,
            "start": start_dt if not all_day else None,
            "end": end_dt if not all_day else None,
            "all_day": "Yes" if all_day else "No",
            "event": summary,
            "location": location,
            "description": description,
            "uid": uid,
        })

    results.sort(key=lambda row: (
        row["date"] or datetime.max,
        row["start"] or datetime.min,
        row["event"].casefold()
    ))
    return results


def export_to_excel(events, output_path, start_date, end_date):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Calendar Events"

    sheet.append([
        "Date", "Start Time", "End Time", "All Day",
        "Event", "Location", "Description", "ChurchSuite UID"
    ])

    for event in events:
        sheet.append([
            event["date"],
            event["start"],
            event["end"],
            event["all_day"],
            event["event"],
            event["location"],
            event["description"],
            event["uid"],
        ])

    # Basic readable formatting.
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center")

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for cell in sheet["A"][1:]:
        cell.number_format = "ddd dd mmm yyyy"
    for col in ("B", "C"):
        for cell in sheet[col][1:]:
            cell.number_format = "hh:mm"

    widths = {
        "A": 20, "B": 13, "C": 13, "D": 12,
        "E": 34, "F": 28, "G": 55, "H": 38
    }
    for col, width in widths.items():
        sheet.column_dimensions[col].width = width

    for row in sheet.iter_rows(min_row=2):
        row[6].alignment = Alignment(wrap_text=True, vertical="top")

    sheet.row_dimensions[1].height = 24
    workbook.properties.title = (
        f"ChurchSuite calendar {start_date:%d %b %Y} to {end_date:%d %b %Y}"
    )
    workbook.save(output_path)


class CalendarExporter:
    def __init__(self, root):
        self.root = root
        self.root.title("ChurchSuite Calendar Export")
        self.root.geometry("560x300")
        self.root.minsize(500, 280)

        frame = ttk.Frame(root, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame, text="ChurchSuite Calendar Export",
            font=("Helvetica", 16, "bold")
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 18))

        today = date.today()
        default_start = date(today.year, today.month, 1)
        default_end = default_start + timedelta(days=89)

        ttk.Label(frame, text="Start date (YYYY-MM-DD):").grid(
            row=1, column=0, sticky="w", pady=6
        )
        self.start_entry = ttk.Entry(frame, width=18)
        self.start_entry.grid(row=1, column=1, sticky="w", pady=6)
        self.start_entry.insert(0, default_start.isoformat())

        ttk.Label(frame, text="End date (YYYY-MM-DD):").grid(
            row=2, column=0, sticky="w", pady=6
        )
        self.end_entry = ttk.Entry(frame, width=18)
        self.end_entry.grid(row=2, column=1, sticky="w", pady=6)
        self.end_entry.insert(0, default_end.isoformat())

        self.export_button = ttk.Button(
            frame, text="Download calendar and create Excel file",
            command=self.start_export
        )
        self.export_button.grid(
            row=3, column=0, columnspan=3, sticky="ew", pady=(20, 12)
        )

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(
            frame, textvariable=self.status_var, wraplength=500
        ).grid(row=4, column=0, columnspan=3, sticky="w")

        ttk.Label(
            frame,
            text=f"Calendar URL is read from {CONFIG_FILE.name} beside this program.",
            foreground="gray"
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(18, 0))

    def set_status(self, message):
        self.root.after(0, self.status_var.set, message)

    def start_export(self):
        try:
            start_date = date.fromisoformat(self.start_entry.get().strip())
            end_date = date.fromisoformat(self.end_entry.get().strip())
        except ValueError:
            messagebox.showerror(
                "Invalid date",
                "Enter dates in YYYY-MM-DD format, for example 2027-01-01."
            )
            return

        if end_date < start_date:
            messagebox.showerror(
                "Invalid date range", "The end date must be on or after the start date."
            )
            return

        default_name = (
            f"ChurchSuite Calendar {start_date:%Y-%m-%d} "
            f"to {end_date:%Y-%m-%d}.xlsx"
        )
        output_path = filedialog.asksaveasfilename(
            title="Save calendar spreadsheet",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")]
        )
        if not output_path:
            return

        self.export_button.config(state="disabled")
        self.set_status("Starting…")

        thread = threading.Thread(
            target=self.run_export,
            args=(start_date, end_date, output_path),
            daemon=True
        )
        thread.start()

    def run_export(self, start_date, end_date, output_path):
        try:
            events = download_events(start_date, end_date, self.set_status)
            self.set_status("Creating Excel workbook…")
            export_to_excel(events, output_path, start_date, end_date)
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Export complete",
                    f"Exported {len(events)} event(s) to:\n{output_path}"
                )
            )
            self.set_status(f"Complete — {len(events)} event(s) exported.")
        except requests.exceptions.RequestException as exc:
            self.root.after(
                0, lambda: messagebox.showerror(
                    "Download failed",
                    "Could not download the ChurchSuite calendar.\n\n"
                    f"Details: {exc}\n\n"
                    "Check that the feed URL in config.json is correct and "
                    "that this Mac has an internet connection."
                )
            )
            self.set_status("Download failed.")
        except Exception as exc:
            self.root.after(
                0, lambda: messagebox.showerror(
                    "Export failed", f"{type(exc).__name__}: {exc}"
                )
            )
            self.set_status("Export failed.")
        finally:
            self.root.after(0, lambda: self.export_button.config(state="normal"))


if __name__ == "__main__":
    root = tk.Tk()
    app = CalendarExporter(root)
    root.mainloop()

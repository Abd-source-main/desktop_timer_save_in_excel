"""
Transparent desktop timer widget — circular button edition.

- Floats as a single transparent CIRCLE (no window frame).
- HOVER the circle  -> shows the elapsed time next to it.
- Quick CLICK        -> start / stop the timer (a "period").
- PRESS & HOLD drag  -> move the circle anywhere on screen.
- RIGHT-CLICK        -> menu: toggle "Save to Excel", check the clock, or Exit.
- Stopping (or exiting) records a period; if "Save to Excel" is on it is
  written to time_log.xlsx, ONE SHEET PER MONTH named year-then-month
  ("2026-1st", "2026-2nd" ... "2026-12th"), one row per date inside it:
      DATE (full, YYYY-MM-DD) | Total Hours (decimal) | Total (H:MM) |
      Period 1 | Period 2 | ...

Dates come from the TIMEZONE set below (Asia/Baghdad), not from the PC's own
timezone, so the log stays correct even if Windows is set wrong. If the two
disagree the widget warns you at start-up.

Run with pythonw.exe so no console appears (use the Desktop shortcut).
"""

import os
import re
import time
import datetime
import tkinter as tk
from tkinter import messagebox

from openpyxl import Workbook, load_workbook

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(SCRIPT_DIR, "time_log.xlsx")

# The zone periods are dated in. Change this one line to move timezone
# (any IANA name, e.g. "Europe/London", "Asia/Dubai").
TIMEZONE = "Asia/Baghdad"
# How far the PC's UTC offset may drift from that zone before we warn.
CLOCK_TOLERANCE_SECONDS = 60

# Any pixel drawn in this exact colour becomes fully transparent (Windows).
TRANSPARENT_KEY = "#010101"

# How far the pointer must move (pixels) before a press counts as a DRAG
# rather than a CLICK.
DRAG_THRESHOLD = 5

# Circle geometry.
CIRCLE = 48            # diameter
PAD = 6               # padding around circle inside its canvas
CANVAS = CIRCLE + PAD * 2
WIN_W = 180            # total window width (leaves room for the hover time)
WIN_H = CANVAS

COL_IDLE = "#2d7d46"      # green
COL_IDLE_HI = "#3a9a58"
COL_RUN = "#b3402f"       # red
COL_RUN_HI = "#cc4a37"

# Sheet layout: A Date | B Total Hours | C Total (H:MM) | D.. periods
HEADERS = ["Date", "Total Hours", "Total (H:MM)", "Periods ->"]
HM_COL = 3
PERIOD_COL = 4


def fmt_duration(seconds):
    seconds = int(round(seconds))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def fmt_hm(seconds):
    """Hours + minutes, rounded to the nearest minute: 5400 -> "1:30".

    Rounds half UP; round() would round 30s down to 0:00 (banker's rounding).
    """
    minutes = int((max(seconds, 0) + 30) // 60)
    h, m = divmod(minutes, 60)
    return f"{h}:{m:02d}"


def fmt_offset(delta):
    """A timedelta UTC offset as "UTC+03:00"."""
    if delta is None:
        return "UTC+??:??"
    total = int(delta.total_seconds())
    sign = "-" if total < 0 else "+"
    h, m = divmod(abs(total) // 60, 60)
    return f"UTC{sign}{h:02d}:{m:02d}"


# ----------------------------------------------------------------------------
# Timezone
# ----------------------------------------------------------------------------
def get_timezone(name=None):
    """-> (name, tzinfo or None, error string or None)."""
    name = name or TIMEZONE
    try:
        from zoneinfo import ZoneInfo
        return name, ZoneInfo(name), None
    except Exception as exc:            # unknown zone, or no tzdata installed
        return name, None, f"{type(exc).__name__}: {exc}"


def now_local(name=None):
    """Current time in the configured zone (falls back to the PC clock)."""
    _, tz, _ = get_timezone(name)
    if tz is None:
        return datetime.datetime.now()
    return datetime.datetime.now(tz)


def clock_status(name=None, reference=None):
    """Compare the PC's UTC offset with the configured timezone's offset.

    -> dict(ok, kind, timezone, expected, actual, delta_seconds, message)
    """
    name, tz, err = get_timezone(name)
    ref = reference or datetime.datetime.now(datetime.timezone.utc)

    if tz is None:
        return {
            "ok": False,
            "kind": "timezone-unavailable",
            "timezone": name,
            "expected": None,
            "actual": ref.astimezone().utcoffset(),
            "delta_seconds": None,
            "message": (
                f"Timezone {name!r} could not be loaded ({err}).\n\n"
                "Times will be recorded using the PC clock instead.\n"
                "Fix: run  pip install tzdata  , or correct TIMEZONE at the "
                "top of timer_widget.pyw."
            ),
        }

    expected = ref.astimezone(tz).utcoffset()
    actual = ref.astimezone().utcoffset()
    delta = (actual - expected).total_seconds()

    if abs(delta) <= CLOCK_TOLERANCE_SECONDS:
        return {
            "ok": True,
            "kind": "in-sync",
            "timezone": name,
            "expected": expected,
            "actual": actual,
            "delta_seconds": delta,
            "message": (
                f"Clock is in sync with {name} ({fmt_offset(expected)}).\n"
                f"Local time there: {ref.astimezone(tz):%Y-%m-%d %H:%M:%S}"
            ),
        }

    hours_off = delta / 3600.0
    return {
        "ok": False,
        "kind": "offset-mismatch",
        "timezone": name,
        "expected": expected,
        "actual": actual,
        "delta_seconds": delta,
        "message": (
            "This PC's clock does not match the timezone this timer logs in.\n\n"
            f"  TIMEZONE : {name}  ({fmt_offset(expected)})\n"
            f"  This PC  : {fmt_offset(actual)}\n"
            f"  Difference: {hours_off:+.2f} h\n\n"
            f"  {name} now : {ref.astimezone(tz):%Y-%m-%d %H:%M:%S}\n"
            f"  This PC now: {ref.astimezone():%Y-%m-%d %H:%M:%S}\n\n"
            f"Periods are logged using {name}, so the sheet stays correct — "
            "but fix the Windows clock/timezone if this is unexpected."
        ),
    }


# ----------------------------------------------------------------------------
# Widget
# ----------------------------------------------------------------------------
class TimerWidget:
    def __init__(self):
        self.running = False
        self.start_epoch = None
        self.elapsed = 0.0

        self._moved = False
        self._hover = False

        self.root = tk.Tk()
        self.root.title("Timer")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=TRANSPARENT_KEY)
        try:
            self.root.attributes("-transparentcolor", TRANSPARENT_KEY)
        except tk.TclError:
            self.root.attributes("-alpha", 0.9)

        self.root.geometry(f"{WIN_W}x{WIN_H}")

        # Time readout (hidden until hover). Sits to the LEFT of the circle.
        self.time_label = tk.Label(
            self.root, text="00:00:00", font=("Segoe UI", 13, "bold"),
            fg="#00e0a0", bg="#22222a", padx=8, pady=4,
        )
        # placed on demand in _show_time()

        # The circle itself, on a transparent canvas so only the oval shows.
        self.canvas = tk.Canvas(
            self.root, width=CANVAS, height=CANVAS,
            bg=TRANSPARENT_KEY, highlightthickness=0, bd=0, cursor="hand2",
        )
        self.canvas.place(relx=1.0, rely=0.5, anchor="e")

        self.oval = self.canvas.create_oval(
            PAD, PAD, PAD + CIRCLE, PAD + CIRCLE,
            fill=COL_IDLE, outline="#0d0d0f", width=2,
        )
        # Glyph in the centre (play / pause).
        self.glyph = self.canvas.create_text(
            CANVAS / 2, CANVAS / 2, text="▶",  # ▶
            font=("Segoe UI", 16, "bold"), fill="white",
        )

        # --- interaction bindings (on the canvas only) ---
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_motion)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Enter>", self.on_enter)
        self.canvas.bind("<Leave>", self.on_leave)
        self.canvas.bind("<Button-3>", self.show_menu)

        # Right-click context menu.
        self.save_var = tk.BooleanVar(value=True)
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label=f"Timezone: {TIMEZONE}", state="disabled")
        self.menu.add_command(label="Check clock…", command=self.check_clock)
        self.menu.add_separator()
        self.menu.add_checkbutton(label="Save to Excel", variable=self.save_var)
        self.menu.add_separator()
        self.menu.add_command(label="Exit", command=self.on_close)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.root.update_idletasks()
        self.place_bottom_right()
        self.tick()
        # Warn once the event loop is running, so the dialog is not modal
        # against a half-built window.
        self.root.after(600, self.warn_if_clock_out_of_sync)

    # ------------------------------------------------------------- placement
    def place_bottom_right(self):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = sw - WIN_W - 20
        y = sh - WIN_H - 60
        self.root.geometry(f"{WIN_W}x{WIN_H}+{x}+{y}")

    # ------------------------------------------------------- press vs. drag
    def on_press(self, event):
        self._press_x = event.x_root
        self._press_y = event.y_root
        self._win_x = self.root.winfo_x()
        self._win_y = self.root.winfo_y()
        self._moved = False

    def on_motion(self, event):
        dx = event.x_root - self._press_x
        dy = event.y_root - self._press_y
        if not self._moved and (abs(dx) > DRAG_THRESHOLD or abs(dy) > DRAG_THRESHOLD):
            self._moved = True
            self.time_label.place_forget()  # hide time while dragging
        if self._moved:
            self.root.geometry(f"+{self._win_x + dx}+{self._win_y + dy}")

    def on_release(self, event):
        if self._moved:
            self._moved = False
            if self._hover:
                self._show_time()
            return
        # A genuine click -> toggle the timer.
        self.toggle()

    # ------------------------------------------------------------- hover
    def on_enter(self, event):
        self._hover = True
        self._show_time()

    def on_leave(self, event):
        self._hover = False
        self.time_label.place_forget()

    def _show_time(self):
        self.time_label.config(text=fmt_duration(self.elapsed))
        # Sit just left of the circle.
        self.time_label.place(relx=1.0, rely=0.5, anchor="e", x=-(CANVAS + 2))

    # ------------------------------------------------------------- menu
    def show_menu(self, event):
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    # ------------------------------------------------------------- clock
    def warn_if_clock_out_of_sync(self):
        status = clock_status()
        if not status["ok"]:
            messagebox.showwarning(
                "Clock / timezone mismatch", status["message"], parent=self.root)

    def check_clock(self):
        status = clock_status()
        if status["ok"]:
            messagebox.showinfo("Clock", status["message"], parent=self.root)
        else:
            messagebox.showwarning(
                "Clock / timezone mismatch", status["message"], parent=self.root)

    # ------------------------------------------------------------- timer
    def toggle(self):
        self.stop() if self.running else self.start()

    def start(self):
        self.running = True
        self.start_epoch = time.time()
        self.elapsed = 0.0
        self.canvas.itemconfig(self.oval, fill=COL_RUN)
        self.canvas.itemconfig(self.glyph, text="⏸")  # ⏸

    def stop(self):
        if not self.running:
            return
        self.elapsed = time.time() - self.start_epoch
        self.running = False
        self.canvas.itemconfig(self.oval, fill=COL_IDLE)
        self.canvas.itemconfig(self.glyph, text="▶")  # ▶

        if self.save_var.get() and self.elapsed >= 1:
            try:
                save_period(self.elapsed)
            except Exception as exc:
                # Never fail silently: under pythonw there is no console, so a
                # print() here would be lost (and would itself raise).
                messagebox.showerror(
                    "Could not save period",
                    f"{fmt_duration(self.elapsed)} was NOT written to\n"
                    f"{EXCEL_PATH}\n\n{type(exc).__name__}: {exc}\n\n"
                    "Close the file in Excel and try again.",
                    parent=self.root,
                )

        self.elapsed = 0.0
        if self._hover:
            self._show_time()

    def tick(self):
        if self.running:
            self.elapsed = time.time() - self.start_epoch
            if self._hover:
                self.time_label.config(text=fmt_duration(self.elapsed))
        self.root.after(200, self.tick)

    # ------------------------------------------------------------- close
    def on_close(self):
        if self.running:
            self.stop()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ----------------------------------------------------------------------------
# Excel writing — one sheet per month
# ----------------------------------------------------------------------------
def ordinal(n):
    """1 -> "1st", 2 -> "2nd", 11 -> "11th"."""
    if 11 <= n % 100 <= 13:
        return f"{n}th"
    suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def sheet_name(day):
    """Sheet title for a date: 2026-05-04 -> "2026-5th"."""
    return f"{day.year}-{ordinal(day.month)}"


def parse_sheet_name(title):
    """(year, month) for a sheet title, or None if it is not one of ours."""
    if not isinstance(title, str):
        return None
    text = title.strip().lower()
    match = re.fullmatch(r"(\d{4})-(\d{1,2})(?:st|nd|rd|th)", text)
    if not match:
        return None
    year, month = int(match.group(1)), int(match.group(2))
    if not 1 <= month <= 12:
        return None
    # Reject a wrong suffix ("2026-1th") rather than silently accepting it.
    if f"{year}-{ordinal(month)}" != text:
        return None
    return year, month


def parse_date(value):
    """Excel cells may hand back a str or a datetime; normalise to a date."""
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if value in (None, ""):
        return None
    try:
        return datetime.date.fromisoformat(str(value).strip()[:10])
    except ValueError:
        return None


def parse_duration(value):
    """A period cell back into seconds: "01:30:00" -> 5400.0. None if unreadable."""
    if isinstance(value, datetime.timedelta):
        return value.total_seconds()
    if isinstance(value, datetime.time):
        return value.hour * 3600 + value.minute * 60 + value.second
    if value in (None, ""):
        return None
    match = re.fullmatch(r"(\d+):([0-5]?\d):([0-5]?\d)", str(value).strip())
    if not match:
        return None
    h, m, s = (int(g) for g in match.groups())
    return float(h * 3600 + m * 60 + s)


def ensure_hm_column(ws):
    """Upgrade a sheet written before the H:MM column existed."""
    if ws.cell(row=1, column=HM_COL).value == HEADERS[HM_COL - 1]:
        return
    ws.insert_cols(HM_COL)
    ws.cell(row=1, column=HM_COL, value=HEADERS[HM_COL - 1])
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value in (None, ""):
            continue
        total = ws.cell(row=row, column=2).value or 0
        ws.cell(row=row, column=HM_COL, value=fmt_hm(float(total) * 3600.0))


def get_sheet(wb, day):
    """The sheet for a date's month, created (in calendar order) if missing."""
    name = sheet_name(day)
    if name in wb.sheetnames:
        ws = wb[name]
        ensure_hm_column(ws)
        return ws

    key = (day.year, day.month)
    position = 0
    for title in wb.sheetnames:
        other = parse_sheet_name(title)
        if other is not None and other < key:
            position += 1
    ws = wb.create_sheet(title=name, index=position)
    ws.append(HEADERS)
    return ws


def find_date_row(ws, day):
    for row in range(2, ws.max_row + 1):
        if parse_date(ws.cell(row=row, column=1).value) == day:
            return row
    return None


def upsert_day(ws, day, hours, periods):
    """Add hours + period cells to `day`'s row, creating the row if needed."""
    target_row = find_date_row(ws, day)
    if target_row is None:
        target_row = ws.max_row + 1
        ws.cell(row=target_row, column=1, value=day.isoformat())
        previous_hours = 0.0
    else:
        previous_hours = float(ws.cell(row=target_row, column=2).value or 0)

    col = PERIOD_COL
    while ws.cell(row=target_row, column=col).value not in (None, ""):
        col += 1
    for period in periods:
        ws.cell(row=target_row, column=col, value=period)
        col += 1

    # Re-total from the period cells. They hold whole seconds, so the sum is
    # exact; carrying forward the ROUNDED total in column B instead would drift
    # by a second or two more with every period added to the day.
    total_seconds = 0.0
    exact = False
    for period_col in range(PERIOD_COL, col):
        seconds = parse_duration(ws.cell(row=target_row, column=period_col).value)
        if seconds is None:          # hand-edited cell we cannot read
            exact = False
            break
        total_seconds += seconds
        exact = True
    if not exact:                    # fall back to the running total
        total_seconds = (previous_hours + hours) * 3600.0

    ws.cell(row=target_row, column=2, value=round(total_seconds / 3600.0, 3))
    ws.cell(row=target_row, column=HM_COL, value=fmt_hm(total_seconds))
    return target_row


def looks_like_log_sheet(ws):
    return ws.cell(row=1, column=1).value == HEADERS[0]


def migrate_legacy_sheets(wb):
    """Re-file rows from older sheet layouts into "<year>-<month>" sheets.

    Handles both the original single "Time Log" sheet and the year-less
    month sheets ("5th"), splitting the latter by each row's own year.
    Sheets that are not ours are left untouched.
    """
    for title in list(wb.sheetnames):
        if parse_sheet_name(title) is not None:
            continue
        ws = wb[title]
        if not looks_like_log_sheet(ws):
            continue

        ensure_hm_column(ws)
        rows = []
        for values in ws.iter_rows(min_row=2, values_only=True):
            day = parse_date(values[0] if values else None)
            if day is None:
                continue
            raw_hours = values[1] if len(values) > 1 else 0
            try:
                hours = float(raw_hours or 0)
            except (TypeError, ValueError):
                hours = 0.0
            periods = [v for v in values[PERIOD_COL - 1:] if v not in (None, "")]
            rows.append((day, hours, periods))

        wb.remove(ws)
        for day, hours, periods in rows:
            upsert_day(get_sheet(wb, day), day, hours, periods)


def save_period(seconds, when=None):
    """Append a period to today's row, in this month's sheet.

    Workbook layout:
        sheet "2026-1st".."2026-12th"  — one per month of each year
        A: Date (YYYY-MM-DD) | B: Total Hours | C: Total (H:MM) | D..: Periods
    """
    when = now_local() if when is None else when
    day = when.date() if isinstance(when, datetime.datetime) else when
    hours = seconds / 3600.0

    if os.path.exists(EXCEL_PATH):
        wb = load_workbook(EXCEL_PATH)
        migrate_legacy_sheets(wb)
    else:
        wb = Workbook()
        wb.remove(wb.active)        # drop the default empty "Sheet"

    ws = get_sheet(wb, day)
    upsert_day(ws, day, hours, [fmt_duration(seconds)])
    wb.save(EXCEL_PATH)


if __name__ == "__main__":
    TimerWidget().run()

# Desk Timer

A transparent, frameless timer that floats on your desktop as a single circle and writes every session you track straight into an Excel workbook — one sheet per month, one row per day.

No window frame, no taskbar entry, no console. Just a small circle that sits above your other windows until you need it.

```
                                    ┌──────────┐
                          hover →   │ 01:24:07 │  ●   ← click to start / stop
                                    └──────────┘
```

---

## Features

- **One-click tracking.** Click the circle to start a period, click again to stop it. Green ▶ when idle, red ⏸ while running.
- **Out of the way.** Frameless, always-on-top, and transparent apart from the circle itself. Drag it anywhere on screen.
- **Elapsed time on hover.** The running total appears beside the circle only while your pointer is over it.
- **Excel logging.** Each stopped period is appended to `time_log.xlsx`, grouped into a sheet per month with a per-day total in both decimal hours and `H:MM`.
- **Timezone-correct dates.** Rows are dated in a timezone you choose, not whatever the PC clock happens to be set to.
- **Clock-sync warning.** On launch the widget compares the PC's UTC offset against that timezone and warns you if they disagree.
- **Nothing fails silently.** If the workbook can't be written — usually because it's open in Excel — you get a dialog naming the file, not a lost period.

---

## Requirements

| | |
|---|---|
| OS | Windows (the transparent-circle effect uses a Windows-only Tk feature; elsewhere the window falls back to a translucent rectangle) |
| Python | 3.9 or newer (needs `zoneinfo`) |
| Packages | `openpyxl`, plus `tzdata` on Windows |

```bash
pip install openpyxl tzdata
```

`tkinter` ships with the standard python.org installer. If `import tkinter` fails, re-run the installer and enable *tcl/tk and IDLE*.

---

## Running it

```bash
pythonw timer_widget.pyw
```

Use `pythonw.exe` rather than `python.exe` so no console window appears behind the widget.

### Desktop shortcut

Run this once to drop a **Desk Timer** shortcut on your Desktop that launches the widget windowless:

```bash
python create_shortcut.py
```

It asks Windows for your real Desktop folder, so it works even when OneDrive has redirected it.

---

## Controls

| Action | Result |
|---|---|
| **Hover** the circle | Shows the elapsed time beside it |
| **Click** | Start / stop the timer (one *period*) |
| **Press and drag** | Move the circle anywhere on screen |
| **Right-click** | Menu: current timezone, *Check clock…*, *Save to Excel* toggle, *Exit* |

A press only counts as a drag once the pointer moves more than 5 pixels, so a slightly shaky click still starts the timer.

Stopping the timer records the period. Exiting while the timer is still running records it too, so you don't lose the session by closing the widget.

---

## What gets logged

Periods go to `time_log.xlsx`, created next to `timer_widget.pyw` on the first save.

### Sheets

One sheet per month, named year-first: `2026-1st`, `2026-2nd`, … `2026-12th`, `2027-1st`. Sheets are kept in chronological order, and each year gets its own — May 2026 and May 2027 never share a sheet.

### Rows

One row per date, inside that date's month sheet:

| Date | Total Hours | Total (H:MM) | Periods → | | |
|---|---|---|---|---|---|
| 2026-09-05 | 1.512 | 1:31 | 01:00:00 | 00:30:00 | 00:00:45 |
| 2026-09-06 | 3.25 | 3:15 | 02:00:00 | 01:15:00 | |

- **Date** — the full date, `YYYY-MM-DD`, so rows stay sortable and unambiguous.
- **Total Hours** — decimal hours, the summable one if you want to chart or `SUM()` it.
- **Total (H:MM)** — the same total in hours and minutes, rounded to the nearest minute.
- **Periods →** — one cell per period, `HH:MM:SS`, appended left to right.

Both totals are re-summed from the period cells on every write, so they always match the periods you can see and never drift as a day accumulates.

Periods shorter than one second are ignored. The `Save to Excel` toggle in the right-click menu turns logging off entirely if you just want a stopwatch.

---

## Timezone and the clock check

Dates are stamped in the timezone set at the top of `timer_widget.pyw`:

```python
TIMEZONE = "Asia/Baghdad"
```

Any IANA name works — `Europe/London`, `Asia/Dubai`, `America/New_York`. Because the date comes from that zone rather than from Windows, your log stays correct even if the PC's timezone is wrong or gets changed.

To make that visible rather than silent, the widget checks the clock automatically about half a second after launch and shows a warning if the PC's UTC offset doesn't match the configured zone — including both offsets, the difference in hours, and both wall-clock times. You can re-run the check any time from **right-click → Check clock…**, which also confirms when everything is in sync.

A separate warning appears if the timezone itself can't be loaded — a typo in `TIMEZONE`, or `tzdata` not installed.

---

## Configuration

Everything tunable lives in constants at the top of `timer_widget.pyw`:

| Constant | Default | Meaning |
|---|---|---|
| `TIMEZONE` | `"Asia/Baghdad"` | Zone that dates are stamped in |
| `CLOCK_TOLERANCE_SECONDS` | `60` | How far the PC's offset may drift before warning |
| `EXCEL_PATH` | `time_log.xlsx` beside the script | Where the log is written |
| `DRAG_THRESHOLD` | `5` | Pixels of movement before a press becomes a drag |
| `CIRCLE` | `48` | Circle diameter in pixels |
| `COL_IDLE` / `COL_RUN` | green / red | Circle colour when idle / running |

---

## Upgrading from an older version

Existing workbooks are migrated in place the first time a period is saved. Two older layouts are handled:

- a single **`Time Log`** sheet — rows are split into month sheets, and the `Total (H:MM)` column is added and backfilled;
- **year-less month sheets** (`5th`, `9th`) — renamed to `2026-5th` style, and split by each row's own year if one sheet holds more than one.

Migration is idempotent, and sheets that aren't the timer's own are left untouched — if you keep your own notes tab in that workbook, it survives.

---

## Project layout

```
timer_widget.pyw     the widget and all the Excel logic
create_shortcut.py   one-shot helper that builds the Desktop shortcut
time_log.xlsx        your log (git-ignored, created on first save)
```

---

## Notes

- The transparent circle relies on Tk's `-transparentcolor` attribute, which is Windows-only. On other platforms the widget still runs but falls back to a translucent rectangular window.
- Keep `time_log.xlsx` closed while the timer is running. Excel locks the file, and a save attempted against a locked file will pop the error dialog instead of writing.
- A period's *length* is measured as a plain elapsed-seconds difference and is unaffected by the timezone setting — that only decides which date the period is filed under.

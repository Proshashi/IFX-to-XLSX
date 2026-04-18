# IFX to XLSX Converter — Technical Documentation

Reference for developers maintaining or extending `ifx_to_xlsx_gui.py`.
For end-user instructions see [README.md](README.md).

---

## 1. Project layout

```
ifx-to-xlsx/
├── ifx_to_xlsx_gui.py     Single-file application (parser + writers + Tk GUI)
├── Convert_IFX.command    Double-clickable launcher for macOS Finder
├── README.md              End-user installation and usage guide
└── TECHNICAL.md           This file
```

The whole tool is one ~500-line Python module. There is no package
structure, no external configuration, and no persistent state.

### Runtime dependencies

| Package    | Used for                              | Source           |
|------------|---------------------------------------|------------------|
| `openpyxl` | Workbook construction, cell styling   | PyPI             |
| `tkinter`  | GUI (Tk widgets, file dialogs)        | Python stdlib    |
| `re`, `pathlib`, `threading` | Parsing / file paths / async | Python stdlib |

Python 3.8+ is required (uses f-strings with `=`, `Path`, type-friendly
constructs but no PEP 604 unions).

---

## 2. The IFX file format

`.ifx` is the ISS spectrofluorometer experiment export. The parser
treats it as a plain UTF-8 text file split by the literal token
`[Data]`:

```
ISS_Experiment_Ver_1_0
Title=...
Timestamp=...
Measurable=Intensity, unit:None
Space=Iteration, EmissionWavelength
Columns=Iteration,EmissionWavelength,Intensity,IntensityStdError,ExcitationChannel
...other Key=Value lines...
[Data]
1   400   583.9   1.2   0
1   401   554.6   1.1   0
...
```

The contract the parser relies on:

- **Header section** — every meaningful line is `Key=Value`. Order does
  not matter; unknown keys are kept in the metadata dict.
- **`[Data]` separator** — a single literal occurrence; everything
  before is header, everything after is data rows.
- **`Columns=`** — comma-separated list naming each whitespace-delimited
  column in the data block.
- **`Space=`** — comma-separated list of independent axes (the
  measurement domain). Length controls whether a pivot is possible.
- **`Measurable=`** — name of the dependent variable; must also appear
  in `Columns=`. May include a trailing `, unit:<u>` segment.
- **Per-axis lines** (e.g. `EmissionWavelength=...unit:nm,...`) — the
  unit is parsed from the `unit:<u>` substring on that header line.

If `Space=` or `Measurable=` are absent (older files), `build_pivot`
falls back to a hardcoded list of likely names (`Iteration`,
`EmissionWavelength`, `ExcitationWavelength`, `Time`, `Intensity`,
`Anisotropy`, `Polarization`, `Lifetime`).

---

## 3. Module architecture

`ifx_to_xlsx_gui.py` is organized into five layers, separated by
section banners:

```
Parsing            parse_ifx
Styling helpers    HEADER_FONT, _style_header_row, BORDER, ...
Pivot helper       build_pivot, _axis_label, _axis_unit, PivotNotPossible
Writers            _write_spectra_sheet, _write_data_sheet,
                   _write_metadata_sheet, write_mb_format,
                   write_comprehensive, convert_file
GUI                App, main
```

Each layer depends only on the ones above it. `convert_file` is the
single entry point used by both the GUI and any future CLI / batch
caller — it is pure, takes paths, and returns a path.

### Data flow

```
.ifx path
   │
   ▼
parse_ifx ──► (metadata: dict, columns: list[str], rows: list[list])
   │
   ▼
build_pivot (optional) ──► (row_label, col_label, row_vals, col_vals,
                            matrix: dict[(col_val,row_val)→value], measurable)
   │
   ▼
write_mb_format / write_comprehensive
   │
   ▼
.xlsx path
```

---

## 4. Parsing — `parse_ifx(path)`

Location: `ifx_to_xlsx_gui.py:32`.

Steps:

1. Read the file as UTF-8 with `errors="replace"` (tolerates stray
   non-UTF-8 bytes that occasionally appear in instrument exports).
2. Reject files lacking `[Data]` with a `ValueError` carrying the
   filename.
3. `partition("[Data]")` splits header from body in one pass.
4. **Header pass** — for every non-empty line containing `=`, split
   once on `=` and store the trimmed key/value in `metadata`.
5. **Columns** — split `metadata["Columns"]` on commas, trim, drop
   empties.
6. **Body pass** — for every non-empty line, split on whitespace
   (`re.split(r"\s+", s)`), then coerce each token:
   - contains `.` or `e`/`E` → `float`
   - otherwise → `int`
   - on `ValueError` → keep raw `str`

Returns `(metadata, columns, rows)` as a tuple. No validation of column
arity per row is done — pivot construction tolerates ragged rows by
indexing into the row list, and will raise `IndexError` if a row is too
short for the declared columns.

---

## 5. Pivot — `build_pivot(metadata, columns, rows)`

Location: `ifx_to_xlsx_gui.py:109`.

Converts long-format rows into a 2D matrix indexed by two axes.

### Inputs

- `metadata["Space"]` — comma-separated axes
- `metadata["Measurable"]` — value column name
- Per-axis header lines for unit extraction

### Algorithm

1. **Resolve axes.** Parse `Space=`. If empty, fall back to a
   hardcoded list of likely axis names that appear in `columns`.
2. **Resolve the value column.** Use `Measurable=`; if absent, scan
   `columns` for `Intensity`, `Anisotropy`, `Polarization`, or
   `Lifetime` in that order.
3. **Validate.** Raise `PivotNotPossible` if:
   - `len(space) != 2`
   - `measurable` is missing or not in `columns`
   - either axis is not in `columns`
4. **Pick row vs column axis.**
   - If one axis is `Iteration`, it becomes the **column** axis (each
     iteration is a series, matching MB's VBA convention).
   - Otherwise, the axis with **fewer unique values** becomes columns.
5. **Build matrix** in a single pass over rows:
   - Maintain insertion-ordered `row_vals` and `col_vals` (preserves
     measurement order).
   - Store `matrix[(col_val, row_val)] = value`.
6. **Build display labels** via `_axis_label(name, unit)`:
   - `_axis_unit` regex-matches `unit:<u>` on the axis's header line;
     `unit:None` is treated as no unit.
   - `_axis_label` inserts spaces before each capital letter
     (`EmissionWavelength` → `Emission Wavelength`) and appends
     `(unit)` if present.

Returns `(row_label, col_label, row_vals, col_vals, matrix, measurable)`.

### Why `PivotNotPossible` is its own class

It signals a recoverable, format-level limitation distinct from
`ValueError` (corrupt file) or `IOError` (filesystem). `convert_file`
uses it to decide between hard failure (MB mode) and graceful skip
(Comprehensive mode keeps Data + Metadata sheets).

---

## 6. Writers

### `_write_spectra_sheet` (`ifx_to_xlsx_gui.py:189`)

Pivoted matrix, dynamic axis labels, full numeric precision.

- Row 1 = `[row_label, col_val_1, col_val_2, ...]`, styled as header.
- Column A = row axis values, bold + center-aligned.
- Body cells use `number_format = "0.######"` for floats — preserves up
  to 6 decimals without forcing trailing zeros.
- Column A width auto-fits the label; series columns get a fixed 14ch
  width.
- `freeze_panes = "B2"` keeps headers and the wavelength column visible
  while scrolling.

### `_write_data_sheet` (`ifx_to_xlsx_gui.py:231`)

Raw long-format dump.

- Header row is the original `Columns=` list (or synthetic `Col1..ColN`
  if absent).
- All floats get `0.######`. Column widths are `max(14, header_len + 4)`.
- `freeze_panes = "A2"`.

### `_write_metadata_sheet` (`ifx_to_xlsx_gui.py:216`)

Two-column key/value table covering every parsed header field.

- Field column 36ch, value column 70ch with wrap.
- Bordered, body-font, top-left frozen.

### `write_mb_format` (`ifx_to_xlsx_gui.py:259`)

- Single sheet named after the source file (sanitized via
  `_safe_sheet_name` — strips `: \ / ? * [ ]`, truncates to 31 chars per
  Excel rules).
- Calls `build_pivot` and writes only the Spectra sheet.
- Propagates `PivotNotPossible` to the caller.

### `write_comprehensive` (`ifx_to_xlsx_gui.py:271`)

- Attempts pivot; on `PivotNotPossible` skips the Spectra sheet and
  promotes Data to the active sheet.
- Always writes Data + Metadata. Sheet order: Spectra (if any), Data,
  Metadata.

### `convert_file(ifx_path, out_path, mode)`

Public façade. Takes the destination `.xlsx` path directly (callers own
naming and overwrite policy), dispatches on `mode in {"mb",
"comprehensive"}`, and re-raises `PivotNotPossible` with a friendlier
message for MB mode. Returns the output path.

---

## 7. Styling constants

Defined once near the top of the styling section
(`ifx_to_xlsx_gui.py:70`):

| Constant      | Value                                                           |
|---------------|-----------------------------------------------------------------|
| `HEADER_FONT` | Arial bold, white text                                          |
| `HEADER_FILL` | Solid `#305496` (Excel "blue, accent 1, darker 25%")            |
| `BODY_FONT`   | Arial regular                                                   |
| `BORDER`      | Thin `#BFBFBF` on all four sides                                |

`_style_header_row(ws, row, n_cols)` applies the header styling to a
contiguous range. Pass `n_cols` explicitly when the header is shorter
than `ws.max_column` (e.g. before body rows have been appended).

---

## 8. GUI — `class App`

Location: `ifx_to_xlsx_gui.py:322`.

### Layout

A single `ttk.Frame` with an 8-row grid:

| Row | Widget                                     |
|-----|--------------------------------------------|
| 0   | Title label                                |
| 1   | Subtitle / instruction                     |
| 2   | File chooser row (Choose / count / Clear)  |
| 3   | Listbox of selected files (expandable)     |
| 4   | `LabelFrame` with mode radio buttons       |
| 5   | Output directory row (label / entry / Browse) |
| 6   | Output filename row (label / entry / `.xlsx` hint) — hidden unless 1 file selected |
| 7   | Progress bar                               |
| 8   | Status label                               |
| 9   | Convert button                             |

The output filename row is shown only when exactly one file is
selected. `refresh_list` calls `self.name_row.grid()` /
`self.name_row.grid_remove()` to toggle visibility (the row remembers
its grid options across hide/show) and pre-fills the entry with the
source file's stem. For multi-file batches the row is hidden and the
source filenames are used directly.

Resizing: row 3 and column 0 carry `weight=1`, so the file list
expands to fill available space; everything else stays pinned.

Theme: tries `style.theme_use("aqua")` for native macOS look; failure
is silent.

### State

- `self.files: list[Path]` — current selection
- `self.out_dir: StringVar` — defaults to `~/Desktop`
- `self.out_name: StringVar` — custom output stem (single-file only)
- `self.mode: StringVar` — `"mb"` or `"comprehensive"`

### Threading

`start_convert` spawns a daemon `threading.Thread` running
`_run_conversion`. The worker:

1. Iterates `self.files`. For each, calls `_resolve_out_path` to pick
   the destination (custom name when one file is selected, source stem
   otherwise; trailing `.xlsx` stripped if the user typed it).
2. If the destination already exists, prompts via
   `messagebox.askyesno`. A "no" response skips that file and records
   it in the `skipped` list — the batch continues.
3. Calls `convert_file(ifx_path, out_path, mode)` per item.
4. Updates `self.status` and `self.progress` after each file.
5. Re-enables the Convert button and shows a final
   `messagebox.showinfo` / `showwarning` summary that lists `done`,
   `skipped`, and `failed` separately.

`self.root.update_idletasks()` is called after each progress tick to
flush trivial widget updates (`StringVar` writes, `progress.config`)
from the worker — Tk tolerates these on macOS in practice.

**Modal dialogs cannot be invoked from the worker.** Calling
`messagebox.askyesno` / `showinfo` / `showwarning` directly from a
background thread raises `_tkinter.TclError` on Python 3.13+ Tk. The
`App._on_ui(fn, *args, **kwargs)` helper marshals the call onto the
main thread via `root.after(0, ...)` and blocks the worker on a
`threading.Event` until the dialog returns (or re-raises any
exception). All overwrite prompts and end-of-batch summary dialogs go
through this helper.

Errors during a single file are caught and collected in `failed`; the
loop continues so one bad file does not abort the batch.

---

## 9. Error handling matrix

| Condition                            | Where raised        | User-visible result                           |
|--------------------------------------|---------------------|-----------------------------------------------|
| Missing `[Data]` section             | `parse_ifx`         | "No [Data] section in <name>" in failure list |
| `Space` axes ≠ 2 in MB mode          | `build_pivot`       | "Cannot pivot: ... Try Comprehensive mode"    |
| `Space` axes ≠ 2 in Comprehensive    | `build_pivot`       | Silent — Spectra sheet omitted                |
| Missing `Measurable` / unknown axis  | `build_pivot`       | `PivotNotPossible` with descriptive message   |
| Output folder does not exist         | `start_convert`     | Auto-`mkdir(parents=True)`; error dialog on failure |
| No files selected                    | `start_convert`     | Warning dialog                                |
| Filesystem / permission error        | openpyxl `wb.save`  | Caught per-file, added to failure list        |

---

## 10. Building a macOS `.app`

```bash
pip3 install pyinstaller
pyinstaller --windowed --name "IFX to Excel" ifx_to_xlsx_gui.py
```

Produces `dist/IFX to Excel.app`. `--windowed` suppresses the Terminal
window on launch. The bundle is unsigned, so first launch needs
right-click → Open to bypass Gatekeeper.

There is no spec file checked in; PyInstaller's defaults are sufficient
because the script imports nothing outside `openpyxl` + stdlib and
loads no data files at runtime.

---

## 11. Extension points

### Adding a new output mode

1. Write a new `write_<name>(metadata, columns, rows, out_path)`.
2. Add the dispatch branch in `convert_file` (`ifx_to_xlsx_gui.py:303`).
3. Add a `Radiobutton` in the GUI's `fmt_frame` with a matching `value=`.

### Supporting a new IFX axis

If a new axis name appears in `Space=` and is already in `Columns=`,
no code change is needed — `build_pivot` is metadata-driven.

If you need to support **older** files that omit `Space=` or
`Measurable=`, extend the fallback lists in `build_pivot`
(`ifx_to_xlsx_gui.py:127–134`).

### Headless / CLI use

`convert_file(path, out_dir, mode)` is import-safe — the GUI is only
constructed inside `main()` under `if __name__ == "__main__"`. A CLI
wrapper can do:

```python
from pathlib import Path
from ifx_to_xlsx_gui import convert_file
convert_file(
    Path("run.ifx"),
    Path("./out/run.xlsx"),    # destination .xlsx, not a directory
    mode="comprehensive",
)
```

---

## 12. Known limitations

- **Single-pass parser, full file in memory.** Fine for typical IFX
  files (a few MB) but will not stream multi-GB exports.
- **Pivot requires exactly 2 axes.** 3D EEM time-courses fall back to
  Data + Metadata only; no faceted / multi-sheet pivot is generated.
- **Tk theming.** The "aqua" theme is requested but not required;
  on Linux/Windows the default Tk theme is used and the layout has
  not been visually tuned for those platforms.
- **Worker thread touches Tk state for cheap updates.** `StringVar`
  writes and `progress.config` calls are made directly from the worker
  and rely on Tk's macOS forgiveness. Modal dialogs are correctly
  marshalled via `_on_ui`; if portability becomes a goal, the
  status/progress updates should follow the same pattern.

# IFX to Excel Converter — macOS

A GUI tool to convert ISS spectrofluorometer `.ifx` experiment files into
Excel `.xlsx` workbooks. Reads the whole folder in one go and lets you
pick the output format per conversion.

---

## What files are supported?

Any `.ifx` file with the ISS `ISS_Experiment_Ver_1_0` signature. The tool
reads the file's own `Measurable=`, `Space=`, and `Columns=` metadata, so
it's not hardcoded to any specific measurement type.

**Pivot (MB mode / Spectra sheet) works for:**
- Emission spectra (Iteration × EmissionWavelength)
- Excitation spectra (Iteration × ExcitationWavelength)
- Time-resolved scans (Iteration × Time)
- Anisotropy, Polarization, Lifetime, etc. (any `Measurable`)
- 2D EEM scans (Excitation × Emission, no iteration)
- Any file with exactly 2 axes in `Space=`

Axis labels and units are pulled from the file header automatically
(e.g. "Excitation Wavelength (nm)", "Time (s)").

**Files where the pivot can't be built** (e.g. 3+ axes like a 3D EEM
time-course) still work in **Comprehensive** mode — you get the Data
sheet + Metadata sheet, just without the pivoted Spectra sheet. MB mode
will show a clear error and suggest switching to Comprehensive.


---

## Output formats

The GUI has a radio-button toggle for two formats:

### 1. MB compatible (default)

Single sheet, **pivoted matrix**: wavelengths down column A, iterations
across the top row, intensity values in the body. Matches the output of
MB's `ImportFluorimeterSpectra` VBA macro. Good for plotting each column
as a spectrum.

Drops `IntensityStdError` and `ExcitationChannel` (same as MB's macro).

| Wavelength (nm) | 1      | 2      | 3      |
|-----------------|--------|--------|--------|
| 400             | 583.9  | 2122.9 | 2115.1 |
| 401             | 554.6  | 2234.2 | 2327.7 |
| …               | …      | …      | …      |

### 2. Comprehensive

Three sheets in one workbook:
- **Spectra** — same pivoted matrix as above
- **Data** — full long-format table with all original columns (Iteration,
  EmissionWavelength, Intensity, IntensityStdError, ExcitationChannel)
- **Metadata** — file header fields (Title, Timestamp, AcquisitionType,
  wavelength bandwidths, iteration count, etc.)

Intensity values are preserved at **full precision** in both formats
(e.g. `386.379935`, `-59.4643976`).

---

## Option 1 — Run the script directly (easiest, ~2 minutes)

### 1. Install Python (if you don't have it)

Open **Terminal** (⌘+Space → "Terminal") and run:

```bash
python3 --version
```

If it says "command not found" or prompts for developer tools, install
Python from **https://www.python.org/downloads/macos/** (the official
installer includes Tkinter, which the GUI needs).

### 2. Install the one required library

```bash
pip3 install openpyxl
```

### 3. Run it

Put `ifx_to_xlsx_gui.py` somewhere (e.g. Desktop), then:

```bash
python3 ~/Desktop/ifx_to_xlsx_gui.py
```

### Optional: double-clickable launcher

Make a file called `Convert_IFX.command` next to the script containing:

```bash
#!/bin/bash
cd "$(dirname "$0")"
python3 ifx_to_xlsx_gui.py
```

Then one-time in Terminal:

```bash
chmod +x ~/Desktop/Convert_IFX.command
```

Now you can double-click it from Finder.

---

## Option 2 — Build a real macOS `.app` bundle

If you want a proper Mac app to drop in `/Applications`:

```bash
pip3 install pyinstaller
pyinstaller --windowed --name "IFX to Excel" ifx_to_xlsx_gui.py
```

The result is `dist/IFX to Excel.app`. Drag it to `/Applications`.

> **First-launch gotcha:** unsigned apps are blocked by Gatekeeper.
> Right-click the app → **Open** → confirm **Open** in the dialog.
> You only need to do this once.

---

## Using the GUI

1. Click **Choose .ifx files…** — pick one or many
2. Select the output format (MB compatible / Comprehensive)
3. Pick a **Save to** folder (defaults to Desktop)
4. (Single file only) edit the **Output name** if you want something
   other than the source filename — `.xlsx` is added automatically
5. Click **Convert**

Each `foo.ifx` becomes `foo.xlsx` (or your chosen name) in the chosen
folder. If a file with that name already exists you'll get a confirm
dialog — choose **No** to skip it and keep going with the rest of the
batch.

---

## Troubleshooting

- **"No module named openpyxl"** — run `python3 -m pip install openpyxl`
- **"No module named tkinter"** — your Python was built without Tk;
  reinstall from https://www.python.org/downloads/macos/
- **GUI too cramped** — resize the window, controls reflow

---

## How it differs from MB's VBA macro

Both tools produce the same pivoted matrix when you pick **MB compatible**
mode. Differences:

- MB's macro runs inside Excel and adds a sheet to the active workbook.
  This tool creates a separate `.xlsx` file per `.ifx`.
- MB's macro rounds intensity for display (`NumberFormat = "0"`); this
  tool keeps full precision (`0.######`), so values with many decimals
  display correctly.
- This tool can also produce the Comprehensive format, which keeps the
  StdError, ExcitationChannel, and metadata that MB's macro discards.

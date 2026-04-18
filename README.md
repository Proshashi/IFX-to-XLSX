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

## Setup

### 1. Install Python (if you don't have it)

Open **Terminal** (⌘+Space → "Terminal") and run:

```bash
python3 --version
```

If it says "command not found" or prompts for developer tools, install
Python from **https://www.python.org/downloads/macos/** (the official
installer includes Tkinter, which the GUI needs).

### 2. Launch with `start.sh`

From the project folder:

```bash
./start.sh
```

`start.sh` checks Python is available, installs `openpyxl` on first
run if needed, and launches the GUI. No build step, no app bundle.

If macOS refuses to execute the script, make it executable once:

```bash
chmod +x start.sh
```

### Optional: double-clickable launcher

`Convert_IFX.command` is included for users who prefer to launch from
Finder by double-click. It does the same thing as `start.sh` minus the
dependency check.

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

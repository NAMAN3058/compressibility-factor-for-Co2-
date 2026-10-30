# CO2 absorption dashboard

Upload an Excel workbook of pressure-vs-time readings. The app gets the CO2 compressibility factor Z from the
**Peace Software online CO2 calculator** (not CoolProp), calculates the absorbed CO2, and gives you back the same
workbook with the results filled in, plus a `Calculation Notes` sheet.

## Install (Windows, once)

1. Install Python 3.9 or newer from https://www.python.org/downloads/ - tick **"Add python.exe to PATH"** in the installer.
2. Open Command Prompt in this folder (type `cmd` in the folder's address bar and press Enter).
3. Run:

```
pip install -r requirements.txt
```

(Only if the app later says it needs a browser engine: `playwright install chromium`.)

## Run

```
python app.py
```

The dashboard opens at http://127.0.0.1:5057. Close the Command Prompt window (or press Ctrl+C) to stop it.

## Use

1. **Choose file** - your workbook (one or many sheets; each sheet needs a `Time` column and a pressure column such as `100 psi`).
2. Set the **temperature** (default 30 °C, any value; °C or K). It is used for the website lookup and for T(K) in the moles equation.
3. Optional: open **Calculation settings** (see below).
4. Click **Test website connection** once. It should show `"ok": true` and a Z value. Compare that Z with the website by hand.
5. Click **Start**. Values fill in live. When finished, click **Download Excel**.

The pressure unit is read from the column header (`psi`, `bar`, `kPa`, `MPa`). If the header has no unit, the
"Pressure unit if not in header" box is used.

## What goes into the output workbook

All original sheets, cells and formatting are kept. New columns are added to the right of the original ones
(or filled, if a column with that header already exists):

`Seconds, Minutes, P1, P2, Z1, Z2, Pi (Pa), Pf (Pa), Absorbed mol, Absorbed CO2 (gm), Quantity (gm/L), Concentration, mol/kg, Error`

Everything is stored as numbers. The `Calculation Notes` sheet lists the formulas, units, every setting used, and any failed rows.

| Column | Formula |
|---|---|
| Minutes | Seconds / 60 (Seconds is read from the Time text: `3 min` = 180, `27 min 30 sec` = 1650) |
| Z1 | Z at the sheet's first pressure (looked up once per sheet) |
| Z2 | Z at each row's pressure |
| Pi, Pf (Pa) | P1, P2 x 6894.757293168 (psi); bar x 1e5; kPa x 1e3; MPa x 1e6 |
| Absorbed mol | ((Pi / Z1) - (Pf / Z2)) x V / (R x T), T(K) = T(°C) + 273.15 |
| Absorbed CO2 (gm) | Absorbed mol x molar mass |
| Quantity (gm/L) | Absorbed CO2 (g) / solution volume (L) |
| Concentration | Quantity (gm/L) x concentration factor |
| mol/kg | Absorbed mol / solution mass (kg) |
| Error | 5 % of mol/kg (or a fixed 0.001 mol/kg - choose in settings) |

### Settings (defaults follow your workbook)

| Setting | Default | Note |
|---|---|---|
| Gas volume V | 0.001 m³ | The "0.001". The equation is only dimensionally correct if this is a volume in **m³** (1 L). Confirm it is your cell's gas volume. |
| R | 8.314 J/(mol·K) | |
| CO2 molar mass | 44.01 g/mol | Your original workbook used 44 (0.02 % difference). Set 44 to reproduce it. |
| Solution volume | 300 mL | |
| Concentration factor | 1001.142303 | Copied from your workbook. Units and source are **not verified**. (1000 / 1001.142303 = 0.99886, which looks like a density of about 0.99886 g/mL, but this is only an observation.) |
| Solution density | 1.000 g/mL | Used for mol/kg. 1.000 reproduces your original `mol x 1000 / 300`, which is really mol per litre. Brines are denser; enter the real density for a true mol/kg. |
| Solution mass | (empty) | If filled, used instead of volume x density. |
| Error | 5 % of mol/kg | Or fixed 0.001 mol/kg. |

## If something fails

- A website request that fails is retried automatically. If it still fails, the row is shown in red with its error message, listed in
  "Rows without a result", and **left blank** - never estimated.
- The workbook is then named `..._Z_INCOMPLETE.xlsx` and the Calculation Notes status says INCOMPLETE. The dashboard never says
  "Finished" unless every row has a result.
- **Retry failed rows only** repeats just the failed lookups and keeps everything that already worked.
- The website only calculates pressures **>= 1 bar (14.504 psi)**. Lower readings always fail and are reported as failed rows.
- **Stop** gives a partial, INCOMPLETE workbook. The uploaded file is never modified.

## Speed

Each distinct pressure is requested once per run (cached), Z1 once per sheet, and several requests run in parallel
("Parallel requests", default 4). Lower it to 1-2 if the website starts refusing requests.

## Other

- Files are saved in the `outputs` folder next to `app.py`.
- Use a different address with: `set PEACE_URL=https://...` then `python app.py`.
- Change the port with `set PORT=5060`.

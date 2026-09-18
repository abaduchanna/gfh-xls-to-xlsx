# Legacy Excel Converter (GFH)

One-time batch tool: pick a folder and convert every legacy
`.xls / .xlsm / .xlt / .xlsb` file to modern `.xlsx` using **real Excel
(COM)** so formatting, formulas and data are preserved. Originals are kept by
default (optional delete after success); already-converted files are skipped
unless "overwrite" is ticked. Optional subfolder recursion.

## Build
Windows EXE built via GitHub Actions on push.

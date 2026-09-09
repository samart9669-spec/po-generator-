# PO Generator

One-click monthly PO generator for the Seamless production plan.

## What it does

Transforms the monthly Excel production plan into supplier PO output using the existing PO template structure.

Current model/month mappings:

- AiR: M9, M10, M11
- Freedom: M9, M10, M11
- Kids: M11

The generator reads SKU-level plan data, groups pack-based SKUs, maps description/pack/unit price from the PO template, and writes PO item lines and totals.

## One-click on Mac

1. Put the monthly plan `.xlsx` and PO template `.xlsx` into `INPUT/`.
2. Double-click `Generate_PO.command`.
3. The result is written to `OUTPUT/PO_Generated.xlsx`.

## One-click on Windows

1. Put the monthly plan `.xlsx` and PO template `.xlsx` into `INPUT/`.
2. Double-click `Generate_PO.bat`.
3. The result is written to `OUTPUT/PO_Generated.xlsx`.

The launcher automatically selects the newest matching plan and PO template.

## Python runtime

The generator uses Python standard library only. Python 3.10+ is recommended.

## Important

The repository intentionally does not store production spreadsheets or supplier data. Place those files locally in `INPUT/` before running.

Commercial terms in the source PO template are preserved rather than silently overridden. Review deposit/payment terms before sending a generated PO.

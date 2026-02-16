from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable


def iter_excel_files(root: Path) -> Iterable[Path]:
    # Expect structure: root/Driver*/<files>.xlsx
    for driver_dir in sorted(root.glob("Driver*")):
        if not driver_dir.is_dir():
            continue
        for xlsx in sorted(driver_dir.glob("*.xlsx")):
            yield xlsx


def parse_maneuver_from_filename(xlsx_path: Path) -> str:
    # Requirement: maneuver is the second part when splitting the filename by '_'
    # Example: STISIMData_U-Turnings.xlsx -> 'U-Turnings'
    stem = xlsx_path.stem
    parts = stem.split("_")
    return parts[1] if len(parts) >= 2 else ""


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Merge STISIM maneuver Excel files across Driver folders into a single CSV, "
            "adding driver and maneuver columns."
        )
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("ManiobrasSimulador"),
        help="Root folder containing Driver1/, Driver2/, ... (default: ManiobrasSimulador)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("maniobras_combined.csv"),
        help="Output CSV path (default: maniobras_combined.csv)",
    )
    parser.add_argument(
        "--sheet",
        type=str,
        default=None,
        help=(
            "Excel sheet name to read. If omitted, reads the first sheet of each workbook. "
            "(Useful if files contain multiple sheets.)"
        ),
    )
    parser.add_argument(
        "--skiprows",
        type=int,
        default=0,
        help="Number of initial rows to skip when reading each sheet (default: 0)",
    )

    args = parser.parse_args()

    input_root: Path = args.input
    output_csv: Path = args.output

    if not input_root.exists() or not input_root.is_dir():
        raise SystemExit(f"Input folder not found or not a directory: {input_root}")

    try:
        import pandas as pd  # type: ignore
    except Exception as exc:  # pragma: no cover
        raise SystemExit(
            "Missing dependency: pandas. Install with: pip install pandas openpyxl\n"
            f"Original error: {exc}"
        )

    columns_to_remove = [
        "Accidents",
        "Collisions",
        "Peds Hit",
        "Speeding Tics",
        "Red Lgt Tics",
        "Speed Exceed",
        "Stop Sign Ticks",
    ]

    frames: list["pd.DataFrame"] = []

    for xlsx_path in iter_excel_files(input_root):
        driver = xlsx_path.parent.name
        maneuver = parse_maneuver_from_filename(xlsx_path)

        # Read one sheet per workbook.
        # If you don't pass --sheet, we read the first sheet (index 0).
        sheet_to_read = args.sheet if args.sheet is not None else 0
        df = pd.read_excel(
            xlsx_path,
            sheet_name=sheet_to_read,
            engine="openpyxl",
            skiprows=args.skiprows,
        )

        # Defensive: if sheet_name was None (or a list) pandas can return a dict.
        if isinstance(df, dict):
            if not df:
                continue
            df = next(iter(df.values()))

        # Some sheets may be empty
        if df is None:
            continue

        # Ensure driver/maneuver are present as columns
        df.insert(0, "driver", driver)
        df.insert(1, "maneuver", maneuver)

        # Remove unwanted columns (ignore if a file/sheet doesn't have them)
        df = df.drop(columns=columns_to_remove, errors="ignore")

        frames.append(df)

    if not frames:
        raise SystemExit(
            f"No .xlsx files found under {input_root}. Expected: {input_root}/Driver*/.xlsx"
        )

    combined = pd.concat(frames, ignore_index=True, sort=False)

    # Write CSV with BOM so it opens cleanly in Excel on Windows
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(output_csv, index=False, encoding="utf-8-sig")

    print(f"Wrote {len(combined):,} rows to: {output_csv.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

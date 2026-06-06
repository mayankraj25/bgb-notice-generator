#!/usr/bin/env python3
"""
BGB Notice Generator — CLI entry point.

Usage examples:
  python main.py --excel input.xlsx --type extension
  python main.py --excel input.xlsx --type encashment
  python main.py --excel input.xlsx --type both
  python main.py --excel input.xlsx --type both --row 2
  python main.py --excel input.xlsx --type both --letter-number 800542/230/E8
"""

import argparse
import os
import sys
from datetime import date

from excel_parser import parse_excel
from letter_generator import generate_extension_notice, generate_encashment_notice
from utils import validity_in_window, format_date


def main():
    parser = argparse.ArgumentParser(
        description="Auto-generate Bank Guarantee Bond (BGB) notices as .docx files."
    )
    parser.add_argument("--excel",         required=True, help="Path to the Excel input file.")
    parser.add_argument("--type",          required=True,
                        choices=["extension", "encashment", "both"],
                        help="Type of notice to generate.")
    parser.add_argument("--row",           type=int, default=None,
                        help="Generate notice for a specific Excel row number only (1-based).")
    parser.add_argument("--output-dir",    default="output",
                        help="Directory to save generated .docx files (default: output/).")
    parser.add_argument("--letter-number", default="800542/230/E8",
                        help="Letter number for the notices (default: 800542/230/E8).")
    parser.add_argument("--letter-date",   default=None,
                        help="Letter date as YYYY-MM-DD (default: today).")
    args = parser.parse_args()

    # Resolve letter date
    if args.letter_date:
        try:
            letter_date = date.fromisoformat(args.letter_date)
        except ValueError:
            print(f"ERROR: Invalid --letter-date '{args.letter_date}'. Use YYYY-MM-DD.")
            sys.exit(1)
    else:
        letter_date = date.today()

    today = date.today()

    # Ensure output directory exists
    os.makedirs(args.output_dir, exist_ok=True)

    # Parse Excel
    print(f"\nReading Excel file: {args.excel}")
    try:
        rows = parse_excel(args.excel)
    except FileNotFoundError:
        print(f"ERROR: Excel file not found: {args.excel}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Failed to parse Excel file: {e}")
        sys.exit(1)

    print(f"  Total data rows found: {len(rows)}")

    # Filter to specific row if requested
    if args.row is not None:
        rows = [r for r in rows if r['row_index'] == args.row]
        if not rows:
            print(f"ERROR: Row {args.row} not found or is a header/blank row.")
            sys.exit(1)

    # Counters
    generated   = 0
    skipped_exp = 0
    skipped_not = 0
    skipped_err = 0

    print(f"\nProcessing rows (today = {today})...\n")

    for row in rows:
        rid = row['row_index']

        # Validate required fields
        missing = []
        if not row['ca_number']:
            missing.append("CA number")
        if not row['bgb_number']:
            missing.append("BGB number")
        if row['validity_date'] is None:
            missing.append("validity date")

        if missing:
            print(f"  [WARNING] Row {rid}: Missing fields ({', '.join(missing)}) — skipped.")
            skipped_err += 1
            continue

        validity = row['validity_date']

        # Check expiry
        if validity < today:
            print(f"  [SKIPPED] Row {rid} — BGB already expired on {format_date(validity)}.")
            skipped_exp += 1
            continue

        # Check if within 60-day window
        if not validity_in_window(validity, today):
            # Not yet due — skip silently
            skipped_not += 1
            continue

        # Generate requested notice(s)
        print(f"  [PROCESSING] Row {rid} | CA: {row['ca_number']} | "
              f"Contractor: {row['contractor_name']} | "
              f"Valid upto: {format_date(validity)}")

        counter = generated + 1   # simple running letter sub-number

        try:
            if args.type in ("extension", "both"):
                fpath = generate_extension_notice(
                    row          = row,
                    letter_number= args.letter_number,
                    letter_date  = letter_date,
                    output_dir   = args.output_dir,
                )
                print(f"    -> Extension Notice: {fpath}")
                generated += 1

            if args.type in ("encashment", "both"):
                fpath = generate_encashment_notice(
                    row               = row,
                    letter_number     = args.letter_number,
                    letter_date       = letter_date,
                    prev_letter_number= args.letter_number,   # same ref for linked pair
                    prev_letter_date  = letter_date,
                    output_dir        = args.output_dir,
                )
                print(f"    -> Encashment Notice: {fpath}")
                generated += 1

        except Exception as e:
            print(f"    [ERROR] Row {rid}: Failed to generate notice — {e}")
            skipped_err += 1

    # Summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"  Notices generated   : {generated}")
    print(f"  Skipped (expired)   : {skipped_exp}")
    print(f"  Skipped (not due)   : {skipped_not}")
    print(f"  Skipped (errors)    : {skipped_err}")
    print(f"  Output directory    : {os.path.abspath(args.output_dir)}")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()

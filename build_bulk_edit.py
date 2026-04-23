"""
Populate the Bulk Edit master template (sheet "Appraisal User Qr code Mapping")
from the StockTake inventory file + the QR codes file.

Logic replicates the manual workflow:
  1. Filter StockTake rows where Room = "Inventory".
  2. For each row, XLOOKUP the Location in the QR codes' Description column
     and return the matching UserQrCode.
     Locations annotated like "R58 (AUH-D1-INVOICE-21-04)" are matched
     on the prefix before the first "(".
  3. Write two columns to the template: Appraisal code (Deal Id), User Qr Code.
  4. Sort by Deal Id, then Location.

Usage:
  python build_bulk_edit.py
      -> uses the three default filenames in this folder and writes
         "Bulk Edit (87) - filled.xlsx" next to the template.

  python build_bulk_edit.py <inventory.xlsx> <qrcodes.xlsx> <template.xlsx> [output.xlsx]
"""
import re
import sys
from pathlib import Path
from openpyxl import load_workbook

HERE = Path(__file__).parent
DEFAULT_INVENTORY = HERE / "INV 21.04.26.xlsx"
DEFAULT_QRCODES = HERE / "QR codes 1.xlsx"
DEFAULT_TEMPLATE = HERE / "Bulk Edit (87).xlsx"

INVENTORY_SHEET = "StockTake Template"
QR_SHEET = "User Qr Code"
OUTPUT_SHEET = "Appraisal User Qr code Mapping"

ROOM_COL = 0       # column A
LOCATION_COL = 2   # column C
DEAL_ID_COL = 4    # column E

_LOCATION_STRIP_RE = re.compile(r"^([A-Za-z0-9]+)\s*\(")


def build_qr_lookup(qrcodes_source) -> dict:
    """qrcodes_source can be a path or any file-like object openpyxl accepts."""
    wb = load_workbook(qrcodes_source, data_only=True, read_only=True)
    ws = wb[QR_SHEET]
    lookup = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        qr, desc = row[0], row[1]
        if qr is None or desc is None:
            continue
        key = str(desc).strip()
        lookup.setdefault(key, qr)
    wb.close()
    return lookup


def match_qr(location, lookup: dict):
    if location is None:
        return None
    loc = str(location).strip()
    if loc in lookup:
        return lookup[loc]
    m = _LOCATION_STRIP_RE.match(loc)
    if m and m.group(1) in lookup:
        return lookup[m.group(1)]
    return None


def collect_rows(inventory_source, qr_lookup: dict):
    """inventory_source can be a path or any file-like object.
    Returns (rows, unmatched, skipped_no_deal_id).
    """
    wb = load_workbook(inventory_source, data_only=True, read_only=True)
    ws = wb[INVENTORY_SHEET]
    rows = []
    unmatched = []
    skipped_no_deal_id = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[ROOM_COL] != "Inventory":
            continue
        deal_id = row[DEAL_ID_COL]
        location = row[LOCATION_COL]
        if deal_id is None or str(deal_id).strip().lower() in ("", "no deal id"):
            skipped_no_deal_id += 1
            continue
        qr = match_qr(location, qr_lookup)
        if qr is None:
            unmatched.append((deal_id, location))
        rows.append((deal_id, qr, location))
    wb.close()
    rows.sort(key=lambda r: (str(r[0]), str(r[2]) if r[2] is not None else ""))
    return rows, unmatched, skipped_no_deal_id


def write_output(template_source, output_target, rows):
    """template_source can be a path or file-like; output_target can be a path or file-like."""
    wb = load_workbook(template_source)
    ws = wb[OUTPUT_SHEET]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)
    for deal_id, qr, _loc in rows:
        ws.append([deal_id, qr])
    wb.save(output_target)


def main():
    args = sys.argv[1:]
    if len(args) >= 3:
        inventory = Path(args[0])
        qrcodes = Path(args[1])
        template = Path(args[2])
        output = Path(args[3]) if len(args) >= 4 else template.with_name(f"{template.stem} - filled.xlsx")
    else:
        inventory = DEFAULT_INVENTORY
        qrcodes = DEFAULT_QRCODES
        template = DEFAULT_TEMPLATE
        output = template.with_name(f"{template.stem} - filled.xlsx")

    for p in (inventory, qrcodes, template):
        if not p.exists():
            sys.exit(f"Missing input: {p}")

    print(f"Reading QR codes:  {qrcodes.name}")
    qr_lookup = build_qr_lookup(qrcodes)
    print(f"  {len(qr_lookup)} unique QR descriptions loaded")

    print(f"Reading inventory: {inventory.name}")
    rows, unmatched, skipped = collect_rows(inventory, qr_lookup)
    matched = len(rows) - len(unmatched)
    print(f"  {len(rows)} Inventory rows -> matched {matched}, unmatched {len(unmatched)}")
    print(f"  skipped (No Deal Id): {skipped}")

    print(f"Writing output:    {output.name}")
    write_output(template, output, rows)
    print("Done.")

    if unmatched:
        print("\nWARNING: no QR code found for these rows:")
        for deal_id, loc in unmatched[:20]:
            print(f"  {deal_id}  |  location={loc!r}")
        if len(unmatched) > 20:
            print(f"  ... and {len(unmatched) - 20} more")


if __name__ == "__main__":
    main()

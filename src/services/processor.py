"""
src/services/processor.py - Core data extraction, mapping, validation and Excel generation.
"""

import shutil
import calendar
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

from src.config import (
    PREFIX,
    TARGET_SECTION,
    COL_MAP,
    OUTPUT_COLS,
    NUMERIC_COLS,
    VALIDATION_MAP,
    YELLOW_FILL,
    fin_year,
    resolve_report_period,
)


def _is_target_section(cell_val) -> bool:
    """Check if cell text refers to STATE WISE SALES."""
    if cell_val is None:
        return False
    clean = "".join(c for c in str(cell_val).upper() if c.isalnum())
    return "STATEWISESALES" in clean or "STATEWISE" in clean


def parse_repotic_sheet(repotic_path: Path, sheet_name: str):
    """
    Returns (df_raw, error_string).
    df_raw has uppercase column headers from the STATE WISE SALES section.
    Returns (None, error_msg) on failure.
    """
    wb = load_workbook(repotic_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return None, f"Sheet '{sheet_name}' not found in {repotic_path.name}"

    ws = wb[sheet_name]
    all_rows = list(ws.iter_rows(values_only=True))

    # 1. Find "STATE WISE SALES" title row (checks across all columns in row)
    section_idx = None
    for i, row in enumerate(all_rows):
        if any(_is_target_section(cell) for cell in row):
            section_idx = i
            break

    # If section title found, header is the next non-empty row
    header_idx = None
    if section_idx is not None:
        for i in range(section_idx + 1, len(all_rows)):
            if any(v is not None for v in all_rows[i]):
                header_idx = i
                break
    else:
        # Fallback: Check if sheet directly contains table headers (STATE + TAXABLE/AMOUNT)
        for i, row in enumerate(all_rows[:15]):
            row_vals = [str(c).strip().upper() for c in row if c is not None]
            if any("STATE" in v for v in row_vals) and any("TAXABLE" in v or "AMOUNT" in v or "RATE" in v for v in row_vals):
                header_idx = i
                break

    if header_idx is None:
        return None, f"No header row or '{TARGET_SECTION}' found in '{sheet_name}'"

    raw_headers = [
        str(v).strip().upper() if v is not None else f"_COL{j}"
        for j, v in enumerate(all_rows[header_idx])
    ]

    # Collect data rows
    data_rows = []
    for i in range(header_idx + 1, len(all_rows)):
        row = all_rows[i]
        if all(v is None for v in row):
            continue
        if row[0] is None or str(row[0]).strip() == "":
            continue
        first_cell = str(row[0]).strip().upper()
        if first_cell in ("TOTAL", "GRAND TOTAL", "SUMMARY"):
            break
        data_rows.append(row)

    if not data_rows:
        return None, f"No data rows under '{TARGET_SECTION}' in '{sheet_name}'"

    return pd.DataFrame(data_rows, columns=raw_headers), None


def _normalize_key(s) -> str:
    """Normalize a state string for lookup: strip, collapse spaces, uppercase."""
    return " ".join(str(s).strip().upper().split())


def load_ledger_all_sheets(ledger_path: Path, logs: list) -> tuple:
    """
    Loads every sheet of the ledger file into a dict:
        { normalized_sheet_name: { normalized_state_key: (Pty_Name, Sales Ledger) } }
    """
    wb = load_workbook(ledger_path, data_only=True)
    available = wb.sheetnames
    logs.append(f"Ledger file: {ledger_path.name}  |  sheets: {available}")

    ledger: dict = {}
    for sheet_name in available:
        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            continue

        # Search across the first 15 rows for the header row
        col_state = col_sales = col_pty = None
        header_row_idx = None

        for r_idx, row in enumerate(all_rows[:15]):
            c_state = c_sales = c_pty = None
            for idx, h in enumerate(row):
                if h is None:
                    continue
                clean_h = "".join(ch for ch in str(h).lower() if ch.isalnum())
                if clean_h in ("stateofsupply", "state", "placeofsupply", "pos", "states"):
                    c_state = idx
                elif clean_h in ("salesledger", "salesledgername", "ledger", "ledgername"):
                    c_sales = idx
                elif clean_h in ("ptyname", "partyname", "party", "customer", "customername", "pty"):
                    c_pty = idx

            if c_state is not None and c_sales is not None and c_pty is not None:
                col_state, col_sales, col_pty = c_state, c_sales, c_pty
                header_row_idx = r_idx
                break

        if col_state is None or col_sales is None or col_pty is None:
            logs.append(f"  [WARNING] Sheet '{sheet_name}' missing expected columns (State:{col_state is not None}, SalesLedger:{col_sales is not None}, Party:{col_pty is not None}) – skipped.")
            continue

        sheet_dict: dict = {}
        for row in all_rows[header_row_idx + 1:]:
            if col_state >= len(row) or col_sales >= len(row) or col_pty >= len(row):
                continue
            state_raw = row[col_state]
            if state_raw is None or str(state_raw).strip() == "":
                continue
            key       = _normalize_key(state_raw)
            pty_val   = row[col_pty]
            sales_val = row[col_sales]
            sheet_dict[key] = (pty_val, sales_val)

        norm_sheet = sheet_name.strip().lower()
        ledger[norm_sheet] = sheet_dict
        logs.append(f"  Ledger sheet {sheet_name!r}: {len(sheet_dict)} states")

    return ledger, available


def find_matching_ledger_sheet(repotic_sheet_name: str, ledger: dict) -> str:
    """Find the best matching ledger sheet name for a given repotic sheet name."""
    norm_r = repotic_sheet_name.strip().lower()
    if norm_r in ledger:
        return norm_r

    # Known marketplaces match
    known_marketplaces = ["flipkart", "meesho", "snapdeal", "amazon", "myntra", "glowroad", "ajio", "jiomart", "nykaa"]
    for mkt in known_marketplaces:
        if mkt in norm_r:
            for l_name in ledger.keys():
                if mkt in l_name:
                    return l_name

    # Alphanumeric substring match
    clean_r = "".join(c for c in norm_r if c.isalnum())
    for l_name in ledger.keys():
        clean_l = "".join(c for c in l_name if c.isalnum())
        if clean_l and clean_r and (clean_l in clean_r or clean_r in clean_l):
            return l_name

    if len(ledger) == 1:
        return list(ledger.keys())[0]

    return None


def get_marketplace_prefix(sheet_name: str) -> str:
    """Get the 2-letter invoice series code for a marketplace."""
    clean_name = sheet_name.strip().upper()
    if clean_name in PREFIX:
        return PREFIX[clean_name]
    for k, v in PREFIX.items():
        if k in clean_name:
            return v
    letters = "".join(c for c in clean_name if c.isalpha())
    return letters[:2].upper() if len(letters) >= 2 else "MK"


def lookup_state(state_raw, sheet_dict: dict) -> tuple:
    if not state_raw or str(state_raw).strip() == "":
        return None, None
    key = _normalize_key(state_raw)
    result = sheet_dict.get(key)
    if result:
        return result
    return None, None


def _to_python(val):
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(val, "item"):
        return val.item()
    return val


def run_automation(repotic_path: Path, ledger_path: Path,
                   template_path: Path, output_path: Path,
                   month_override: str = None) -> dict:
    logs: list  = []
    unmatched   = []
    all_frames  = []
    summary     = []

    # 1. Report Period
    year, month, source = resolve_report_period(repotic_path, month_override)
    fy_str = fin_year(month, year)
    last_day = calendar.monthrange(year, month)[1]
    inv_date = f"{last_day:02d}-{month:02d}-{year}"

    logs.append(f"Report month: {month:02d}/{year} (Source: {source}) | Financial Year: {fy_str}")
    logs.append(f"Invoice Date (Month-End): {inv_date}")
    if source == "assumed":
        logs.append(f"  [WARNING] Month was not provided or detected; assumed previous calendar month: {year}-{month:02d}")

    # Load ledger
    ledger, available_ledger_sheets = load_ledger_all_sheets(ledger_path, logs)

    # 2. Discover Sheets
    repotic_wb = load_workbook(repotic_path, data_only=True)
    sheets_to_process = []
    sheet_to_ledger_map = {}
    skipped_reasons = []

    for s_name in repotic_wb.sheetnames:
        clean_name = s_name.strip()
        norm_s = clean_name.lower()

        if "other b2b" in norm_s or "cdnr" in norm_s:
            skipped_reasons.append(f"'{s_name}': Skipped (B2B/CDNR section)")
            continue

        matched_ledger = find_matching_ledger_sheet(clean_name, ledger)
        if not matched_ledger:
            skipped_reasons.append(
                f"'{s_name}': No matching sheet found in ledger. (Available ledger sheets: {list(ledger.keys())})"
            )
            continue

        ws = repotic_wb[s_name]
        has_section = any(
            any(_is_target_section(cell) for cell in row)
            for row in ws.iter_rows(values_only=True) if row
        )
        if not has_section:
            sample_rows = list(ws.iter_rows(max_row=15, values_only=True))
            has_table = any(
                any("STATE" in str(c).upper() for c in r if c) and
                any("TAXABLE" in str(c).upper() or "AMOUNT" in str(c).upper() or "RATE" in str(c).upper() for c in r if c)
                for r in sample_rows
            )
            if not has_table:
                skipped_reasons.append(f"'{s_name}': Neither '{TARGET_SECTION}' nor state sales columns found")
                continue

        sheets_to_process.append(s_name)
        sheet_to_ledger_map[s_name] = matched_ledger

    if not sheets_to_process:
        reason_list = "\n".join(f"  - {r}" for r in skipped_reasons)
        raise ValueError(
            f"No eligible marketplace sheets found in REPOTIC workbook with matching ledger sheets.\n"
            f"REPOTIC sheets found: {repotic_wb.sheetnames}\n"
            f"Ledger sheets in file: {available_ledger_sheets}\n"
            f"Ledger sheets recognized: {list(ledger.keys())}\n"
            f"Diagnostic details:\n{reason_list}"
        )

    logs.append(f"Marketplace sheets to process: {sheets_to_process}")

    # 3. Parse + Process
    for sheet_name in sheets_to_process:
        logs.append(f"─── Processing: {sheet_name} ───")
        matched_ledger = sheet_to_ledger_map[sheet_name]
        sheet_dict = ledger[matched_ledger]
        mkt_code = get_marketplace_prefix(sheet_name)

        df_raw, err = parse_repotic_sheet(repotic_path, sheet_name)
        if err:
            logs.append(f"  [SKIP] {err}")
            continue

        n = len(df_raw)
        logs.append(f"  Rows extracted: {n}")

        out_df = pd.DataFrame(index=range(n), columns=OUTPUT_COLS, dtype=object)
        for repotic_col, output_col in COL_MAP.items():
            uc = repotic_col.upper()
            if uc in df_raw.columns:
                out_df[output_col] = df_raw[uc].values
            else:
                logs.append(f"  [WARNING] Column '{repotic_col}' not found – left blank.")

        out_df["Vch_Type"] = "Auto Sales"
        out_df["Bill of Supply"] = out_df["StateOfSupply"]

        inv_numbers = [
            f"{month:02d}/{fy_str}/{mkt_code}-{idx:02d}"
            for idx in range(1, n + 1)
        ]
        out_df["InvNo"] = inv_numbers
        out_df["Inv_Dt"] = inv_date

        inv_series = f"{inv_numbers[0]} to {inv_numbers[-1]}" if inv_numbers else "None"
        logs.append(f"  Invoice series: {inv_series}")

        for col in NUMERIC_COLS:
            if col in out_df.columns:
                out_df[col] = pd.to_numeric(out_df[col], errors="coerce").round(2)

        out_df["_marketplace"] = sheet_name

        records = out_df.to_dict(orient="records")
        pty_names, sales_ledgs, unmatched_flags = [], [], []

        for rec in records:
            state_val = rec.get("StateOfSupply")
            pty, ledg = lookup_state(state_val, sheet_dict)
            pty_names.append(pty)
            sales_ledgs.append(ledg)
            matched = pty is not None

            if not matched and state_val and str(state_val).strip():
                st = str(state_val).strip()
                logs.append(f"  [UNMATCHED] {sheet_name} – {st} not in ledger sheet '{matched_ledger}'")
                unmatched.append({"sheet": sheet_name, "state": st})
            unmatched_flags.append(not matched)

        out_df["Pty_Name"]        = pty_names
        out_df["Sales Ledger"]    = sales_ledgs
        out_df["_unmatched_flag"] = unmatched_flags

        sheet_ok = True
        val_detail = []
        for out_col, raw_col in VALIDATION_MAP.items():
            if raw_col in df_raw.columns:
                src = float(round(pd.to_numeric(df_raw[raw_col], errors="coerce").sum(), 2))
                out = float(round(pd.to_numeric(out_df[out_col],  errors="coerce").sum(), 2))
                ok  = bool(abs(src - out) <= 0.02)
                if not ok:
                    sheet_ok = False
                val_detail.append({"col": out_col, "source": src, "output": out, "ok": ok})
                tag = "OK" if ok else "MISMATCH"
                logs.append(f"    {out_col:<15} src={src:>13,.2f}  out={out:>13,.2f}  {tag}")

        summary.append({
            "sheet"         : sheet_name,
            "rows"          : int(n),
            "invoice_series": inv_series,
            "TaxableAmt"    : float(round(pd.to_numeric(out_df["TaxableAmt"], errors="coerce").sum(), 2)),
            "IGSTAmt"       : float(round(pd.to_numeric(out_df["IGSTAmt"],    errors="coerce").sum(), 2)),
            "CGSTAmt"       : float(round(pd.to_numeric(out_df["CGSTAmt"],    errors="coerce").sum(), 2)),
            "SGSTAmt"       : float(round(pd.to_numeric(out_df["SGSTAmt"],    errors="coerce").sum(), 2)),
            "Net_Amt"       : float(round(pd.to_numeric(out_df["Net_Amt"],    errors="coerce").sum(), 2)),
            "status"        : "OK" if sheet_ok else "MISMATCH",
            "validation"    : val_detail,
        })

        all_frames.append(out_df)

    if not all_frames:
        raise ValueError("No data extracted from any sheet. Check the input file.")

    combined_df = pd.concat(all_frames, ignore_index=True)
    logs.append(f"Total rows written: {int(len(combined_df))}")

    # 4. Write Excel
    shutil.copy2(template_path, output_path)
    wb_out = load_workbook(output_path)
    ws_out = wb_out.active
    ws_out.title = "Final Sample Format Tamilnadu"

    header_map = {
        str(ws_out.cell(row=1, column=c).value).strip(): c
        for c in range(1, ws_out.max_column + 1)
        if ws_out.cell(row=1, column=c).value is not None
    }

    pty_col_idx   = header_map.get("Pty_Name")
    sales_col_idx = header_map.get("Sales Ledger")

    data_records = combined_df[OUTPUT_COLS + ["_unmatched_flag"]].to_dict(orient="records")
    for row_offset, record in enumerate(data_records, start=2):
        is_unmatched = bool(record.get("_unmatched_flag", False))
        for col_name in OUTPUT_COLS:
            col_idx = header_map.get(col_name.strip())
            if col_idx is None:
                continue
            val = _to_python(record.get(col_name))
            cell = ws_out.cell(row=row_offset, column=col_idx, value=val)

            if col_name == "InvNo":
                cell.number_format = "@"
                cell.value = str(val) if val is not None else ""
            elif col_name in NUMERIC_COLS and val is not None:
                cell.number_format = "#,##0.00"

        if is_unmatched:
            for cidx in (pty_col_idx, sales_col_idx):
                if cidx:
                    ws_out.cell(row=row_offset, column=cidx).fill = YELLOW_FILL

    wb_out.save(output_path)
    logs.append(f"Saved: {output_path.name}  (sheet: '{ws_out.title}')")

    # 5. Preview
    preview_cols = [
        "InvNo", "Inv_Dt", "_marketplace", "Vch_Type", "StateOfSupply", "Bill of Supply",
        "Qty", "TaxPer", "TaxableAmt", "IGSTAmt", "SGSTAmt", "CGSTAmt",
        "Net_Amt", "Pty_Name", "Sales Ledger",
    ]
    preview_records = [
        {col: _to_python(row[col]) for col in preview_cols}
        for _, row in combined_df[preview_cols].head(10).iterrows()
    ]

    return {
        "logs"          : logs,
        "summary"       : summary,
        "unmatched"     : unmatched,
        "preview"       : preview_records,
        "total_rows"    : int(len(combined_df)),
        "report_month"  : f"{year}-{month:02d}",
        "financial_year": fy_str,
        "month_source"  : source,
    }

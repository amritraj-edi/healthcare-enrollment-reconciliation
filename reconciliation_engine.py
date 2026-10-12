from __future__ import annotations

import io
import re
import math
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


@dataclass
class ComparisonOptions:
    key_mapping: dict[str, str]
    column_mapping: dict[str, str]
    normalize_whitespace: bool = True
    normalize_case: bool = False
    normalize_dates: bool = True
    normalize_numbers: bool = True
    run_rule_checks: bool = True
    report_unmapped_columns: bool = True
    source_a_name: str = "Source A"
    source_b_name: str = "Source B"


def load_uploaded_file(uploaded, sheet_name=None) -> pd.DataFrame:
    """Read a CSV or Excel upload while keeping identifier-like values as strings."""
    uploaded.seek(0)
    name = getattr(uploaded, "name", "").lower()
    if name.endswith(".csv"):
        raw = uploaded.getvalue() if hasattr(uploaded, "getvalue") else uploaded.read()
        last_error = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False, encoding=encoding)
            except UnicodeDecodeError as exc:
                last_error = exc
        raise ValueError(f"Unable to decode CSV: {last_error}")
    if name.endswith(".xlsx") or name.endswith(".xls"):
        try:
            df = pd.read_excel(uploaded, sheet_name=sheet_name, dtype=str, keep_default_na=False)
        except ImportError as exc:
            raise ValueError("Reading legacy .xls files requires the xlrd dependency.") from exc
        if isinstance(df, dict):
            raise ValueError("Select a worksheet before reading an Excel workbook.")
        return df.fillna("")
    raise ValueError("Unsupported file type. Upload .xlsx, .xls, or .csv.")


def _clean_header(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def _norm_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", _clean_header(value).lower())


def suggest_matching_keys(a_df: pd.DataFrame, b_df: pd.DataFrame) -> list[list[str]]:
    """Suggest stable shared keys using header heuristics; suggestions require user confirmation."""
    a_cols = list(a_df.columns)
    b_norm = {_norm_header(c): c for c in b_df.columns}
    pairs = []
    for a_col in a_cols:
        n = _norm_header(a_col)
        if n in b_norm:
            pairs.append((a_col, b_norm[n]))
    # Canonical header groups with common real-world variants.
    aliases = {
        "employer": {"employerid", "groupid", "clientid", "employeridentifier", "groupnumber"},
        "employee": {"employeeid", "eeid", "employeeidentifier", "workerid", "associateid"},
        "ssn": {"ssn", "socialsecuritynumber", "taxpayerid"},
    }
    found = {}
    for a_col in a_cols:
        n = _norm_header(a_col)
        for canonical, names in aliases.items():
            if n in names:
                match = next((c for c in b_df.columns if _norm_header(c) in names), None)
                if match is not None:
                    found.setdefault(canonical, (a_col, match))
    suggestions = []
    if "employer" in found and "employee" in found:
        suggestions.append([found["employer"][0], found["employee"][0]])
    if "employee" in found:
        suggestions.append([found["employee"][0]])
    if "ssn" in found:
        suggestions.append([found["ssn"][0]])
    if not suggestions:
        exact = [a for a, b in pairs if _norm_header(a) in {"employeeid", "eeid", "workerid", "associateid"}]
        if exact:
            suggestions.append([exact[0]])
        elif pairs:
            # Do not recommend a random shared field as a matching key.
            pass
    # unique preserving order
    unique = []
    for item in suggestions:
        if item not in unique:
            unique.append(item)
    return unique


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


def _normalize(value: Any, field_name: str, options: ComparisonOptions) -> Any:
    if _is_blank(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        if options.normalize_dates:
            return pd.Timestamp(value).date().isoformat()
        return str(value)
    text = str(value)
    if options.normalize_whitespace:
        text = text.strip()
    if options.normalize_dates:
        # Parse only date-looking fields or unambiguous ISO date strings.
        fname = _norm_header(field_name)
        looks_date = any(x in fname for x in ("date", "dob", "birth", "effective", "termination", "hire"))
        if looks_date:
            parsed = pd.to_datetime(text, errors="coerce")
            if not pd.isna(parsed):
                text = parsed.date().isoformat()
    if options.normalize_numbers:
        fname = _norm_header(field_name)
        numeric_field = any(x in fname for x in ("amount", "election", "contribution", "deduction", "count"))
        if numeric_field:
            cleaned = text.replace(",", "").replace("$", "")
            try:
                num = float(cleaned)
                if math.isfinite(num):
                    return int(num) if num.is_integer() else num
            except ValueError:
                pass
    if options.normalize_case:
        text = text.casefold()
    return text


def _safe_value(value: Any) -> str:
    if _is_blank(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value).date().isoformat()
    return str(value)


def _row_key(row: pd.Series, key_cols: list[str], options: ComparisonOptions, key_side: str) -> tuple | None:
    values = []
    for col in key_cols:
        val = row.get(col, "")
        if _is_blank(val):
            return None
        values.append(_normalize(val, col, options))
    return tuple(values)


def _key_display(key: tuple | None) -> str:
    if key is None:
        return ""
    return " | ".join(str(v) for v in key)


def compare_frames(a_df: pd.DataFrame, b_df: pd.DataFrame, options: ComparisonOptions) -> dict[str, Any]:
    """Compare two frames. Changes is long-form: one row per field/rule exception."""
    if a_df is None or b_df is None or a_df.empty or b_df.empty:
        raise ValueError("Both input files must contain at least one data row.")
    if a_df.columns.duplicated().any() or b_df.columns.duplicated().any():
        raise ValueError("Duplicate column headers are not supported. Rename duplicate headers and retry.")
    if not options.key_mapping:
        raise ValueError("Select at least one matching key.")
    for a_key, b_key in options.key_mapping.items():
        if a_key not in a_df.columns:
            raise ValueError(f"Matching key '{a_key}' is missing from Source A.")
        if b_key not in b_df.columns:
            raise ValueError(f"Mapped matching key '{b_key}' is missing from Source B.")
    valid_mapping = {
        a: b for a, b in options.column_mapping.items()
        if a in a_df.columns and b in b_df.columns
    }
    for a_key, b_key in options.key_mapping.items():
        valid_mapping[a_key] = b_key

    a = a_df.copy().fillna("")
    b = b_df.copy().fillna("")
    a_key_cols = list(options.key_mapping.keys())
    b_key_cols = [options.key_mapping[k] for k in a_key_cols]

    a_keys = [_row_key(row, a_key_cols, options, "A") for _, row in a.iterrows()]
    b_keys = [_row_key(row, b_key_cols, options, "B") for _, row in b.iterrows()]

    a_counts = pd.Series([k for k in a_keys if k is not None], dtype=object).value_counts()
    b_counts = pd.Series([k for k in b_keys if k is not None], dtype=object).value_counts()
    dup_a = {k for k, count in a_counts.items() if count > 1}
    dup_b = {k for k, count in b_counts.items() if count > 1}

    # Key -> row index only for unique, non-duplicate keys.
    a_index = {k: i for i, k in enumerate(a_keys) if k is not None and k not in dup_a}
    b_index = {k: i for i, k in enumerate(b_keys) if k is not None and k not in dup_b}

    changes = []
    matched = []
    warnings = []

    def add_change(key, arow, brow, category, field, av, bv, severity, comment, action):
        changes.append({
            "Employer_ID": _find_value(arow, brow, ("Employer_ID", "Employer ID", "Group_ID", "Group ID")),
            "Employee_ID": _find_value(arow, brow, ("Employee_ID", "Employee ID", "EEID", "Worker_ID", "Worker ID")) or _key_display(key),
            "Employee_Name": _find_value(arow, brow, ("Employee_Name", "Employee Name", "Employee_First_Name", "First_Name", "First Name")),
            "Difference_Category": category,
            "Field_Name": field,
            "Vendor_Value": _safe_value(av),
            "BenAdmin_Value": _safe_value(bv),
            "Severity": severity,
            "Comment": comment,
            "Recommended_Action": action,
        })

    def arow_for(i):
        return a.iloc[i] if i is not None else None
    def brow_for(i):
        return b.iloc[i] if i is not None else None

    # Duplicates and missing keys are surfaced as explicit exceptions, never silently deduplicated.
    for i, key in enumerate(a_keys):
        if key is None:
            add_change(None, a.iloc[i], None, "Missing/Invalid Key", "Matching key", "", "", "High",
                       "Source A record has a blank value in at least one selected matching-key field. It cannot be matched reliably.",
                       "Verify the source record and matching-key configuration.")
        elif key in dup_a:
            add_change(key, a.iloc[i], None, "Duplicate Key", "Matching key", _key_display(key), "", "High",
                       "Duplicate matching key occurs in the vendor/client file. Automatic one-to-one reconciliation is ambiguous.",
                       "Investigate duplicate rows and confirm whether they represent duplicate transactions or distinct coverage records.")
    for i, key in enumerate(b_keys):
        if key is None:
            add_change(None, None, b.iloc[i], "Missing/Invalid Key", "Matching key", "", "", "High",
                       "BenAdmin record has a blank value in at least one selected matching-key field. It cannot be matched reliably.",
                       "Verify the source record and matching-key configuration.")
        elif key in dup_b:
            add_change(key, None, b.iloc[i], "Duplicate Key", "Matching key", "", _key_display(key), "High",
                       "Duplicate matching key occurs in the BenAdmin file. Automatic one-to-one reconciliation is ambiguous.",
                       "Investigate duplicate rows and confirm whether the export contains multiple coverage or transaction rows per employee.")

    # If a key is duplicated in either source, do not reconcile that key automatically.
    all_keys = set(a_index) | set(b_index)
    matched_keys = set(a_index) & set(b_index)
    for key in sorted(matched_keys, key=lambda x: tuple(str(y) for y in x)):
        ai, bi = a_index[key], b_index[key]
        ar, br = a.iloc[ai], b.iloc[bi]
        record_changes_start = len(changes)
        # Compare mapped fields only; selected keys are included as comparison fields.
        for acol, bcol in valid_mapping.items():
            av, bv = ar.get(acol, ""), br.get(bcol, "")
            an = _normalize(av, acol, options)
            bn = _normalize(bv, bcol, options)
            if an != bn:
                # SSN mismatch is especially sensitive and should never be echoed into comments.
                is_ssn = _norm_header(acol) in {"ssn", "socialsecuritynumber", "taxpayerid"} or _norm_header(bcol) in {"ssn", "socialsecuritynumber", "taxpayerid"}
                cat = "Identity Discrepancy" if is_ssn else _category_for_field(acol)
                sev = "Critical" if is_ssn else _severity_for_field(acol)
                comment = (
                    "Identity field differs between the two sources. Verify employee identity using the approved secure process; do not automatically overwrite either value."
                    if is_ssn else
                    f"{acol} differs between the vendor/client file and BenAdmin. Verify the source election or employee data and confirm which system should be corrected."
                )
                action = "Route through approved identity/data correction process." if is_ssn else "Review source documentation and correct the authoritative system if required."
                add_change(key, ar, br, cat, acol, av, bv, sev, comment, action)

        if options.run_rule_checks:
            for issue in _business_rule_checks(ar, br, valid_mapping, options):
                add_change(key, ar, br, issue["category"], issue["field"], issue["a"], issue["b"],
                           issue["severity"], issue["comment"], issue["action"])
        if len(changes) == record_changes_start:
            record = {"Matching_Key": _key_display(key)}
            for acol, bcol in valid_mapping.items():
                record[f"Vendor::{acol}"] = _safe_value(ar.get(acol, ""))
                record[f"BenAdmin::{bcol}"] = _safe_value(br.get(bcol, ""))
            matched.append(record)

    # Unmatched unique keys are only classified as missing if not ambiguous due to duplicates.
    for key, ai in a_index.items():
        if key not in b_index:
            add_change(key, a.iloc[ai], None, "Missing in BenAdmin", "Employee record", "Present", "Missing", "High",
                       "Employee is present in the vendor/client file but no unambiguous corresponding record was found in BenAdmin.",
                       "Verify file completeness, employee identifier mapping, and enrollment setup.")
    for key, bi in b_index.items():
        if key not in a_index:
            add_change(key, None, b.iloc[bi], "Missing in Vendor", "Employee record", "Missing", "Present", "High",
                       "Employee is present in BenAdmin but no unambiguous corresponding record was found in the vendor/client file.",
                       "Verify whether the vendor file is incomplete, the record is stale, or the employee is outside the expected population.")

    if options.report_unmapped_columns:
        mapped_a = set(valid_mapping.keys())
        mapped_b = set(valid_mapping.values())
        for col in a.columns:
            if col not in mapped_a:
                warnings.append({
                    "Employer_ID": "", "Employee_ID": "", "Employee_Name": "",
                    "Difference_Category": "Unmapped Column", "Field_Name": col,
                    "Vendor_Value": "Column present", "BenAdmin_Value": "",
                    "Severity": "Low",
                    "Comment": f"Vendor column '{col}' was not mapped to a BenAdmin column and was not compared.",
                    "Recommended_Action": "Map the column if it should be reconciled, or explicitly exclude it.",
                })
        for col in b.columns:
            if col not in mapped_b:
                warnings.append({
                    "Employer_ID": "", "Employee_ID": "", "Employee_Name": "",
                    "Difference_Category": "Unmapped Column", "Field_Name": col,
                    "Vendor_Value": "", "BenAdmin_Value": "Column present",
                    "Severity": "Low",
                    "Comment": f"BenAdmin column '{col}' was not mapped to a vendor column and was not compared.",
                    "Recommended_Action": "Map the column if it should be reconciled, or explicitly exclude it.",
                })

    changes_df = pd.DataFrame(changes, columns=[
        "Employer_ID", "Employee_ID", "Employee_Name", "Difference_Category", "Field_Name",
        "Vendor_Value", "BenAdmin_Value", "Severity", "Comment", "Recommended_Action"
    ])
    if warnings:
        changes_df = pd.concat([changes_df, pd.DataFrame(warnings)], ignore_index=True)
    match_df = pd.DataFrame(matched)
    if match_df.empty:
        match_df = pd.DataFrame(columns=["Matching_Key"] + [
            item for a_col, b_col in valid_mapping.items()
            for item in (f"Vendor::{a_col}", f"BenAdmin::{b_col}")
        ])
    return {
        "match": match_df,
        "changes": changes_df,
        "metadata": {
            "source_a_rows": len(a),
            "source_b_rows": len(b),
            "unique_matched_keys": len(matched_keys),
            "source_a_duplicate_keys": len(dup_a),
            "source_b_duplicate_keys": len(dup_b),
            "source_a_invalid_key_rows": sum(k is None for k in a_keys),
            "source_b_invalid_key_rows": sum(k is None for k in b_keys),
            "source_a_name": options.source_a_name,
            "source_b_name": options.source_b_name,
            "generated_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        },
    }


def _find_value(arow, brow, candidates):
    for row in (arow, brow):
        if row is None:
            continue
        for col in candidates:
            if col in row.index and not _is_blank(row[col]):
                return _safe_value(row[col])
    return ""


def _category_for_field(field_name: str) -> str:
    n = _norm_header(field_name)
    if any(x in n for x in ("termination", "employmentstatus", "status")):
        return "Employment/Termination"
    if any(x in n for x in ("medical", "dental", "vision", "plan", "tier", "coverage", "waive")):
        return "Benefit Enrollment"
    if any(x in n for x in ("effective", "eligibility", "hire", "birth", "date")):
        return "Date Discrepancy"
    if any(x in n for x in ("hsa", "fsa", "amount", "contribution", "deduction")):
        return "Election/Contribution"
    return "Field Mismatch"


def _severity_for_field(field_name: str) -> str:
    n = _norm_header(field_name)
    if any(x in n for x in ("termination", "employmentstatus", "ssn", "eligibility", "effective")):
        return "High"
    if any(x in n for x in ("plan", "tier", "coverage")):
        return "High"
    return "Medium"


def _business_rule_checks(arow, brow, mapping, options):
    issues = []
    # Run basic consistency checks against available fields; these are warnings, not definitive benefit determinations.
    def get(row, names):
        if row is None:
            return ""
        for col in row.index:
            if _norm_header(col) in names:
                return row[col]
        return ""
    status = str(get(arow, {"employmentstatus", "status"})).strip().casefold()
    term = get(arow, {"terminationdate", "termdate"})
    if status in {"active", "a"} and not _is_blank(term):
        issues.append({
            "category": "Data Quality Warning", "field": "Employment status / Termination date",
            "a": f"Status={get(arow, {'employmentstatus','status'})}; Termination_Date={term}",
            "b": "", "severity": "High",
            "comment": "Source A indicates Active status but also has a termination date populated. Confirm the source's status/date convention.",
            "action": "Validate employment status and termination date with the authoritative source.",
        })
    # Eligibility/effective date ordering, when both are available in Source A.
    elig = get(arow, {"benefiteligibilitydate", "eligibilitydate"})
    eff = get(arow, {"medicaleffectivedate", "healtheffectivedate", "effectivedate"})
    if not _is_blank(elig) and not _is_blank(eff):
        e1, e2 = pd.to_datetime(elig, errors="coerce"), pd.to_datetime(eff, errors="coerce")
        if not pd.isna(e1) and not pd.isna(e2) and e2 < e1:
            issues.append({
                "category": "Data Quality Warning", "field": "Eligibility vs effective date",
                "a": f"Eligibility={_safe_value(elig)}; Effective={_safe_value(eff)}",
                "b": "", "severity": "High",
                "comment": "Medical effective date precedes benefit eligibility date in Source A. Confirm whether retroactive coverage or a client-specific rule explains the sequence.",
                "action": "Check plan eligibility rules and enrollment documentation.",
            })
    # Extreme amount is an outlier warning only; no tax limit is assumed.
    for col in arow.index:
        n = _norm_header(col)
        if "hsa" in n and "election" in n and not _is_blank(arow[col]):
            try:
                amount = float(str(arow[col]).replace("$", "").replace(",", "").strip())
                if amount > 50000 or amount < 0:
                    issues.append({
                        "category": "Data Quality Warning", "field": col,
                        "a": _safe_value(arow[col]), "b": "",
                        "severity": "High",
                        "comment": "HSA election amount is outside the configurable starter outlier threshold. This is a QA warning, not a tax or eligibility determination.",
                        "action": "Confirm the intended election amount and apply the correct plan-year validation rules.",
                    })
            except ValueError:
                pass
    return issues


def build_report_workbook(result: dict[str, Any]) -> bytes:
    """Create an Excel workbook with Match and Changes sheets plus a summary sheet."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result["match"].to_excel(writer, index=False, sheet_name="Match")
        result["changes"].to_excel(writer, index=False, sheet_name="Changes")
        summary_rows = [
            ("Generated at (UTC)", result["metadata"].get("generated_at_utc", "")),
            ("Source A file", result["metadata"].get("source_a_name", "")),
            ("Source B file", result["metadata"].get("source_b_name", "")),
            ("Source A rows", result["metadata"].get("source_a_rows", 0)),
            ("Source B rows", result["metadata"].get("source_b_rows", 0)),
            ("Matched records with no reported differences", len(result["match"])),
            ("Changes rows (field/rule/column exceptions)", len(result["changes"])),
            ("Source A duplicate key groups", result["metadata"].get("source_a_duplicate_keys", 0)),
            ("Source B duplicate key groups", result["metadata"].get("source_b_duplicate_keys", 0)),
        ]
        pd.DataFrame(summary_rows, columns=["Metric", "Value"]).to_excel(writer, index=False, sheet_name="Summary")
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 32
            for col_idx, cells in enumerate(ws.columns, 1):
                width = 12
                for cell in list(cells)[:300]:
                    width = max(width, min(len(str(cell.value or "")) + 2, 48))
                ws.column_dimensions[get_column_letter(col_idx)].width = min(width, 42)
            if ws.title == "Changes":
                header_map = {cell.value: cell.column for cell in ws[1]}
                sev_col = header_map.get("Severity")
                if sev_col:
                    for row in range(2, ws.max_row + 1):
                        cell = ws.cell(row, sev_col)
                        if cell.value in {"Critical", "High"}:
                            cell.fill = PatternFill("solid", fgColor="FCE4D6")
                        elif cell.value == "Medium":
                            cell.fill = PatternFill("solid", fgColor="FFF2CC")
    return output.getvalue()


def summarize_result(result: dict[str, Any]) -> dict[str, int]:
    changes = result["changes"]
    def count_category(category):
        if changes.empty or "Difference_Category" not in changes:
            return 0
        return int((changes["Difference_Category"] == category).sum())
    return {
        "matched_records": len(result["match"]),
        "records_with_changes": int(changes["Employee_ID"].replace("", pd.NA).dropna().nunique()) if not changes.empty else 0,
        "exception_rows": len(changes),
        "duplicate_rows": count_category("Duplicate Key"),
        "source_a_only": count_category("Missing in BenAdmin"),
        "source_b_only": count_category("Missing in Vendor"),
        "warning_rows": int(changes["Difference_Category"].astype(str).str.contains("Warning|Unmapped", case=False, regex=True).sum()) if not changes.empty else 0,
    }

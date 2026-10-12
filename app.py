from __future__ import annotations

import io
import time
from datetime import datetime, timezone

import pandas as pd
import streamlit as st

from reconciliation_engine import (
    ComparisonOptions, compare_frames, load_uploaded_file,
    suggest_matching_keys, build_report_workbook, summarize_result,
)

st.set_page_config(
    page_title="Healthcare Enrollment Reconciliation",
    page_icon="🩺",
    layout="wide",
)

st.title("Healthcare Enrollment Reconciliation Tool")
st.caption("Compare vendor/client enrollment data against a Benefit Administration system extract.")
st.info(
    "Synthetic-data demo. Do not upload real SSNs, dates of birth, or other employee information. "
    "Use only synthetic or explicitly approved test data."
)

with st.expander("How it works", expanded=False):
    st.markdown(
        "1. Upload the vendor/client file and BenAdmin extract.\n"
        "2. Choose worksheets, confirm suggested matching keys, and map equivalent columns if needed.\n"
        "3. Run reconciliation and download an Excel workbook with **Match** and **Changes** sheets.\n\n"
        "The app processes files temporarily in this session and does not intentionally persist uploads or reports."
    )

left, right = st.columns(2)
with left:
    st.subheader("1. Vendor / Client file")
    source_a_file = st.file_uploader(
        "Upload vendor/client enrollment file",
        type=["xlsx", "xls", "csv"],
        key="source_a",
        help="Supported: .xlsx, .xls (when dependency is installed), .csv",
    )
with right:
    st.subheader("2. BenAdmin file")
    source_b_file = st.file_uploader(
        "Upload Benefit Administration extract",
        type=["xlsx", "xls", "csv"],
        key="source_b",
        help="Supported: .xlsx, .xls (when dependency is installed), .csv",
    )

def get_sheet_names(uploaded):
    if uploaded is None:
        return []
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return ["CSV"]
    try:
        uploaded.seek(0)
        return pd.ExcelFile(uploaded).sheet_names
    except Exception:
        return []

def read_preview(uploaded, sheet):
    uploaded.seek(0)
    return load_uploaded_file(uploaded, sheet_name=None if sheet == "CSV" else sheet)

a_df = b_df = None
a_sheet = b_sheet = None
if source_a_file and source_b_file:
    a_sheets = get_sheet_names(source_a_file)
    b_sheets = get_sheet_names(source_b_file)
    if not a_sheets or not b_sheets:
        st.error("Could not read worksheet metadata. Check that both files are valid and not password-protected.")
        st.stop()
    c1, c2 = st.columns(2)
    with c1:
        a_sheet = st.selectbox("Vendor worksheet", a_sheets, key="a_sheet")
    with c2:
        b_sheet = st.selectbox("BenAdmin worksheet", b_sheets, key="b_sheet")
    try:
        a_df = read_preview(source_a_file, a_sheet)
        b_df = read_preview(source_b_file, b_sheet)
    except Exception as exc:
        st.error(f"Unable to read one of the selected inputs: {exc}")
        st.stop()

    st.subheader("File overview")
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Vendor rows", f"{len(a_df):,}")
    p2.metric("BenAdmin rows", f"{len(b_df):,}")
    p3.metric("Vendor columns", len(a_df.columns))
    p4.metric("BenAdmin columns", len(b_df.columns))

    with st.expander("Preview source files", expanded=False):
        st.markdown("**Vendor / Client preview**")
        st.dataframe(a_df.head(10), use_container_width=True)
        st.markdown("**BenAdmin preview**")
        st.dataframe(b_df.head(10), use_container_width=True)

    st.subheader("3. Confirm matching and comparison settings")
    suggestions = suggest_matching_keys(a_df, b_df)
    a_cols = list(a_df.columns)
    b_cols = list(b_df.columns)
    common_cols = [c for c in a_cols if c in b_cols]

    st.write("**Suggested matching keys**")
    if suggestions:
        suggested_label = " + ".join(suggestions[0])
        st.caption(f"Best heuristic suggestion: {suggested_label}. Suggestions are not guaranteed; verify before comparing.")
        key_defaults = suggestions[0]
    else:
        st.warning("No strong matching key could be suggested. Select equivalent columns manually.")
        key_defaults = []

    key_options = list(dict.fromkeys(a_cols + b_cols))
    selected_keys = st.multiselect(
        "Matching key column(s) from the Vendor file",
        options=a_cols,
        default=[k for k in key_defaults if k in a_cols],
        help="Select one or more stable identifiers. Employee ID plus Employer/Group ID is usually preferable.",
    )

    # Map each selected vendor key to a BenAdmin key.
    key_mapping = {}
    if selected_keys:
        map_cols = st.columns(min(3, len(selected_keys)))
        for i, key in enumerate(selected_keys):
            with map_cols[i % len(map_cols)]:
                options = b_cols
                default_idx = options.index(key) if key in options else 0
                mapped = st.selectbox(
                    f"BenAdmin equivalent for '{key}'",
                    options=options,
                    index=default_idx,
                    key=f"keymap_{key}",
                )
                key_mapping[key] = mapped

    st.markdown("**Column mapping**")
    st.caption("Exact matching headers are mapped automatically. Add or change mappings where equivalent fields have different names.")
    mapping_rows = []
    for col in a_cols:
        default = col if col in b_cols else None
        mapping_rows.append((col, default if default is not None else "(Do not compare)"))
    mapping_df = pd.DataFrame(mapping_rows, columns=["Vendor column", "Suggested BenAdmin column"])
    editable = st.data_editor(
        mapping_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Vendor column": st.column_config.TextColumn(disabled=True),
            "Suggested BenAdmin column": st.column_config.SelectboxColumn(
                options=["(Do not compare)"] + b_cols,
                required=True,
            ),
        },
        key="column_mapping_editor",
    )
    column_mapping = {}
    for _, row in editable.iterrows():
        vcol = row["Vendor column"]
        bcol = row["Suggested BenAdmin column"]
        if pd.notna(bcol) and bcol != "(Do not compare)":
            column_mapping[vcol] = bcol

    # Ensure selected key mappings are reflected in the overall mapping.
    for vkey, bkey in key_mapping.items():
        column_mapping[vkey] = bkey

    opts_col1, opts_col2, opts_col3 = st.columns(3)
    with opts_col1:
        normalize_whitespace = st.checkbox("Trim whitespace", value=True)
        normalize_case = st.checkbox("Ignore case for text comparison", value=False)
    with opts_col2:
        normalize_dates = st.checkbox("Normalize common date formats", value=True)
        normalize_numbers = st.checkbox("Normalize numeric representations", value=True)
    with opts_col3:
        missing_column_policy = st.selectbox(
            "Unmapped/missing columns",
            ["Report as warning", "Ignore unmapped columns"],
            index=0,
        )
        include_rule_checks = st.checkbox("Run basic data-quality checks", value=True)

    st.caption(
        "Business-rule checks are configurable starter checks, not a substitute for client-specific eligibility, "
        "plan-year, waiting-period, or termination rules."
    )

    max_rows = st.number_input(
        "Maximum rows per input file",
        min_value=100,
        max_value=500000,
        value=100000,
        step=1000,
    )

    can_compare = bool(selected_keys and key_mapping and all(k in column_mapping for k in selected_keys))
    if st.button("Compare enrollment files", type="primary", disabled=not can_compare, use_container_width=True):
        if len(a_df) > max_rows or len(b_df) > max_rows:
            st.error(f"One or both files exceed the configured row limit of {max_rows:,}.")
            st.stop()
        if a_df.empty or b_df.empty:
            st.error("Both files must contain at least one data row.")
            st.stop()
        try:
            progress = st.progress(0, text="Validating input files…")
            status = st.empty()
            status.write("Validating column mappings and matching keys")
            progress.progress(10, text="Validating input files…")
            time.sleep(0.1)
            status.write("Preparing records and normalized comparison values")
            progress.progress(25, text="Preparing data…")
            time.sleep(0.1)
            status.write("Matching employee records and comparing selected fields")
            progress.progress(60, text="Comparing records…")
            options = ComparisonOptions(
                key_mapping=key_mapping,
                column_mapping=column_mapping,
                normalize_whitespace=normalize_whitespace,
                normalize_case=normalize_case,
                normalize_dates=normalize_dates,
                normalize_numbers=normalize_numbers,
                run_rule_checks=include_rule_checks,
                report_unmapped_columns=(missing_column_policy == "Report as warning"),
                source_a_name=source_a_file.name,
                source_b_name=source_b_file.name,
            )
            result = compare_frames(a_df, b_df, options)
            status.write("Generating formatted Excel report")
            progress.progress(85, text="Generating report…")
            workbook = build_report_workbook(result)
            progress.progress(100, text="Completed")
            status.success("Comparison completed. Your report is ready.")
            summary = summarize_result(result)
            st.session_state["recon_result"] = result
            st.session_state["recon_workbook"] = workbook
            st.session_state["recon_summary"] = summary
            st.session_state["recon_filename"] = (
                f"enrollment_reconciliation_{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}.xlsx"
            )
        except Exception as exc:
            st.error(f"Comparison could not be completed: {exc}")
            st.caption("Check the matching keys, column mappings, and file structure, then retry.")

if "recon_result" in st.session_state and "recon_workbook" in st.session_state:
    st.divider()
    st.subheader("4. Reconciliation results")
    summary = st.session_state["recon_summary"]
    cols = st.columns(4)
    cols[0].metric("Matched records", f"{summary['matched_records']:,}")
    cols[1].metric("Records with changes", f"{summary['records_with_changes']:,}")
    cols[2].metric("Field/rule exceptions", f"{summary['exception_rows']:,}")
    cols[3].metric("Duplicate-key rows", f"{summary['duplicate_rows']:,}")
    cols2 = st.columns(3)
    cols2[0].metric("Vendor-only records", f"{summary['source_a_only']:,}")
    cols2[1].metric("BenAdmin-only records", f"{summary['source_b_only']:,}")
    cols2[2].metric("Warnings", f"{summary['warning_rows']:,}")
    st.download_button(
        "Download Excel reconciliation report",
        data=st.session_state["recon_workbook"],
        file_name=st.session_state["recon_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    result = st.session_state["recon_result"]
    with st.expander("Preview Match sheet"):
        st.dataframe(result["match"].head(100), use_container_width=True)
    with st.expander("Preview Changes sheet"):
        st.dataframe(result["changes"].head(200), use_container_width=True)

st.divider()
st.caption("Demo build · Synthetic data only · No intentional long-term file retention")

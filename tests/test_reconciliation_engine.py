import io
import pandas as pd
import pytest
from reconciliation_engine import (
    ComparisonOptions, compare_frames, build_report_workbook,
    suggest_matching_keys, summarize_result, load_uploaded_file,
)

def opts():
    return ComparisonOptions(
        key_mapping={"Employer_ID":"Employer_ID", "Employee_ID":"Employee_ID"},
        column_mapping={
            "Employer_ID":"Employer_ID",
            "Employee_ID":"Employee_ID",
            "Employment_Status":"Employment_Status",
            "Medical_Plan":"Medical_Plan",
            "Medical_Effective_Date":"Medical_Effective_Date",
        },
        run_rule_checks=True,
        report_unmapped_columns=False,
    )

def test_exact_match_goes_to_match_sheet():
    a = pd.DataFrame([{
        "Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active",
        "Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"
    }])
    b = a.copy()
    result = compare_frames(a,b,opts())
    assert len(result["match"]) == 1
    assert result["changes"].empty

def test_plan_mismatch_is_reported():
    a = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    b = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"HDHP","Medical_Effective_Date":"2026-01-01"}])
    result = compare_frames(a,b,opts())
    assert len(result["match"]) == 0
    assert (result["changes"]["Field_Name"] == "Medical_Plan").any()

def test_missing_in_benadmin():
    a = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    b = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E2","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    result = compare_frames(a,b,opts())
    assert (result["changes"]["Difference_Category"] == "Missing in BenAdmin").any()
    assert (result["changes"]["Difference_Category"] == "Missing in Vendor").any()

def test_duplicate_key_is_not_silently_matched():
    a = pd.DataFrame([
        {"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"},
        {"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"HDHP","Medical_Effective_Date":"2026-01-01"},
    ])
    b = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    result = compare_frames(a,b,opts())
    assert (result["changes"]["Difference_Category"] == "Duplicate Key").any()
    assert len(result["match"]) == 0

def test_blank_key_is_reported():
    a = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    b = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    result = compare_frames(a,b,opts())
    assert (result["changes"]["Difference_Category"] == "Missing/Invalid Key").any()

def test_whitespace_normalization():
    a = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":" E1 ","Employment_Status":"Active","Medical_Plan":"PPO ","Medical_Effective_Date":"2026-01-01"}])
    b = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    result = compare_frames(a,b,opts())
    assert len(result["match"]) == 1

def test_date_normalization():
    a = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"01/01/2026"}])
    b = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    result = compare_frames(a,b,opts())
    assert len(result["match"]) == 1

def test_active_with_termination_date_warning():
    a = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Termination_Date":"2026-09-01","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    b = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Termination_Date":"","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    o = opts()
    o.column_mapping["Termination_Date"] = "Termination_Date"
    result = compare_frames(a,b,o)
    assert (result["changes"]["Difference_Category"] == "Data Quality Warning").any()

def test_report_has_match_changes_summary_sheets():
    a = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1","Employment_Status":"Active","Medical_Plan":"PPO","Medical_Effective_Date":"2026-01-01"}])
    result = compare_frames(a,a.copy(),opts())
    content = build_report_workbook(result)
    book = pd.ExcelFile(io.BytesIO(content))
    assert book.sheet_names == ["Match","Changes","Summary"]

def test_key_suggestions_prefer_group_plus_employee_id():
    a = pd.DataFrame(columns=["Employer_ID","Employee_ID","SSN"])
    b = pd.DataFrame(columns=["Employer_ID","Employee_ID","SSN"])
    suggestions = suggest_matching_keys(a,b)
    assert ["Employer_ID","Employee_ID"] in suggestions

def test_empty_input_raises():
    a = pd.DataFrame(columns=["Employer_ID","Employee_ID"])
    b = pd.DataFrame([{"Employer_ID":"G1","Employee_ID":"E1"}])
    with pytest.raises(ValueError):
        compare_frames(a,b,opts())

def test_csv_loader_keeps_identifiers_as_strings():
    class FakeUpload:
        name = "test.csv"
        def __init__(self, data): self.data = data
        def seek(self, n): pass
        def getvalue(self): return self.data
    df = load_uploaded_file(FakeUpload(b"Employee_ID,SSN\n00123,900-10-1001\n"))
    assert df.iloc[0]["Employee_ID"] == "00123"

import pandas as pd
from datetime import date
from src.generate_report import compute_kpis

def test_compute_kpis_basic():
    data = {
        "claim_id": ["A", "B", "C"],
        "branch": ["X", "X", "Y"],
        "line_of_service": ["Mitigation", "Mitigation", "Reconstruction"],
        "is_assignment": ["Yes", "No", "Yes"],
        "received_date": [date(2025,9,22)]*3,
        "assigned_pm": ["Alex", "", "Jamie"],
        "assigned_date": [date(2025,9,22), None, date(2025,9,23)],
        "status": ["Completed","New","In Progress"],
        "dash_job_id": ["D-1","","D-2"],
        "completed_date": [date(2025,9,25), None, None],
    }
    df = pd.DataFrame(data)
    kpis, by_branch, by_service, by_pm = compute_kpis(df, 1, 7)
    assert kpis["Total Claims"] == 3
    assert kpis["Status: Completed"] == 1
    assert by_branch["count"].sum() == 3

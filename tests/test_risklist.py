#tests/test_risklist.py
from excel_merge.risklist import RiskList

def test_load_from_excel_book():
    min_row = 2
    min_col = 2
    file_name = "./tests/school_members_from.xlsx"
    sheet_name = "Sheet1"

    risk_list = RiskList()
    risk_list.load_from_excel_book(file_name, sheet_name, min_row, min_col)

    assert risk_list.min_row == min_row
    assert risk_list.min_col == min_col
    assert risk_list.file_name == file_name
    assert risk_list.sheet_name == sheet_name


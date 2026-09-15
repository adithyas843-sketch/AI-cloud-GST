import pandas as pd
from validators import valid_gstin, validation_report
def test_gstin_structure():
    assert valid_gstin("27ABCDE1234F1Z5")
    assert not valid_gstin("27ABCDE1234F1Z")
    assert not valid_gstin("00ABCDE1234F1Z5")
    assert not valid_gstin(pd.NA)
def test_duplicate_detection():
    df=pd.DataFrame({"GSTIN":["27ABCDE1234F1Z5"]*2,"Invoice Number":["A1"]*2,"Invoice Date":["2026-01-01"]*2,"Taxable Value":[10]*2,"IGST":[0]*2,"CGST":[1]*2,"SGST":[1]*2})
    assert "Duplicate invoice" in validation_report(df).Issue.tolist()

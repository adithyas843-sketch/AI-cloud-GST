import pandas as pd
from reconciliation import reconcile, health_score
from sample_data_generator import generate_samples
def source(value=100):
 return pd.DataFrame({"GSTIN":["27ABCDE1234F1Z5"],"Invoice Number":["A1"],"Invoice Date":["2026-01-01"],"Customer":["A"],"Taxable Value":[value],"IGST":[18],"CGST":[0],"SGST":[0]})
def test_invoice_matching_and_exposure():
    r=reconcile({"GSTR-1":source(),"Tally":source(120)})
    assert r.iloc[0]["Status"] == "Mismatch"
    assert r.iloc[0]["Taxable Difference"] == 20
    assert r.iloc[0]["Exposure Amount"] == 0
def test_health_score():
    r=reconcile({"GSTR-1":source(),"Tally":source()})
    score,_=health_score(r,pd.DataFrame(columns=["Issue"]))
    assert score == 100
def test_gstin_mismatch_is_classified_not_just_missing():
    tally=source(); tally.loc[0,"GSTIN"]="29ABCDE1234F1Z5"
    r=reconcile({"GSTR-1":source(),"Tally":tally})
    assert "GSTIN mismatch with Tally" in r.iloc[0]["Mismatch Type"]
def test_invoice_number_mismatch_is_classified():
    tally=source(); tally.loc[0,"Invoice Number"]="A2"
    r=reconcile({"GSTR-1":source(),"Tally":tally})
    assert "Invoice number mismatch with Tally" in r.iloc[0]["Mismatch Type"]
def test_sample_generator_writes_all_workbooks(tmp_path):
    files=generate_samples(tmp_path)
    assert len(files) == 4
    assert all(path.exists() for path in files)
    assert list(pd.read_excel(files[0]).columns) == ["GSTIN", "Invoice Number", "Invoice Date", "Customer", "Taxable Value", "IGST", "CGST", "SGST"]
    result=reconcile({"GSTR-1":pd.read_excel(files[0]),"Tally":pd.read_excel(files[1]),"Zoho":pd.read_excel(files[2])})
    assert (result["Status"] == "Matched").any()

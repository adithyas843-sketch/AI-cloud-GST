"""Create demonstration GST source workbooks."""
from pathlib import Path
import pandas as pd

BASE = [
 ["27ABCDE1234F1Z5","INV-1001","2026-04-05","Aster Retail",100000,18000,0,0],
 ["29PQRSX6789L1Z2","INV-1002","2026-04-08","Bluebird Foods",50000,0,4500,4500],
 ["07AAACA1234A1ZP","INV-1003","2026-05-14","Cedar Labs",75000,13500,0,0],
 ["33AAAAA0000A1Z5","INV-1004","2026-05-20","Delta Stores",25000,0,2250,2250],
 ["27ABCDE1234F1Z5","INV-1005","2026-06-11","Aster Retail",120000,21600,0,0],
]
COLUMNS=["GSTIN","Invoice Number","Invoice Date","Customer","Taxable Value","IGST","CGST","SGST"]
def generate_samples(directory="."):
    p=Path(directory); p.mkdir(parents=True,exist_ok=True)
    gstr1=pd.DataFrame(BASE, columns=COLUMNS)
    tally=pd.DataFrame(BASE[:-1]+[["29PQRSX6789L1Z2","INV-1006","2026-06-17","Bluebird Foods",40000,0,3600,3600]], columns=COLUMNS)
    tally.loc[2,"Taxable Value"]=70000; tally.loc[3,"Invoice Date"]="2026-05-22"
    # Duplicate INV-1006 while leaving INV-1001 as a clean cross-source match.
    tally=pd.concat([tally,tally.iloc[[-1]]],ignore_index=True)
    zoho=pd.DataFrame(BASE, columns=COLUMNS); zoho.loc[1,"GSTIN"]="29PQRSX6789L1Z9"
    gstr3b=pd.DataFrame({"Period":["Apr-2026","May-2026","Jun-2026"],"IGST":[18000,13500,21600],"CGST":[4500,2250,0],"SGST":[4500,2250,0]})
    gstr1.to_excel(p/"sample_gstr1.xlsx",index=False); tally.to_excel(p/"sample_tally.xlsx",index=False)
    zoho.to_excel(p/"sample_zoho.xlsx",index=False); gstr3b.to_excel(p/"sample_gstr3b.xlsx",index=False)
    return [p/f"sample_{n}.xlsx" for n in ["gstr1","tally","zoho","gstr3b"]]
if __name__ == "__main__": print("\n".join(map(str,generate_samples())))

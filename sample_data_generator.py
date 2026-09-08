from pathlib import Path

import pandas as pd


OUTPUT_DIR = Path("sample_data")
OUTPUT_DIR.mkdir(exist_ok=True)


gstr1 = pd.DataFrame(
    [
        {
            "Invoice Number": "INV-1001",
            "Invoice Date": "2026-04-01",
            "Customer Name": "Acme Manufacturing",
            "GSTIN": "27AACCA1111A1Z5",
            "Taxable Value": 100000,
            "IGST": 18000,
            "CGST": 0,
            "SGST": 0,
        },
        {
            "Invoice Number": "INV-1002",
            "Invoice Date": "2026-04-02",
            "Customer Name": "Beta Retail",
            "GSTIN": "29AABCB2222B1Z6",
            "Taxable Value": 50000,
            "IGST": 0,
            "CGST": 4500,
            "SGST": 4500,
        },
        {
            "Invoice Number": "INV-1003",
            "Invoice Date": "2026-04-03",
            "Customer Name": "Gamma Services",
            "GSTIN": "07AACCG3333C1Z7",
            "Taxable Value": 75000,
            "IGST": 13500,
            "CGST": 0,
            "SGST": 0,
        },
        {
            "Invoice Number": "INV-1004",
            "Invoice Date": "2026-04-04",
            "Customer Name": "Delta Traders",
            "GSTIN": "24AACCD4444D1Z8",
            "Taxable Value": 90000,
            "IGST": 0,
            "CGST": 8100,
            "SGST": 8100,
        },
        {
            "Invoice Number": "INV-1005",
            "Invoice Date": "2026-04-05",
            "Customer Name": "Epsilon Foods",
            "GSTIN": "06AACEF5555E1Z9",
            "Taxable Value": 40000,
            "IGST": 7200,
            "CGST": 0,
            "SGST": 0,
        },
    ]
)

tally = pd.DataFrame(
    [
        {
            "Voucher No": "INV-1001",
            "Voucher Date": "2026-04-01",
            "Party Name": "Acme Manufacturing",
            "Party GSTIN": "27AACCA1111A1Z5",
            "Taxable Amount": 100000,
            "IGST Amount": 18000,
            "CGST Amount": 0,
            "SGST Amount": 0,
        },
        {
            "Voucher No": "INV-1002",
            "Voucher Date": "2026-04-02",
            "Party Name": "Beta Retail",
            "Party GSTIN": "29AABCB2222B1Z6",
            "Taxable Amount": 52000,
            "IGST Amount": 0,
            "CGST Amount": 4680,
            "SGST Amount": 4680,
        },
        {
            "Voucher No": "INV-1003",
            "Voucher Date": "2026-04-03",
            "Party Name": "Gamma Services",
            "Party GSTIN": "07AACCG3333C1Z7",
            "Taxable Amount": 75000,
            "IGST Amount": 13500,
            "CGST Amount": 0,
            "SGST Amount": 0,
        },
        {
            "Voucher No": "INV-1004",
            "Voucher Date": "2026-04-04",
            "Party Name": "Delta Traders",
            "Party GSTIN": "24WRONG4444D1Z8",
            "Taxable Amount": 90000,
            "IGST Amount": 0,
            "CGST Amount": 8100,
            "SGST Amount": 8100,
        },
        {
            "Voucher No": "INV-1006",
            "Voucher Date": "2026-04-06",
            "Party Name": "Zeta Components",
            "Party GSTIN": "27AACFZ6666F1Z0",
            "Taxable Amount": 30000,
            "IGST Amount": 5400,
            "CGST Amount": 0,
            "SGST Amount": 0,
        },
        {
            "Voucher No": "INV-1006",
            "Voucher Date": "2026-04-06",
            "Party Name": "Zeta Components",
            "Party GSTIN": "27AACFZ6666F1Z0",
            "Taxable Amount": 30000,
            "IGST Amount": 5400,
            "CGST Amount": 0,
            "SGST Amount": 0,
        },
    ]
)

gstr1.to_excel(OUTPUT_DIR / "sample_gstr1.xlsx", index=False)
tally.to_excel(OUTPUT_DIR / "sample_tally_sales_register.xlsx", index=False)

print(f"Created files in {OUTPUT_DIR.resolve()}")

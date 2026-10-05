import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const templates = {
  "sample_sales_register.xlsx": [["Document Type", "GSTIN", "Voucher Number", "Voucher Date", "Customer Name", "Taxable Value", "IGST", "CGST", "SGST"], ["Invoice", "27ABCDE1234F1Z5", "INV-1001", new Date("2026-04-05"), "Aster Retail", 100000, 18000, 0, 0], ["Credit Note", "29PQRSX6789L1Z2", "CN-1002", new Date("2026-04-08"), "Bluebird Foods", 5000, 900, 0, 0]],
  "sample_purchase_register.xlsx": [["Document Type", "GSTIN", "Voucher Number", "Voucher Date", "Vendor Name", "Taxable Value", "IGST", "CGST", "SGST"], ["Invoice", "27ABCDE1234F1Z5", "PINV-1001", new Date("2026-04-05"), "Aster Supplies", 85000, 15300, 0, 0], ["Debit Note", "29PQRSX6789L1Z2", "DN-1002", new Date("2026-04-08"), "Bluebird Traders", 5000, 0, 450, 450]],
  "sample_gstr1.xlsx": [["Document Type", "Customer GSTIN", "Invoice No", "Date", "Customer Name", "Taxable Value", "IGST", "CGST", "SGST"], ["Invoice", "27ABCDE1234F1Z5", "INV-1001", new Date("2026-04-05"), "Aster Retail", 100000, 18000, 0, 0], ["Credit Note", "29PQRSX6789L1Z2", "CN-1002", new Date("2026-04-08"), "Bluebird Foods", 5000, 900, 0, 0]],
  "sample_gstr2b.xlsx": [["Document Type", "Supplier GSTIN", "Invoice No", "Invoice Date", "Supplier Name", "Taxable Value", "IGST", "CGST", "SGST"], ["Invoice", "27ABCDE1234F1Z5", "PINV-1001", new Date("2026-04-05"), "Aster Supplies", 85000, 15300, 0, 0], ["Debit Note", "29PQRSX6789L1Z2", "DN-1002", new Date("2026-04-08"), "Bluebird Traders", 5000, 0, 450, 450]],
  "sample_gstr3b.xlsx": [["Month", "Taxable Value", "IGST", "CGST", "SGST"], ["Apr-2026", 150000, 18000, 4500, 4500], ["May-2026", 75000, 13500, 0, 0]],
};

await fs.mkdir("templates", { recursive: true });
for (const [filename, values] of Object.entries(templates)) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Sample Data");
  const range = sheet.getRangeByIndexes(0, 0, values.length, values[0].length);
  range.values = values;
  sheet.getRangeByIndexes(0, 0, 1, values[0].length).format = { fill: "#1F4E78", font: { name: "Book Antiqua", bold: true, color: "#FFFFFF" }, borders: { preset: "all", style: "thin", color: "#D9E2F3" } };
  sheet.getRangeByIndexes(1, 0, values.length - 1, values[0].length).format = { font: { name: "Book Antiqua" }, borders: { preset: "all", style: "thin", color: "#D9D9D9" } };
  sheet.getUsedRange().format.autofitColumns();
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
  workbook.recalculate();
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(`templates/${filename}`);
}

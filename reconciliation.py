from __future__ import annotations

import re
from io import BytesIO
from typing import Any

import numpy as np
import pandas as pd


CANONICAL_FIELDS = [
    "invoice_number",
    "invoice_date",
    "customer_name",
    "gstin",
    "taxable_value",
    "igst",
    "cgst",
    "sgst",
]

FIELD_ALIASES = {
    "invoice_number": [
        "invoice number",
        "invoice no",
        "invoice no.",
        "invoice #",
        "invoice",
        "bill no",
        "bill number",
        "document number",
        "document no",
        "voucher number",
        "voucher no",
        "inv no",
        "inv number",
    ],
    "invoice_date": [
        "invoice date",
        "date",
        "document date",
        "bill date",
        "voucher date",
    ],
    "customer_name": [
        "customer",
        "customer name",
        "party",
        "party name",
        "buyer",
        "buyer name",
        "recipient",
        "recipient name",
        "name",
    ],
    "gstin": [
        "gstin",
        "gstin/uin",
        "gstin uin",
        "gst number",
        "gst no",
        "party gstin",
        "customer gstin",
        "buyer gstin",
        "recipient gstin",
    ],
    "taxable_value": [
        "taxable value",
        "taxable amount",
        "taxable",
        "assessable value",
        "net taxable value",
        "taxable value rs",
        "taxable amount rs",
    ],
    "igst": [
        "igst",
        "igst amount",
        "integrated tax",
        "integrated tax amount",
        "igst amount rs",
    ],
    "cgst": [
        "cgst",
        "cgst amount",
        "central tax",
        "central tax amount",
        "cgst amount rs",
    ],
    "sgst": [
        "sgst",
        "sgst amount",
        "state tax",
        "state tax amount",
        "utgst",
        "sgst/utgst",
        "sgst amount rs",
    ],
}

MISMATCH_LABELS = {
    "missing_in_gstr1": "Missing invoice in GSTR-1",
    "missing_in_tally": "Missing invoice in Tally",
    "gstin_mismatch": "GSTIN mismatch",
    "invoice_number_mismatch": "Invoice number mismatch",
    "taxable_value_mismatch": "Taxable value mismatch",
    "igst_mismatch": "IGST mismatch",
    "cgst_mismatch": "CGST mismatch",
    "sgst_mismatch": "SGST mismatch",
    "duplicate_invoice": "Duplicate invoice",
}


def normalize_header(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_invoice_number(value: Any) -> str:
    if pd.isna(value):
        return ""

    text = str(value).strip().upper()
    text = re.sub(r"\.0$", "", text)
    text = re.sub(r"[^A-Z0-9]", "", text)
    return text


def normalize_gstin(value: Any) -> str:
    if pd.isna(value):
        return ""

    return re.sub(r"\s+", "", str(value).strip().upper())


def numeric_value(value: Any) -> float:
    if pd.isna(value) or value == "":
        return 0.0

    if isinstance(value, str):
        value = value.replace(",", "").replace("₹", "").strip()
        if value in {"-", "NA", "N/A", "NONE", "NAN"}:
            return 0.0

    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def read_excel_file(uploaded_file: Any) -> pd.DataFrame:
    uploaded_file.seek(0)
    data = uploaded_file.read()
    workbook = pd.ExcelFile(BytesIO(data), engine="openpyxl")

    for sheet_name in workbook.sheet_names:
        candidate = pd.read_excel(
            BytesIO(data),
            sheet_name=sheet_name,
            engine="openpyxl",
        )
        candidate = candidate.dropna(how="all").dropna(axis=1, how="all")

        if not candidate.empty:
            return candidate

    return pd.DataFrame()


def infer_column_mapping(columns: list[str]) -> dict[str, str | None]:
    normalized_columns = {
        column: normalize_header(column)
        for column in columns
    }

    mapping: dict[str, str | None] = {}

    for field, aliases in FIELD_ALIASES.items():
        alias_set = {normalize_header(alias) for alias in aliases}
        selected = None

        for column, normalized in normalized_columns.items():
            if normalized in alias_set:
                selected = column
                break

        if selected is None:
            for column, normalized in normalized_columns.items():
                if any(
                    alias in normalized or normalized in alias
                    for alias in alias_set
                ):
                    selected = column
                    break

        mapping[field] = selected

    return mapping


def standardize_dataframe(
    dataframe: pd.DataFrame,
    mapping: dict[str, str | None],
) -> pd.DataFrame:
    result = pd.DataFrame(index=dataframe.index)

    for field in CANONICAL_FIELDS:
        source_column = mapping.get(field)

        if source_column and source_column in dataframe.columns:
            result[field] = dataframe[source_column]
        else:
            result[field] = ""

    result["invoice_number"] = result["invoice_number"].map(
        normalize_invoice_number
    )
    result["gstin"] = result["gstin"].map(normalize_gstin)
    result["invoice_date"] = pd.to_datetime(
        result["invoice_date"],
        errors="coerce",
    )
    result["taxable_value"] = result["taxable_value"].map(numeric_value)
    result["igst"] = result["igst"].map(numeric_value)
    result["cgst"] = result["cgst"].map(numeric_value)
    result["sgst"] = result["sgst"].map(numeric_value)

    result["customer_name"] = (
        result["customer_name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    result = result[
        result["invoice_number"].astype(bool)
        | (result["taxable_value"] != 0)
        | (result["igst"] != 0)
        | (result["cgst"] != 0)
        | (result["sgst"] != 0)
    ].copy()

    result["source_row"] = result.index + 2
    result["invoice_key"] = result["invoice_number"]

    return result.reset_index(drop=True)


def find_duplicate_rows(dataframe: pd.DataFrame) -> set[str]:
    duplicated = dataframe[
        dataframe["invoice_key"].ne("")
        & dataframe["invoice_key"].duplicated(keep=False)
    ]

    return set(duplicated["invoice_key"].tolist())


def amount_matches(
    left: float,
    right: float,
    tolerance: float,
) -> bool:
    return bool(np.isclose(left, right, atol=tolerance, rtol=0))


def explanation_for_issue(
    issue_type: str,
    gstr1_row: pd.Series | None,
    tally_row: pd.Series | None,
    difference: float = 0.0,
) -> str:
    if issue_type == "missing_in_gstr1":
        return (
            "This invoice exists in the Tally sales register but was not found "
            "in GSTR-1. Verify whether it was omitted from the return or filed "
            "under a different invoice number."
        )

    if issue_type == "missing_in_tally":
        return (
            "This invoice exists in GSTR-1 but was not found in the Tally sales "
            "register. Verify the accounting entry, cancellation status, and "
            "source register export."
        )

    if issue_type == "gstin_mismatch":
        return (
            "The buyer GSTIN differs between the two sources. Confirm the "
            "customer master and correct the GSTIN before relying on the "
            "invoice for input tax credit or outward-supply reporting."
        )

    if issue_type == "invoice_number_mismatch":
        return (
            "The invoice identifier does not align between the sources. Check "
            "prefixes, separators, leading zeroes, amended invoices, and credit "
            "note references."
        )

    if issue_type == "taxable_value_mismatch":
        return (
            f"Taxable value differs by Rs {abs(difference):,.2f}. Review "
            "discounts, freight, rounding, amendments, and whether both files "
            "use the same invoice-level aggregation."
        )

    if issue_type in {"igst_mismatch", "cgst_mismatch", "sgst_mismatch"}:
        tax_name = issue_type.replace("_mismatch", "").upper()
        return (
            f"{tax_name} differs between the two sources by "
            f"Rs {abs(difference):,.2f}. Review tax rate, place of supply, "
            "rounding, and whether the invoice was amended or split across "
            "multiple lines."
        )

    if issue_type == "duplicate_invoice":
        return (
            "The invoice number appears more than once in the source data. "
            "Check whether the rows are legitimate line items or duplicate "
            "invoice records before filing or posting the reconciliation."
        )

    return "Review the invoice-level records in both sources."


def reconcile_data(
    gstr1: pd.DataFrame,
    tally: pd.DataFrame,
    tolerance: float = 1.0,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    gstr1 = gstr1.copy()
    tally = tally.copy()

    gstr1_duplicates = find_duplicate_rows(gstr1)
    tally_duplicates = find_duplicate_rows(tally)

    gstr1_by_key = {
        key: group.iloc[0]
        for key, group in gstr1.groupby("invoice_key", sort=False)
        if key
    }
    tally_by_key = {
        key: group.iloc[0]
        for key, group in tally.groupby("invoice_key", sort=False)
        if key
    }

    all_keys = list(dict.fromkeys(
        list(gstr1_by_key.keys()) + list(tally_by_key.keys())
    ))

    rows: list[dict[str, Any]] = []

    for invoice_key in all_keys:
        gstr1_row = gstr1_by_key.get(invoice_key)
        tally_row = tally_by_key.get(invoice_key)

        issue_types: list[str] = []

        if gstr1_row is None:
            issue_types.append("missing_in_gstr1")

        if tally_row is None:
            issue_types.append("missing_in_tally")

        if invoice_key in gstr1_duplicates or invoice_key in tally_duplicates:
            issue_types.append("duplicate_invoice")

        if gstr1_row is not None and tally_row is not None:
            if gstr1_row["gstin"] != tally_row["gstin"]:
                issue_types.append("gstin_mismatch")

            if gstr1_row["invoice_number"] != tally_row["invoice_number"]:
                issue_types.append("invoice_number_mismatch")

            amount_fields = {
                "taxable_value": "taxable_value_mismatch",
                "igst": "igst_mismatch",
                "cgst": "cgst_mismatch",
                "sgst": "sgst_mismatch",
            }

            for field, issue_type in amount_fields.items():
                if not amount_matches(
                    gstr1_row[field],
                    tally_row[field],
                    tolerance,
                ):
                    issue_types.append(issue_type)

        if not issue_types:
            status = "Matched"
        elif any(
            issue in issue_types
            for issue in [
                "missing_in_gstr1",
                "missing_in_tally",
                "duplicate_invoice",
            ]
        ):
            status = "Review"
        else:
            status = "Mismatch"

        taxable_value = (
            gstr1_row["taxable_value"]
            if gstr1_row is not None
            else tally_row["taxable_value"]
        )
        igst = (
            gstr1_row["igst"]
            if gstr1_row is not None
            else tally_row["igst"]
        )
        cgst = (
            gstr1_row["cgst"]
            if gstr1_row is not None
            else tally_row["cgst"]
        )
        sgst = (
            gstr1_row["sgst"]
            if gstr1_row is not None
            else tally_row["sgst"]
        )

        tax_exposure = 0.0

        if "missing_in_gstr1" in issue_types:
            tax_exposure += tally_row["igst"] + tally_row["cgst"] + tally_row["sgst"]

        if "missing_in_tally" in issue_types:
            tax_exposure += gstr1_row["igst"] + gstr1_row["cgst"] + gstr1_row["sgst"]

        for field in ["igst", "cgst", "sgst"]:
            mismatch_name = f"{field}_mismatch"
            if mismatch_name in issue_types:
                tax_exposure += abs(
                    gstr1_row[field] - tally_row[field]
                )

        explanations = [
            explanation_for_issue(
                issue_type,
                gstr1_row,
                tally_row,
                (
                    abs(
                        gstr1_row[issue_type.replace("_mismatch", "")]
                        - tally_row[issue_type.replace("_mismatch", "")]
                    )
                    if gstr1_row is not None
                    and tally_row is not None
                    and issue_type.endswith("_mismatch")
                    and issue_type.replace("_mismatch", "")
                    in {"taxable_value", "igst", "cgst", "sgst"}
                    else 0.0
                ),
            )
            for issue_type in issue_types
        ]

        rows.append(
            {
                "invoice_number": invoice_key,
                "status": status,
                "issues": ", ".join(
                    MISMATCH_LABELS[issue] for issue in issue_types
                ),
                "issue_codes": issue_types,
                "ai_explanation": " ".join(explanations),
                "gstr1_gstin": (
                    gstr1_row["gstin"] if gstr1_row is not None else ""
                ),
                "tally_gstin": (
                    tally_row["gstin"] if tally_row is not None else ""
                ),
                "gstr1_taxable_value": (
                    gstr1_row["taxable_value"]
                    if gstr1_row is not None
                    else np.nan
                ),
                "tally_taxable_value": (
                    tally_row["taxable_value"]
                    if tally_row is not None
                    else np.nan
                ),
                "gstr1_igst": (
                    gstr1_row["igst"] if gstr1_row is not None else np.nan
                ),
                "tally_igst": (
                    tally_row["igst"] if tally_row is not None else np.nan
                ),
                "gstr1_cgst": (
                    gstr1_row["cgst"] if gstr1_row is not None else np.nan
                ),
                "tally_cgst": (
                    tally_row["cgst"] if tally_row is not None else np.nan
                ),
                "gstr1_sgst": (
                    gstr1_row["sgst"] if gstr1_row is not None else np.nan
                ),
                "tally_sgst": (
                    tally_row["sgst"] if tally_row is not None else np.nan
                ),
                "taxable_value": taxable_value,
                "gst_exposure": tax_exposure,
            }
        )

    result = pd.DataFrame(rows)

    total_invoices = len(result)
    matched_invoices = int((result["status"] == "Matched").sum())
    mismatch_invoices = total_invoices - matched_invoices
    gst_exposure = float(result["gst_exposure"].sum())

    if total_invoices:
        health_score = round(
            matched_invoices / total_invoices * 100,
            1,
        )
    else:
        health_score = 0.0

    issue_counts = {}
    for issue_code, label in MISMATCH_LABELS.items():
        issue_counts[label] = int(
            result["issue_codes"].apply(
                lambda codes: issue_code in codes
            ).sum()
        )

    summary = {
        "total_invoices": total_invoices,
        "matched_invoices": matched_invoices,
        "mismatch_invoices": mismatch_invoices,
        "gst_exposure": gst_exposure,
        "health_score": health_score,
        "issue_counts": issue_counts,
    }

    return result, summary


def generate_management_summary(summary: dict[str, Any]) -> str:
    total = summary["total_invoices"]
    matched = summary["matched_invoices"]
    mismatches = summary["mismatch_invoices"]
    exposure = summary["gst_exposure"]
    score = summary["health_score"]

    issue_counts = summary["issue_counts"]
    top_issues = sorted(
        issue_counts.items(),
        key=lambda item: item[1],
        reverse=True,
    )
    top_issues = [
        f"{label} ({count})"
        for label, count in top_issues
        if count > 0
    ][:3]

    issue_text = ", ".join(top_issues) if top_issues else "none"

    return f"""
GST Compliance Copilot Management Summary

The reconciliation covered {total:,} invoice keys. {matched:,} invoices
matched across GSTR-1 and Tally, while {mismatches:,} invoices require review.
The calculated GST exposure is Rs {exposure:,.2f}, and the GST Health Score is
{score:.1f}/100.

The most significant exception categories are: {issue_text}.

Management actions:
1. Prioritize invoices with missing records and duplicate invoice numbers.
2. Validate GSTIN and invoice numbering against the customer master and source
   documents.
3. Review tax-value differences before filing, amendment, or period close.
4. Retain the reconciliation output and supporting source files for audit trail.
""".strip()

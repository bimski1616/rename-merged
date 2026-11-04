
from __future__ import annotations

from collections import Counter
from io import BytesIO
from pathlib import Path
import re
import zipfile

import pandas as pd


# =========================================================
# CONFIGURATION
# =========================================================

REQUIRED_COLUMNS = [
    "PO Number",
    "Site ID",
    "Invoice Proposal",
]


# =========================================================
# NORMALIZATION
# =========================================================

def normalize(value: object) -> str:
    """Normalize IDs for matching."""
    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(value or "").upper(),
    )


def safe_filename_part(value: object) -> str:
    """Remove characters that are invalid or unsafe in filenames."""

    text = str(value or "").strip()

    text = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "-",
        text,
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    ).strip(" .")

    return text


# =========================================================
# SITE ID MATCHER
# =========================================================

class SiteMatcher:
    """
    Prepare Site ID matching data once.

    This avoids normalizing all Site IDs and compiling
    fallback regex patterns for every PDF.
    """

    def __init__(self, site_ids: list[str]):

        self.normalized_sites = []

        for site_id in site_ids:

            normalized = normalize(site_id)

            if not normalized:
                continue

            self.normalized_sites.append({
                "original": site_id,
                "normalized": normalized,
                "pattern": re.compile(
                    rf"(?<![A-Za-z0-9])"
                    rf"{re.escape(str(site_id).strip())}"
                    rf"(?![A-Za-z0-9])",
                    flags=re.IGNORECASE,
                ),
            })

    def find(self, filename: str) -> str | None:

        stem = Path(filename).stem

        segments = {
            normalize(part)
            for part in re.split(
                r"[^A-Za-z0-9]+",
                stem,
            )
        }

        # -------------------------------------------------
        # PRIMARY MATCHING
        # -------------------------------------------------

        matches = [
            item["original"]
            for item in self.normalized_sites
            if item["normalized"] in segments
        ]

        if len(matches) == 1:
            return matches[0]

        if len(matches) > 1:
            return None

        # -------------------------------------------------
        # FALLBACK MATCHING
        # -------------------------------------------------

        fallback_matches = [
            item["original"]
            for item in self.normalized_sites
            if item["pattern"].search(stem)
        ]

        if len(fallback_matches) == 1:
            return fallback_matches[0]

        return None


def find_site_id(
    filename: str,
    site_ids: list[str],
) -> str | None:
    """
    Public matching function.

    Compatible with the existing app.py.
    """

    matcher = SiteMatcher(site_ids)

    return matcher.find(filename)


# =========================================================
# LOAD EXCEL DATABASE
# =========================================================

def load_database(excel_source) -> pd.DataFrame:

    if hasattr(excel_source, "seek"):
        excel_source.seek(0)

    df = pd.read_excel(
        excel_source,
        dtype=str,
        engine="openpyxl",
    ).fillna("")

    missing = [
        col
        for col in REQUIRED_COLUMNS
        if col not in df.columns
    ]

    if missing:

        raise ValueError(
            "Kolom wajib tidak ditemukan: "
            + ", ".join(missing)
        )

    for col in REQUIRED_COLUMNS:

        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
        )

    df["_site_norm"] = df["Site ID"].map(normalize)

    df = df[
        df["_site_norm"] != ""
    ].copy()

    return df


# =========================================================
# VALIDATE DATABASE
# =========================================================

def validate_database(excel_source) -> dict:

    try:

        df = load_database(excel_source)

        duplicated = (
            df[
                df["_site_norm"].duplicated(keep=False)
            ]["Site ID"]
            .drop_duplicates()
            .tolist()
        )

        if duplicated:

            return {
                "ok": False,
                "errors": [
                    "Ada Site ID duplikat di Excel. "
                    "Perbaiki agar satu Site ID hanya memiliki "
                    "satu mapping: "
                    + ", ".join(duplicated[:20])
                ],
            }

        site_ids = df["Site ID"].tolist()

        matcher = SiteMatcher(site_ids)

        return {
            "ok": True,
            "errors": [],
            "preview_match": (
                lambda filename:
                matcher.find(filename) or "Tidak terdeteksi"
            ),
        }

    except Exception as exc:

        return {
            "ok": False,
            "errors": [str(exc)],
        }


# =========================================================
# PROCESS RENAME
# =========================================================

def process_files(excel_source, pdf_files):
    """
    Create a ZIP containing renamed PDFs
    and return a detailed processing log.
    """

    df = load_database(excel_source)

    # -----------------------------------------------------
    # CHECK DUPLICATE SITE ID
    # -----------------------------------------------------

    duplicated = (
        df[
            df["_site_norm"].duplicated(keep=False)
        ]["Site ID"]
        .drop_duplicates()
        .tolist()
    )

    if duplicated:

        raise ValueError(
            "Site ID duplikat ditemukan di Excel: "
            + ", ".join(duplicated[:20])
        )

    # -----------------------------------------------------
    # BUILD LOOKUP
    # -----------------------------------------------------

    records = df.to_dict("records")

    lookup = {
        row["_site_norm"]: {
            "site_id": row["Site ID"],
            "po_number": row["PO Number"],
            "invoice_proposal": row["Invoice Proposal"],
        }
        for row in records
    }

    all_site_ids = [
        item["site_id"]
        for item in lookup.values()
    ]

    matcher = SiteMatcher(all_site_ids)

    # -----------------------------------------------------
    # OUTPUT PREPARATION
    # -----------------------------------------------------

    output_buffer = BytesIO()

    log_rows = []

    used_names = Counter()
    used_output_names = set()

    renamed_count = 0

    # -----------------------------------------------------
    # PROCESS PDF
    # -----------------------------------------------------

    with zipfile.ZipFile(
        output_buffer,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        allowZip64=True,
    ) as archive:

        for pdf in pdf_files:

            original_name = Path(pdf.name).name

            row = {
                "Original Filename": original_name,
                "Detected Site ID": "",
                "Invoice Proposal": "",
                "PO Number": "",
                "New Filename": "",
                "Status": "",
                "Remarks": "",
            }

            # ---------------------------------------------
            # CHECK FILE EXTENSION
            # ---------------------------------------------

            if Path(original_name).suffix.lower() != ".pdf":

                row["Status"] = "SKIPPED"
                row["Remarks"] = "File bukan PDF."

                log_rows.append(row)

                continue

            # ---------------------------------------------
            # MATCH SITE ID
            # ---------------------------------------------

            site_id = matcher.find(original_name)

            if not site_id:

                row["Status"] = "NOT MATCHED"

                row["Remarks"] = (
                    "Site ID tidak ditemukan atau "
                    "lebih dari satu kandidat cocok."
                )

                log_rows.append(row)

                continue

            record = lookup.get(
                normalize(site_id)
            )

            if record is None:

                row["Status"] = "NOT MATCHED"
                row["Remarks"] = (
                    "Mapping Site ID tidak ditemukan."
                )

                log_rows.append(row)

                continue

            row["Detected Site ID"] = site_id

            row["Invoice Proposal"] = (
                record["invoice_proposal"]
            )

            row["PO Number"] = (
                record["po_number"]
            )

            # ---------------------------------------------
            # CHECK REQUIRED MAPPING DATA
            # ---------------------------------------------

            if (
                not record["invoice_proposal"]
                or not record["po_number"]
            ):

                row["Status"] = "MISSING DATA"

                row["Remarks"] = (
                    "Invoice Proposal atau PO Number "
                    "kosong di Excel."
                )

                log_rows.append(row)

                continue

            # ---------------------------------------------
            # GENERATE OUTPUT FILENAME
            # ---------------------------------------------

            invoice = safe_filename_part(
                record["invoice_proposal"]
            )

            po_number = safe_filename_part(
                record["po_number"]
            )

            base_name = f"{invoice}_{po_number}"

            candidate = f"{base_name}.pdf"

            used_names[base_name.casefold()] += 1

            sequence = used_names[
                base_name.casefold()
            ]

            if sequence > 1:

                candidate = (
                    f"{base_name} ({sequence}).pdf"
                )

            # Prevent filename collisions
            while candidate.casefold() in used_output_names:

                sequence += 1

                candidate = (
                    f"{base_name} ({sequence}).pdf"
                )

            used_output_names.add(
                candidate.casefold()
            )

            # ---------------------------------------------
            # WRITE PDF INTO ZIP
            # ---------------------------------------------

            pdf.seek(0)

            pdf_bytes = pdf.read()

            archive.writestr(
                candidate,
                pdf_bytes,
            )

            row["New Filename"] = candidate

            row["Status"] = "RENAMED"

            row["Remarks"] = "Berhasil diproses."

            log_rows.append(row)

            renamed_count += 1

    # -----------------------------------------------------
    # SUMMARY
    # -----------------------------------------------------

    log_df = pd.DataFrame(log_rows)

    summary = {
        "total": len(pdf_files),
        "renamed": renamed_count,
        "issues": len(pdf_files) - renamed_count,
    }

    output_buffer.seek(0)

    return (
        output_buffer.getvalue(),
        log_df,
        summary,
    )

from __future__ import annotations

from collections import Counter
from io import BytesIO
from pathlib import Path
import re
import zipfile

import pandas as pd

REQUIRED_COLUMNS = ["PO Number", "Site ID", "Invoice Proposal"]


def normalize(value: object) -> str:
    """Normalize IDs for matching: uppercase and remove non-alphanumeric characters."""
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def safe_filename_part(value: object) -> str:
    """Remove characters that are invalid or unsafe in common file systems."""
    text = str(value or "").strip()
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text


def find_site_id(filename: str, site_ids: list[str]) -> str | None:
    """
    Detect Site ID from a filename.

    Primary method: compare normalized filename segments split by punctuation
    (for example WCC_NPO_INDOSAT_02KBJ0023 -> 02KBJ0023).
    Fallback: search for a complete Site ID bounded by non-alphanumeric chars.
    """
    stem = Path(filename).stem
    segments = [normalize(part) for part in re.split(r"[^A-Za-z0-9]+", stem)]
    normalized_sites = [(site_id, normalize(site_id)) for site_id in site_ids]

    matches = []
    for original, normalized_site in normalized_sites:
        if not normalized_site:
            continue
        if normalized_site in segments:
            matches.append(original)

    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        return None

    # Fallback for IDs that are surrounded by separators but not isolated by underscores.
    fallback_matches = []
    for original, _ in normalized_sites:
        escaped = re.escape(str(original).strip())
        pattern = rf"(?<![A-Za-z0-9]){escaped}(?![A-Za-z0-9])"
        if re.search(pattern, stem, flags=re.IGNORECASE):
            fallback_matches.append(original)

    if len(fallback_matches) == 1:
        return fallback_matches[0]
    return None


def load_database(excel_source) -> pd.DataFrame:
    if hasattr(excel_source, "seek"):
        excel_source.seek(0)
    df = pd.read_excel(excel_source, dtype=str, engine="openpyxl").fillna("")
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError("Kolom wajib tidak ditemukan: " + ", ".join(missing))

    for col in REQUIRED_COLUMNS:
        df[col] = df[col].astype(str).str.strip()

    df["_site_norm"] = df["Site ID"].map(normalize)
    df = df[df["_site_norm"] != ""].copy()
    return df


def validate_database(excel_source) -> dict:
    try:
        df = load_database(excel_source)
        duplicated = df[df["_site_norm"].duplicated(keep=False)]["Site ID"].drop_duplicates().tolist()
        if duplicated:
            return {
                "ok": False,
                "errors": [
                    "Ada Site ID duplikat di Excel. Perbaiki agar satu Site ID hanya memiliki satu mapping: "
                    + ", ".join(duplicated[:20])
                ],
            }

        site_ids = df["Site ID"].tolist()
        return {
            "ok": True,
            "errors": [],
            "preview_match": lambda filename: find_site_id(filename, site_ids) or "Tidak terdeteksi",
        }
    except Exception as exc:
        return {"ok": False, "errors": [str(exc)]}


def process_files(excel_source, pdf_files):
    """Create a ZIP containing renamed PDFs and return a detailed processing log."""
    df = load_database(excel_source)

    duplicated = df[df["_site_norm"].duplicated(keep=False)]["Site ID"].drop_duplicates().tolist()
    if duplicated:
        raise ValueError(
            "Site ID duplikat ditemukan di Excel: " + ", ".join(duplicated[:20])
        )

    lookup = {
        row["_site_norm"]: {
            "site_id": row["Site ID"],
            "po_number": row["PO Number"],
            "invoice_proposal": row["Invoice Proposal"],
        }
        for _, row in df.iterrows()
    }
    all_site_ids = [item["site_id"] for item in lookup.values()]

    output_buffer = BytesIO()
    log_rows = []
    used_names = Counter()
    used_output_names = set()
    renamed_count = 0

    with zipfile.ZipFile(output_buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
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

            if Path(original_name).suffix.lower() != ".pdf":
                row["Status"] = "SKIPPED"
                row["Remarks"] = "File bukan PDF."
                log_rows.append(row)
                continue

            site_id = find_site_id(original_name, all_site_ids)
            if not site_id:
                row["Status"] = "NOT MATCHED"
                row["Remarks"] = "Site ID tidak ditemukan atau lebih dari satu kandidat cocok."
                log_rows.append(row)
                continue

            record = lookup.get(normalize(site_id))
            row["Detected Site ID"] = site_id
            row["Invoice Proposal"] = record["invoice_proposal"]
            row["PO Number"] = record["po_number"]

            if not record["invoice_proposal"] or not record["po_number"]:
                row["Status"] = "MISSING DATA"
                row["Remarks"] = "Invoice Proposal atau PO Number kosong di Excel."
                log_rows.append(row)
                continue

            invoice = safe_filename_part(record["invoice_proposal"])
            po_number = safe_filename_part(record["po_number"])
            base_name = f"{invoice}_{po_number}"
            candidate = f"{base_name}.pdf"

            used_names[base_name.casefold()] += 1
            sequence = used_names[base_name.casefold()]
            if sequence > 1:
                candidate = f"{base_name} ({sequence}).pdf"

            # Guard against collisions, including names that already contain a suffix.
            while candidate.casefold() in used_output_names:
                sequence += 1
                candidate = f"{base_name} ({sequence}).pdf"
            used_output_names.add(candidate.casefold())

            pdf.seek(0)
            pdf_bytes = pdf.read()
            archive.writestr(candidate, pdf_bytes)

            row["New Filename"] = candidate
            row["Status"] = "RENAMED"
            row["Remarks"] = "Berhasil diproses."
            log_rows.append(row)
            renamed_count += 1

    log_df = pd.DataFrame(log_rows)
    summary = {
        "total": len(pdf_files),
        "renamed": renamed_count,
        "issues": len(pdf_files) - renamed_count,
    }
    output_buffer.seek(0)
    return output_buffer.getvalue(), log_df, summary

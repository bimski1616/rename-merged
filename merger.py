from io import BytesIO
from pathlib import Path
from collections import defaultdict
import re
import zipfile

import pandas as pd
from pypdf import PdfReader, PdfWriter


def normalize(value):
    """Normalize PO Number for matching."""
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def unique_filename(filename, used_names):
    """Prevent duplicate filenames in output ZIP."""
    path = Path(filename)
    base = path.stem
    extension = path.suffix

    candidate = filename
    counter = 2

    while candidate.casefold() in used_names:
        candidate = f"{base} ({counter}){extension}"
        counter += 1

    used_names.add(candidate.casefold())
    return candidate


def read_renamed_pdfs(renamed_zip=None, renamed_pdf_files=None):
    """Read renamed invoice PDFs from ZIP and/or individual uploads."""
    pdf_items = []

    if renamed_zip is not None:
        renamed_zip.seek(0)

        with zipfile.ZipFile(renamed_zip, "r") as archive:
            for name in archive.namelist():
                if name.lower().endswith(".pdf") and not name.endswith("/"):
                    pdf_items.append({
                        "filename": Path(name).name,
                        "content": archive.read(name),
                    })

    for pdf in renamed_pdf_files or []:
        pdf.seek(0)
        pdf_items.append({
            "filename": Path(pdf.name).name,
            "content": pdf.read(),
        })

    return pdf_items


def read_po_files(po_files):
    """Build PO Number -> list of PDF files mapping."""
    po_lookup = defaultdict(list)

    for pdf in po_files:
        filename = Path(pdf.name).name

        if Path(filename).suffix.lower() != ".pdf":
            continue

        po_number = Path(filename).stem
        po_key = normalize(po_number)

        if not po_key:
            continue

        pdf.seek(0)
        po_lookup[po_key].append({
            "filename": filename,
            "content": pdf.read(),
        })

    return po_lookup


def extract_po_from_invoice(filename):
    """
    Expected renamed invoice format:
    InvoiceProposal_PO Number.pdf

    PO Number is taken from the last underscore-separated segment.
    """
    stem = Path(filename).stem

    if "_" not in stem:
        return ""

    return stem.rsplit("_", 1)[-1].strip()


def merge_pdf_bytes(invoice_bytes, po_bytes):
    """Merge invoice first, followed by PO."""
    invoice_stream = BytesIO(invoice_bytes)
    po_stream = BytesIO(po_bytes)

    invoice_reader = PdfReader(invoice_stream, strict=False)
    po_reader = PdfReader(po_stream, strict=False)

    if invoice_reader.is_encrypted or po_reader.is_encrypted:
        raise ValueError("PDF terenkripsi/password protected.")

    writer = PdfWriter()

    for page in invoice_reader.pages:
        writer.add_page(page)

    for page in po_reader.pages:
        writer.add_page(page)

    output = BytesIO()
    writer.write(output)

    return output.getvalue()


def process_merge(renamed_zip, renamed_pdf_files, po_files):
    """
    Process renamed invoices against PO PDFs.

    Missing or duplicate PO:
    Keep invoice unchanged and add remarks to the log.
    """
    invoices = read_renamed_pdfs(
        renamed_zip=renamed_zip,
        renamed_pdf_files=renamed_pdf_files,
    )

    po_lookup = read_po_files(po_files)

    if not invoices:
        raise ValueError("Tidak ada PDF hasil rename yang ditemukan.")

    if not po_lookup:
        raise ValueError("Tidak ada PDF PO yang valid ditemukan.")

    output_buffer = BytesIO()
    log_rows = []
    used_names = set()

    merged_count = 0
    missing_count = 0
    duplicate_count = 0
    error_count = 0

    with zipfile.ZipFile(
        output_buffer,
        "w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archive:

        for invoice in invoices:
            filename = invoice["filename"]
            invoice_bytes = invoice["content"]

            po_number = extract_po_from_invoice(filename)
            po_key = normalize(po_number)

            row = {
                "Invoice Filename": filename,
                "PO Number": po_number,
                "Matched PO Filename": "",
                "Final Filename": "",
                "Status": "",
                "Remarks": "",
            }

            final_bytes = invoice_bytes

            if not po_key:
                row["Status"] = "NOT MATCHED"
                row["Remarks"] = (
                    "PO Number tidak dapat dibaca dari nama file invoice."
                )
                error_count += 1

            elif po_key not in po_lookup:
                row["Status"] = "PO NOT FOUND"
                row["Remarks"] = (
                    "PDF PO tidak ditemukan. Invoice tetap disimpan tanpa merge."
                )
                missing_count += 1

            elif len(po_lookup[po_key]) > 1:
                duplicate_files = [
                    item["filename"] for item in po_lookup[po_key]
                ]

                row["Status"] = "DUPLICATE PO"
                row["Remarks"] = (
                    "Ditemukan lebih dari satu PDF untuk PO Number ini: "
                    + ", ".join(duplicate_files)
                    + ". Invoice disimpan tanpa merge."
                )
                duplicate_count += 1

            else:
                po_item = po_lookup[po_key][0]
                row["Matched PO Filename"] = po_item["filename"]

                try:
                    final_bytes = merge_pdf_bytes(
                        invoice_bytes,
                        po_item["content"],
                    )

                    row["Status"] = "SUCCESS - MERGED"
                    row["Remarks"] = (
                        "Invoice dan PO berhasil digabung. "
                        "Urutan: Invoice lalu PO."
                    )
                    merged_count += 1

                except Exception as exc:
                    final_bytes = invoice_bytes
                    row["Status"] = "PROCESSING ERROR"
                    row["Remarks"] = (
                        "Merge gagal. Invoice tetap disimpan tanpa merge. "
                        f"Detail: {str(exc)}"
                    )
                    error_count += 1

            output_name = unique_filename(filename, used_names)

            archive.writestr(output_name, final_bytes)

            row["Final Filename"] = output_name
            log_rows.append(row)

    log_df = pd.DataFrame(log_rows)

    summary = {
        "total": len(invoices),
        "merged": merged_count,
        "missing": missing_count,
        "duplicate": duplicate_count,
        "errors": error_count,
    }

    output_buffer.seek(0)

    return output_buffer.getvalue(), log_df, summary


def create_excel_log(log_df):
    """Create an Excel processing log."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        log_df.to_excel(
            writer,
            index=False,
            sheet_name="Merge Log",
        )

        worksheet = writer.sheets["Merge Log"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for column_cells in worksheet.columns:
            max_length = max(
                len(str(cell.value or "")) for cell in column_cells
            )
            worksheet.column_dimensions[
                column_cells[0].column_letter
            ].width = min(max_length + 3, 60)

    output.seek(0)
    return output.getvalue()
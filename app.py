import streamlit as st
import pandas as pd

from renamer import (
    REQUIRED_COLUMNS,
    validate_database,
    process_files,
)

from merger import (
    process_merge,
    create_excel_log,
)


st.set_page_config(
    page_title="PDF Renamer & Merger",
    page_icon="📄",
    layout="wide",
)


st.title("📄 PDF Renamer & Merger")
st.caption(
    "Rename invoice PDFs using Excel references, "
    "then merge them with their corresponding PO PDFs."
)


with st.sidebar:
    st.header("Workflow")
    st.markdown(
        """
        **Step 1 — Rename PDF**
        1. Upload Excel database.
        2. Upload invoice PDFs.
        3. Review matching.
        4. Download renamed PDFs.

        **Step 2 — Merge PDF**
        1. Upload ZIP hasil rename atau PDF hasil rename.
        2. Upload PDF PO.
        3. Match berdasarkan PO Number.
        4. Download final PDFs and processing log.
        """
    )

    st.info(
        "Original files are not modified. "
        "Each process generates a separate ZIP output."
    )


tab_rename, tab_merge = st.tabs([
    "1️⃣ Rename PDF",
    "2️⃣ Merge PDF",
])


# =========================================================
# TAB 1 — RENAME PDF
# =========================================================

with tab_rename:

    st.header("Step 1 — Rename Invoice PDFs")

    st.subheader("1. Upload Excel database")

    excel_file = st.file_uploader(
        "Choose Excel reference file (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
        key="excel_upload",
    )

    st.subheader("2. Upload invoice PDF files")

    pdf_files = st.file_uploader(
        "Choose invoice PDF files",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_upload",
    )

    if excel_file is not None:

        try:
            database = pd.read_excel(
                excel_file,
                dtype=str,
                engine="openpyxl",
            ).fillna("")

            missing = [
                col for col in REQUIRED_COLUMNS
                if col not in database.columns
            ]

            if missing:
                st.error(
                    "Kolom wajib tidak ditemukan: "
                    + ", ".join(missing)
                )

            else:
                st.success(
                    f"Database berhasil dibaca: {len(database):,} baris."
                )

                with st.expander("Preview database"):
                    st.dataframe(
                        database[REQUIRED_COLUMNS].head(20),
                        use_container_width=True,
                    )

        except Exception as exc:
            st.error(f"Gagal membaca Excel: {exc}")

    if excel_file is not None and pdf_files:

        try:
            excel_file.seek(0)
            validation = validate_database(excel_file)

            if validation["ok"]:

                preview_rows = []

                for pdf in pdf_files:
                    preview_rows.append({
                        "Original Filename": pdf.name,
                        "Detected Site ID": validation[
                            "preview_match"
                        ](pdf.name),
                    })

                preview_df = pd.DataFrame(preview_rows)

                st.subheader("3. File preview")

                st.dataframe(
                    preview_df,
                    use_container_width=True,
                    hide_index=True,
                )

                st.caption(
                    "Output format: Invoice Proposal_PO Number.pdf"
                )

                if st.button(
                    "🚀 Process & Rename PDFs",
                    type="primary",
                    use_container_width=True,
                    key="rename_button",
                ):

                    with st.spinner("Renaming PDF files..."):

                        excel_file.seek(0)

                        zip_bytes, log_df, summary = process_files(
                            excel_file,
                            pdf_files,
                        )

                        st.session_state["renamer_zip"] = zip_bytes
                        st.session_state["renamer_log"] = log_df
                        st.session_state["renamer_summary"] = summary

                if "renamer_summary" in st.session_state:

                    summary = st.session_state["renamer_summary"]

                    st.subheader("4. Rename result")

                    col1, col2, col3 = st.columns(3)

                    col1.metric("Total PDF", summary["total"])
                    col2.metric(
                        "Successfully renamed",
                        summary["renamed"],
                    )
                    col3.metric(
                        "Need attention",
                        summary["issues"],
                    )

                    st.dataframe(
                        st.session_state["renamer_log"],
                        use_container_width=True,
                        hide_index=True,
                    )

                    st.download_button(
                        "⬇️ Download renamed PDFs (ZIP)",
                        data=st.session_state["renamer_zip"],
                        file_name="renamed_pdfs.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True,
                        key="download_rename_zip",
                    )

                    rename_log_bytes = (
                        st.session_state["renamer_log"]
                        .to_csv(index=False)
                        .encode("utf-8-sig")
                    )

                    st.download_button(
                        "⬇️ Download rename log (CSV)",
                        data=rename_log_bytes,
                        file_name="rename_log.csv",
                        mime="text/csv",
                        key="download_rename_log",
                    )

            else:
                for error in validation["errors"]:
                    st.error(error)

        except Exception as exc:
            st.error(f"Rename validation failed: {exc}")


# =========================================================
# TAB 2 — MERGE PDF
# =========================================================

with tab_merge:

    st.header("Step 2 — Merge Invoice with PO")

    st.info(
        "Upload the ZIP generated from Step 1, "
        "or upload the extracted renamed PDFs individually."
    )

    st.subheader("1. Upload renamed invoice files")

    renamed_zip = st.file_uploader(
        "Upload renamed_pdfs.zip (optional)",
        type=["zip"],
        accept_multiple_files=False,
        key="renamed_zip_upload",
    )

    renamed_pdf_files = st.file_uploader(
        "Or upload renamed invoice PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key="renamed_pdf_upload",
    )

    st.subheader("2. Upload PO PDF files")

    po_files = st.file_uploader(
        "Choose PO PDF files (example: 51938355.pdf)",
        type=["pdf"],
        accept_multiple_files=True,
        key="po_pdf_upload",
    )

    has_invoices = (
        renamed_zip is not None
        or bool(renamed_pdf_files)
    )

    if has_invoices and po_files:

        st.success(
            f"Invoice input ready. "
            f"PO PDF uploaded: {len(po_files):,} files."
        )

        st.caption(
            "Matching rule: "
            "PJIP1126000573_51938355.pdf "
            "will be matched with 51938355.pdf."
        )

        if st.button(
            "🔗 Process & Merge PDFs",
            type="primary",
            use_container_width=True,
            key="merge_button",
        ):

            try:
                with st.spinner(
                    "Matching PO files and merging PDF documents..."
                ):

                    output_zip, merge_log, summary = process_merge(
                        renamed_zip,
                        renamed_pdf_files,
                        po_files,
                    )

                    log_excel = create_excel_log(merge_log)

                    st.session_state["merge_zip"] = output_zip
                    st.session_state["merge_log"] = merge_log
                    st.session_state["merge_summary"] = summary
                    st.session_state["merge_log_excel"] = log_excel

            except Exception as exc:
                st.error(f"Merge process failed: {exc}")

    elif not has_invoices:
        st.warning(
            "Upload ZIP hasil rename atau PDF hasil rename terlebih dahulu."
        )

    elif not po_files:
        st.warning("Upload kumpulan PDF PO terlebih dahulu.")

    if "merge_summary" in st.session_state:

        summary = st.session_state["merge_summary"]

        st.divider()
        st.subheader("3. Merge processing result")

        col1, col2, col3, col4, col5 = st.columns(5)

        col1.metric("Total Invoice", summary["total"])
        col2.metric("Successfully Merged", summary["merged"])
        col3.metric("PO Not Found", summary["missing"])
        col4.metric("Duplicate PO", summary["duplicate"])
        col5.metric("Errors", summary["errors"])

        st.subheader("Processing log")

        st.dataframe(
            st.session_state["merge_log"],
            use_container_width=True,
            hide_index=True,
        )

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                "⬇️ Download Final PDFs (ZIP)",
                data=st.session_state["merge_zip"],
                file_name="final_merged_pdfs.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key="download_merge_zip",
            )

        with col2:
            st.download_button(
                "⬇️ Download Merge Log (Excel)",
                data=st.session_state["merge_log_excel"],
                file_name="merge_processing_log.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
                key="download_merge_log",
            )


st.divider()

st.caption(
    "PDF Renamer & Merger • "
    "Original files remain unchanged. "
    "Rename and Merge are independent processes."
)
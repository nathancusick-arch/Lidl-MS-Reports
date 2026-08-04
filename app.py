from __future__ import annotations

import hashlib
from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st

from report_generator import ReportGenerationError, generate_reports_from_csv


REPORT_TIMEZONE = ZoneInfo("Europe/London")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


st.set_page_config(
    page_title="Lidl MS Monthly Report Generator",
    page_icon="📊",
    layout="wide",
)

st.title("Lidl MS Monthly Report Generator")
current_period = datetime.now(REPORT_TIMEZONE).strftime("%B %Y")
st.write(
    "Upload the monthly `audits_basic_data_export.csv`. "
    "The app will split MCO and SCO, extract aborts and zero charge discrepancies, "
    "and create the standard and password-encrypted MD workbooks."
)
st.caption(f"Output filenames will use the current month: **{current_period}**.")


def get_md_password() -> str | None:
    try:
        password = str(st.secrets["MD_PASSWORD"])
    except KeyError:
        return None
    return password if password else None


def mime_type(filename: str) -> str:
    return "text/csv" if filename.lower().endswith(".csv") else XLSX_MIME


uploaded_file = st.file_uploader(
    "Upload audits_basic_data_export.csv",
    type=["csv"],
    help="One export containing both Till Compliance and Till Compliance - SCO rows.",
)

md_password = get_md_password()
if md_password is None:
    st.error(
        "The Streamlit secret `MD_PASSWORD` has not been configured. "
        "Add it in the app's Streamlit Cloud settings before generating reports."
    )

generate_clicked = st.button(
    "Generate reports",
    type="primary",
    disabled=uploaded_file is None or md_password is None,
)

if generate_clicked and uploaded_file is not None and md_password is not None:
    csv_bytes = uploaded_file.getvalue()
    with st.spinner("Generating and encrypting seven reports…"):
        try:
            result = generate_reports_from_csv(csv_bytes, md_password)
        except ReportGenerationError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"The reports could not be generated: {exc}")
        else:
            cache_key = hashlib.sha256(csv_bytes).hexdigest()
            st.session_state["lidl_reports"] = result
            st.session_state["lidl_reports_source"] = cache_key

if uploaded_file is not None and "lidl_reports" in st.session_state:
    current_hash = hashlib.sha256(uploaded_file.getvalue()).hexdigest()
    if st.session_state.get("lidl_reports_source") != current_hash:
        st.info("The uploaded file has changed. Select **Generate reports** to refresh the downloads.")
    else:
        result = st.session_state["lidl_reports"]
        st.success(f"Seven {result.reporting_label} reports generated successfully.")

        stats = result.stats
        first_metrics = st.columns(4)
        first_metrics[0].metric("Uploaded rows", f"{stats['uploaded_rows']:,}")
        first_metrics[1].metric("MCO report", f"{stats['mco_report_rows']:,}")
        first_metrics[2].metric("MCO aborts", f"{stats['mco_abort_rows']:,}")
        first_metrics[3].metric("Zero discrepancies", f"{stats['zero_discrepancy_rows']:,}")

        second_metrics = st.columns(3)
        second_metrics[0].metric("SCO report", f"{stats['sco_report_rows']:,}")
        second_metrics[1].metric("SCO aborts", f"{stats['sco_abort_rows']:,}")
        second_metrics[2].metric("Ignored rows", f"{stats['ignored_rows']:,}")

        zip_name = f"Lidl MS Reports - {result.reporting_label}.zip"
        st.download_button(
            "Download all reports (.zip)",
            data=result.as_zip(),
            file_name=zip_name,
            mime="application/zip",
            type="primary",
            width="stretch",
            on_click="ignore",
        )

        with st.expander("Download individual reports"):
            download_columns = st.columns(2)
            for index, (filename, contents) in enumerate(result.files.items()):
                with download_columns[index % 2]:
                    st.download_button(
                        filename,
                        data=contents,
                        file_name=filename,
                        mime=mime_type(filename),
                        key=f"download_{index}_{filename}",
                        width="stretch",
                        on_click="ignore",
                    )

        if result.warnings:
            with st.expander(f"Mapping warnings ({len(result.warnings)})"):
                for warning in result.warnings:
                    st.warning(warning)

import streamlit as st
from converter_vas import xml_to_excel_vas, excel_to_xml_vas, read_excel_bytes_vas

st.set_page_config(page_title="VAS (Vastrm) Order Converter", layout="wide")

with st.sidebar:
    st.title("VAS Order Converter")
    st.caption("Vastrm order feed (Red Rooster VAS)")
    mode = st.radio(
        "Chọn chức năng",
        ["XML → Excel", "Excel → XML"],
        key="vas_mode",
    )

st.title("VAS (Vastrm) Order Converter")

if mode == "XML → Excel":
    st.subheader("Convert VAS XML to Excel")
    file = st.file_uploader("Upload a .xml file", type=["xml"], key="vas_xml_upload")
    if file:
        raw = file.read()
        try:
            excel_bytes = xml_to_excel_vas(raw)
            st.download_button(
                label="⬇️ Download Excel",
                data=excel_bytes,
                file_name="vas_orders.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            preview_rows = read_excel_bytes_vas(excel_bytes)
            st.success(f"Converted {len(preview_rows)} order detail row(s).")
            if preview_rows:
                import pandas as pd
                st.dataframe(pd.DataFrame(preview_rows[:5]), use_container_width=True)
                st.caption(f"Showing up to 5 of {len(preview_rows)} rows")

        except Exception as e:
            st.error(f"Error processing file: {e}")

else:
    st.subheader("Convert Excel to VAS XML")
    file = st.file_uploader("Upload a .xlsx file", type=["xlsx"], key="vas_excel_upload")
    if file:
        raw = file.read()
        try:
            xml_str, total = excel_to_xml_vas(raw)
            xml_bytes = xml_str.encode("utf-8")

            st.download_button(
                label="⬇️ Download XML",
                data=xml_bytes,
                file_name="vas_orders.xml",
                mime="application/xml",
            )

            st.success(f"Converted {total} order detail row(s) successfully.")

            with st.expander("Preview XML (first 100 lines)"):
                preview_lines = xml_str.splitlines()[:100]
                st.code("\n".join(preview_lines), language="xml")

        except Exception as e:
            st.error(f"Error processing file: {e}")

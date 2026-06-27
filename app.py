import json
import streamlit as st
from converter import json_to_excel, excel_to_json
from excel_to_xml_jhil import excel_to_xml

st.set_page_config(page_title="TAL Order Converter", layout="wide")
st.title("TAL Order Converter")

tab1, tab2, tab3 = st.tabs(["JSON → Excel", "Excel → JSON", "Excel → XML (JHIL Knit)"])

with tab1:
    st.subheader("Convert JSON to Excel")
    file = st.file_uploader("Upload a .json file", type=["json"], key="json_upload")
    if file:
        raw = file.read()
        try:
            data = json.loads(raw)
            orders = data.get("Order", [])

            excel_bytes = json_to_excel(raw)
            st.download_button(
                label="⬇️ Download Excel",
                data=excel_bytes,
                file_name="orders.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            import pandas as pd
            preview_rows = []
            for order in orders[:5]:
                src = order.get("OrderHeader", order)
                header = {k: src.get(k, "") for k in ["OrderNumber", "ReferenceNo", "CustomerName", "OrderDate", "OrderType"]}
                header["LineCount"] = len(order.get("OrderLine", []))
                preview_rows.append(header)

            if preview_rows:
                st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)
                st.caption(f"Showing up to 5 of {len(orders)} orders")

        except Exception as e:
            st.error(f"Error processing file: {e}")

with tab2:
    st.subheader("Convert Excel to JSON")
    file = st.file_uploader("Upload a .xlsx file", type=["xlsx"], key="excel_upload")
    if file:
        raw = file.read()
        try:
            json_bytes = excel_to_json(raw)
            data = json.loads(json_bytes)
            orders = data.get("Order", [])

            st.download_button(
                label="⬇️ Download JSON",
                data=json_bytes,
                file_name="orders.json",
                mime="application/json",
            )

            st.json(orders[:5])
            st.caption(f"Showing up to 5 of {len(orders)} orders")

        except Exception as e:
            st.error(f"Error processing file: {e}")

with tab3:
    st.subheader("Convert Excel to XML (JHIL Knit PO)")
    file = st.file_uploader("Upload a .xlsx file", type=["xlsx"], key="xml_upload")
    if file:
        raw = file.read()
        try:
            xml_str, total, errors = excel_to_xml(raw)
            xml_bytes = xml_str.encode("utf-8")

            st.download_button(
                label="⬇️ Download XML",
                data=xml_bytes,
                file_name="knit_orders.xml",
                mime="application/xml",
            )

            st.success(f"Converted {total} records successfully.")

            if errors:
                st.warning(f"{len(errors)} record(s) had errors:")
                for err in errors:
                    st.text(err)

            with st.expander("Preview XML (first 100 lines)"):
                preview_lines = xml_str.splitlines()[:100]
                st.code("\n".join(preview_lines), language="xml")

        except Exception as e:
            st.error(f"Error processing file: {e}")

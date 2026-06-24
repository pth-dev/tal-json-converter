import json
import streamlit as st
from converter import json_to_excel, excel_to_json

st.set_page_config(page_title="TAL JSON ↔ Excel Converter", layout="wide")
st.title("TAL JSON ↔ Excel Converter")

tab1, tab2 = st.tabs(["JSON → Excel", "Excel → JSON"])

with tab1:
    st.subheader("Convert JSON to Excel")
    file = st.file_uploader("Upload a .json file", type=["json"], key="json_upload")
    if file:
        raw = file.read()
        try:
            data = json.loads(raw)
            orders = data.get("Order", [])

            # Build preview rows
            preview_rows = []
            for order in orders[:5]:
                header = {k: order.get(k, "") for k in ["OrderNumber", "ReferenceNo", "CustomerName", "OrderDate", "OrderType"]}
                lines = order.get("OrderLine", [])
                header["LineCount"] = len(lines)
                preview_rows.append(header)

            if preview_rows:
                import pandas as pd
                st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)
                st.caption(f"Showing up to 5 of {len(orders)} orders")

            excel_bytes = json_to_excel(raw)
            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name="orders.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
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

            st.json(orders[:5])
            st.caption(f"Showing up to 5 of {len(orders)} orders")

            st.download_button(
                label="Download JSON",
                data=json_bytes,
                file_name="orders.json",
                mime="application/json",
            )
        except Exception as e:
            st.error(f"Error processing file: {e}")

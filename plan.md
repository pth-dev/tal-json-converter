# MVP Plan — TAL JSON ↔ Excel Converter

## Stack
- Python 3.11+, Streamlit, openpyxl, pandas
- Deploy: Streamlit Cloud (free)

---

## File structure
```
tal-converter/
├── app.py            # UI only
├── converter.py      # logic only
└── requirements.txt
```

`requirements.txt`:
```
streamlit
openpyxl
pandas
```

---

## Phase 1 — converter.py

### 1A. json_to_excel(json_bytes) → bytes

```
1. Parse JSON → data["Order"]
2. Loop orders:
   - Lấy OrderHeader fields (flat)
   - Pivot OrderHeaderDetail [{ref,val}] → dict
   - Loop OrderLine → mỗi line 1 row:
       - Pivot OrderLineDetail [{ref,val}] → dict
       - Merge header + headerDetail + lineFields + lineDetail → 1 row
3. openpyxl: write rows, header style, freeze A2, auto_filter
4. Return BytesIO
```

### 1B. excel_to_json(excel_bytes) → bytes

```
1. pandas read_excel → DataFrame
2. Group by OrderNumber
3. Mỗi group → 1 Order dict:
   - HEADER_COLS → OrderHeader object
   - HEADER_DETAIL_COLS (Courier, CourierService...) → OrderHeaderDetail [{ref,val}]
   - Mỗi row trong group → OrderLine:
       - LINE_COLS (Product, LineRefNo...) → OrderLine fields
       - LINE_DETAIL_COLS (Style, MainFabric...) → OrderLineDetail [{ref,val}]
4. Wrap: {"Order": [...]}
5. Return json.dumps bytes
```

**Column mapping constants** (define ở đầu converter.py):
```python
HEADER_COLS = ["OrderNumber","ReferenceNo","RSVS","OrderDate","OrderType",
               "ShipmentMethod","PackagingMethod","CustomerName",
               "ShiptoCustomerName","CustomerAddress1","CustomerAddress2",
               "CustomerAddressCity","CustomerAddressState","CustomerAddressZip",
               "CustomerAddressCountry","CustomerPhoneNumber","SalesPerson"]

HEADER_DETAIL_COLS = ["Courier","CourierService","ShipFromName",
                      "ShipFromAdd1","ShipFromAdd2","ShipFromAdd3",
                      "DCName","DCNumber","StoreNumber","StoreName"]

LINE_COLS = ["Product","LineRefNo","UnitRetailPrice","Tier","Quantity"]

# LINE_DETAIL_COLS = tất cả cột còn lại
```

---

## Phase 2 — app.py

```
st.title("TAL JSON ↔ Excel Converter")
tab1, tab2 = st.tabs(["JSON → Excel", "Excel → JSON"])

Tab 1:
  file = st.file_uploader(".json")
  if file:
    preview 5 rows st.dataframe
    st.download_button → json_to_excel(file)

Tab 2:
  file = st.file_uploader(".xlsx")
  if file:
    preview st.json (5 orders)
    st.download_button → excel_to_json(file)
```

---

## Phase 3 — Deploy

```
1. git init → push lên GitHub (public)
2. share.streamlit.io → New app → connect repo → app.py
3. Done — URL: https://<name>.streamlit.app
```

---

## Thứ tự build

| Bước | Việc | Thời gian |
|------|------|-----------|
| 1 | converter.py — hàm json_to_excel | 30 phút |
| 2 | converter.py — hàm excel_to_json | 1 giờ |
| 3 | Test 2 hàm bằng file thật | 30 phút |
| 4 | app.py — UI 2 tab | 30 phút |
| 5 | Deploy Streamlit Cloud | 15 phút |
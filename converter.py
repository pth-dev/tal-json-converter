import json
import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

HEADER_COLS = [
    "OrderNumber", "ReferenceNo", "RSVS", "OrderDate", "OrderType",
    "ShipmentMethod", "PackagingMethod", "CustomerName",
    "ShiptoCustomerName", "CustomerAddress1", "CustomerAddress2", "CustomerAddress3",
    "CustomerAddressCity", "CustomerAddressState", "CustomerAddressZip",
    "CustomerAddressCountry", "CustomerPhoneNumber", "SalesPerson",
]

HEADER_DETAIL_COLS = [
    "Courier", "CourierService", "ShipFromName",
    "ShipFromAdd1", "ShipFromAdd2", "ShipFromAdd3",
    "DCName", "DCNumber", "StoreNumber", "StoreName",
]

LINE_COLS = ["Product", "LineRefNo", "UnitRetailPrice", "Tier", "Quantity"]


def _pivot_detail(detail_list):
    """Convert [{ref, val}] list to a flat dict."""
    if not detail_list:
        return {}
    return {item.get("ref", ""): item.get("val", "") for item in detail_list}


def json_to_excel(json_bytes: bytes) -> bytes:
    data = json.loads(json_bytes)
    orders = data.get("Order", [])

    rows = []
    for order in orders:
        order_header = order.get("OrderHeader", order)  # fallback to order itself for backwards compat
        header = {col: order_header.get(col, "") for col in HEADER_COLS}
        header_detail = _pivot_detail(order_header.get("OrderHeaderDetail", []))

        for line in order.get("OrderLine", []):
            line_fields = {col: line.get(col, "") for col in LINE_COLS}
            line_detail = _pivot_detail(line.get("OrderLineDetail", []))
            row = {**header, **header_detail, **line_fields, **line_detail}
            rows.append(row)

    if not rows:
        rows = [{}]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orders"

    if rows:
        # Collect ALL columns across all rows to avoid missing fields
        # that only appear in some orders (e.g. JokerTag, LInseam short, etc.)
        seen = {}
        for row in rows:
            for k in row.keys():
                if k not in seen:
                    seen[k] = True
        columns = list(seen.keys())
        # Write header row
        header_fill = PatternFill("solid", fgColor="4472C4")
        header_font = Font(bold=True, color="FFFFFF")
        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        # Write data rows
        for row_idx, row in enumerate(rows, start=2):
            for col_idx, col_name in enumerate(columns, start=1):
                ws.cell(row=row_idx, column=col_idx, value=row.get(col_name, ""))

        # Auto-fit columns (approximate)
        for col_idx, col_name in enumerate(columns, start=1):
            max_len = max(len(str(col_name)), 10)
            for row_idx in range(2, min(len(rows) + 2, 102)):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val:
                    max_len = max(max_len, min(len(str(val)), 40))
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = max_len + 2

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def excel_to_json(excel_bytes: bytes) -> bytes:
    df = pd.read_excel(io.BytesIO(excel_bytes), dtype=str).fillna("")

    all_cols = list(df.columns)
    line_detail_cols = [
        c for c in all_cols
        if c not in HEADER_COLS and c not in HEADER_DETAIL_COLS and c not in LINE_COLS
    ]

    orders = []
    for order_number, group in df.groupby("OrderNumber", sort=False):
        first = group.iloc[0]

        order_header = {col: first[col] for col in HEADER_COLS if col in df.columns}

        header_detail = [
            {"ref": col, "val": first[col]}
            for col in HEADER_DETAIL_COLS
            if col in df.columns
        ]

        order_lines = []
        for _, row in group.iterrows():
            line_fields = {col: row[col] for col in LINE_COLS if col in df.columns}

            # Determine whether this line is short or regular based on which
            # inseam fields have values. Mutually exclusive: if LInseam has a
            # non-empty value use the regular pair; otherwise use the short pair.
            has_regular_inseam = bool(row.get("LInseam", "").strip()) if "LInseam" in df.columns else False
            has_short_inseam   = bool(row.get("LInseam short", "").strip()) if "LInseam short" in df.columns else False

            INSEAM_REGULAR = {"LInseam", "RInseam"}
            INSEAM_SHORT   = {"LInseam short", "RInseam short"}

            line_detail = []
            for col in line_detail_cols:
                # Skip the inseam key that doesn't apply to this line
                if col in INSEAM_REGULAR and not has_regular_inseam:
                    continue
                if col in INSEAM_SHORT and not has_short_inseam:
                    continue
                line_detail.append({"ref": col, "val": row[col]})

            order_line = {**line_fields, "OrderLineDetail": line_detail}
            order_lines.append(order_line)

        order = {
            "OrderHeader": {
                **order_header,
                "OrderHeaderDetail": header_detail,
            },
            "OrderLine": order_lines,
        }
        orders.append(order)

    result = {"Order": orders}
    return json.dumps(result, indent=2, ensure_ascii=False).encode("utf-8")

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

# Canonical order for OrderLineDetail fields.
# LInseam short / RInseam short occupy the same position as LInseam / RInseam.
LINE_DETAIL_ORDER = [
    "Style", "MainFabric", "StanttFabric", "FabricName", "PocketingFabric",
    "FrontStyle", "FrontCrease", "Hem", "FrenchFly", "WaistbandExtension",
    "WaistbandGripper", "Button", "Washing", "Dipping", "Monogram",
    "MonogramInitial", "MonogramLocation", "MonogramFont", "MonogramColor",
    "Size", "Fit", "Hip", "Waist",
    "LInseam", "RInseam",           # regular trousers
    "LInseam short", "RInseam short",  # shorts — same position slot
    "IncrRise", "DecrRise", "Thigh", "Knee", "Ankle",
    "FrontPocket", "FrontPocketReinforcement", "BackPocket", "BackPocketLabel",
    "Label", "EDI_PO", "HangTag1", "HangTag2", "Rush",
    "StandardLabel1", "StandardLabel2", "StandardLabel3",
    # JokerTag injected here when StandardLabel3 has value
    "StandardLabel4", "PackingSlipSize", "Zipper", "Rivet",
    # UPC/ColorCode/ColorName/StyleName/StoreStyle injected here when StandardLabel3 has value
    "NewAlteration",
]


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
    df = pd.read_excel(io.BytesIO(excel_bytes), dtype=str, keep_default_na=False, na_values=[]).fillna("")

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

            # Determine inseam type: regular (LInseam) or short (LInseam short)
            has_regular_inseam = bool(str(row.get("LInseam", "")).strip()) if "LInseam" in df.columns else False
            has_short_inseam   = bool(str(row.get("LInseam short", "")).strip()) if "LInseam short" in df.columns else False

            std_label3_val = str(row.get("StandardLabel3", "")).strip() if "StandardLabel3" in df.columns else ""

            # Build a lookup of all available values for this row
            row_vals = {col: row[col] for col in df.columns}

            line_detail = []
            for col in LINE_DETAIL_ORDER:
                # Skip inseam pair that doesn't apply
                if col in ("LInseam", "RInseam") and not has_regular_inseam:
                    continue
                if col in ("LInseam short", "RInseam short") and not has_short_inseam:
                    continue
                # Only emit if column exists in this Excel file
                if col not in df.columns:
                    continue
                line_detail.append({"ref": col, "val": row_vals[col]})
                # Inject JokerTag immediately after StandardLabel3 (if has value)
                if col == "StandardLabel3" and std_label3_val:
                    line_detail.append({"ref": "JokerTag", "val": std_label3_val})
                # Inject retail fields immediately after Rivet (if StandardLabel3 has value)
                if col == "Rivet" and std_label3_val:
                    for rf in ["UPC", "ColorCode", "ColorName", "StyleName", "StoreStyle"]:
                        line_detail.append({"ref": rf, "val": row_vals.get(rf, "")})

            # Append any extra columns not in LINE_DETAIL_ORDER (future-proof)
            known = set(LINE_DETAIL_ORDER) | {"JokerTag", "UPC", "ColorCode", "ColorName", "StyleName", "StoreStyle"}
            for col in line_detail_cols:
                if col not in known and col in df.columns:
                    line_detail.append({"ref": col, "val": row_vals[col]})

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

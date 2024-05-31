# Order Identification Section (Sales Order, Web Order, Purchase Order)
ORDER_ID_TYPE = "Sales Order"

# Excel Details (including mandatory Ship Set Number and SKU fields)
EXCEL_FILE_NAME = "delivery_date_example.xlsx"
SHIP_SET_COLUMN_NAME = "Ship Set Number"
SKU_COLUMN_NAME = "SKU"

# Process multiple orders on the same sheet, or Sheet Per Order feature (if multiple orders on a single sheet,
# ORDER_COLUMN must exist!)
SINGLE_SHEET = True
ORDER_COLUMN_NAME = "Sales Order"  # Ignored if SINGLE_SHEET is False

# Fields to track dictionary (mapping CCW Class fields or Line Item Fields to Excel Columns Names)
FIELDS_TO_TRACK = {"deliveryDate": "Delivery Date"}

# Configurable Parameters for special field tracking over time!
KEEP_HISTORY = True
KEEP_HISTORY_FIELD = "deliveryDate"

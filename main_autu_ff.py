from google_sheets import GoogleSheetHandler

# Sheet chứa danh sách Order ID kiểm tra
SHOES_SHEET_ID = "1Y_EnKwWThJaxLaLQyAWGojCjcahJscZPCve5qHbwGIs"
JERSEY_SHEET_ID = "13agKuW62InJ_Sdj0qA5SmiHJYjiPFqguUllLjr3CzM4"

# Các Sheet cần kiểm tra
SHEET_IDS = [
    "1iU5kAhVSC0pIP2szucrTm4PaplUh501H2oUvLgx0mw8",
    "1cGF0JBFX1dkTq_56-23IblzLKpdqgVkPxNb-ZX5-sQA"
]

# Khởi tạo handler
shoes_handler = GoogleSheetHandler(SHOES_SHEET_ID)
jersey_handler = GoogleSheetHandler(JERSEY_SHEET_ID)

# Lấy danh sách order ID từ Shoes (cột D) và Jersey (cột K)
shoes_sheet = shoes_handler.client.open_by_key(SHOES_SHEET_ID).worksheet("Shoes")
shoes_orders = [row[3] for row in shoes_sheet.get_all_values()[1:] if len(row) >= 4]  # Cột D = index 3

jersey_sheet = jersey_handler.client.open_by_key(JERSEY_SHEET_ID).worksheet("JERSEY")
jersey_data = jersey_sheet.get_all_values()
jersey_orders = [row[10] for row in jersey_sheet.get_all_values()[1:] if len(row) >= 11]  # Cột K = index 10

# Gộp thành tập hợp để tra nhanh
existing_order_ids = set(shoes_orders + jersey_orders)

# Duyệt qua từng Google Sheet cần check
for sheet_id in SHEET_IDS:
    print(f"\n📄 Đang kiểm tra Sheet: {sheet_id}")
    handler = GoogleSheetHandler(sheet_id)
    sheet1, _ = handler.get_sheets()
    data = sheet1.get_all_values()[1:]  # Bỏ tiêu đề

    for row in data:
        if len(row) > 1:
            order_id = row[1]  # Cột B = Order ID
            order_status = row[2]
            factory = row[34] 
            if order_id and order_id not in existing_order_ids and order_status == "processing":
                print(f"❗ Order ID chưa có trong Shoes/Jersey: {order_id}")
                if factory == "TP":
                    print(f"➡️ Thêm Order ID '{order_id}' vào cột K của sheet JERSEY")
                    jersey_sheet.append_row([""] * 10 + [order_id])  # Thêm dòng trống đến cột K, gán order_id
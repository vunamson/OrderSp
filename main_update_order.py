import requests
from google_sheets import GoogleSheetHandler  # Import class xử lý Google Sheets
import gspread
from gspread_formatting import (
    format_cell_ranges , set_row_heights,
    CellFormat, Color
)# 🌟 Danh sách WooCommerce Stores & Google Sheets
WOOCOMMERCE_STORES = [
    {
        "url": "https://onesimpler.com/wp-json/wc/v3/orders",
        "product_url": "https://onesimpler.com/wp-json/wc/v3/products/",
        "consumer_key": "ck_eb670ea5cee90d559872e5f29386eb3dbbb8f2da",
        "consumer_secret": "cs_cffd7acb2e5b6c5629e1a30ae580efdf73411fba",
        "sheet_id": "1j5VHpm1g3hlXK-HncynZNybubWLLmlsWt-rK5ws9UFM"
    },
    # {
    #     "url": "https://lovasuit.com/wp-json/wc/v3/orders",
    #     "product_url": "https://lovasuit.com/wp-json/wc/v3/products/",
    #     "consumer_key": "ck_30eaceb240581910ea9a679053ea7801485cd634",
    #     "consumer_secret": "cs_71f70cbb54b1cf5566867635165fa783482c6919",
    #     "sheet_id": "1oTKNUs_3XRJ7GD4C8q5ay-1JjRub2wKdOF1HDFSXEo8"
    # },
    {
        "url": "https://luxinshoes.com/wp-json/wc/v3/orders",
        "product_url": "https://luxinshoes.com/wp-json/wc/v3/products/",    
        "consumer_key": "ck_762adb5c45a88080ded28b5259e971f2274bc586",
        "consumer_secret": "cs_21df6fc65867df61725e29743f4bd6260f28d2af",
        "sheet_id": "1cGF0JBFX1dkTq_56-23IblzLKpdqgVkPxNb-ZX5-sQA"
    },
    {
        "url": "https://davidress.com/wp-json/wc/v3/orders",
        "product_url": "https://davidress.com/wp-json/wc/v3/products/",
        "consumer_key": "ck_140a74832b999d10f1f5b7b6f97ae8ddc25e835a",
        "consumer_secret": "cs_d290713d3e1199c51a22dc1e85707bb24bcce769",
        "sheet_id": "1iU5kAhVSC0pIP2szucrTm4PaplUh501H2oUvLgx0mw8"
    }
]

SHEET_SOURCES = {
    "shoes": {
        "sheet_id": "1Y_EnKwWThJaxLaLQyAWGojCjcahJscZPCve5qHbwGIs",
        "sheet_name": "Shoes",
        "order_id_col": 3,  # Cột D (Order ID)
        "checking_number_col": 2  # Cột C (Checking Number)
    },
    "cn": {
        "sheet_id": "1Y_EnKwWThJaxLaLQyAWGojCjcahJscZPCve5qHbwGIs",
        "sheet_name": "CN",
        "order_id_col": 2,  # Cột C (Order ID)
        "checking_number_col": 22  # Cột W (Checking Number)
    },
    "merchfox": {
        "sheet_id": "13agKuW62InJ_Sdj0qA5SmiHJYjiPFqguUllLjr3CzM4",
        "sheet_name": "JERSEY",
        "order_id_col": 10,  # Cột E (Order ID)
        "checking_number_col": 1  # Cột B (Checking Number)
    },
    "webbb": {
        "sheet_id": "1mCdTlRUw2OlNLBipZWycfP6CDhJ29DZBNF7zv2snoB4",
        "sheet_name": "WEB",
        "order_id_col": 1,  # Cột L (Order ID)
        "checking_number_col": 4  # Cột E (Checking Number)
    }
}

def fetch_checking_numbers():
    """
    Trả về dictionary ánh xạ Order ID → Checking Number từ nhiều nguồn dữ liệu khác nhau
    """
    checking_maps = {
        "shoes": {},
        "cn": {},
        "merchfox": {},
        "webbb": {}
    }

    for source_name, source in SHEET_SOURCES.items():
        try:
            google_sheet = GoogleSheetHandler(source["sheet_id"])
            sheet = google_sheet.client.open_by_key(source["sheet_id"]).worksheet(source["sheet_name"])
            data = sheet.get_all_values()

            # Duyệt qua từng dòng, bỏ qua dòng tiêu đề
            for row in data[1:]:
                order_id = row[source["order_id_col"]].strip() if len(row) > source["order_id_col"] else ""
                checking_number = row[source["checking_number_col"]].strip() if len(row) > source["checking_number_col"] else ""

                if order_id:
                    checking_maps[source_name][order_id] = checking_number

            print(f"✅ Lấy dữ liệu từ {source_name}: {len(checking_maps[source_name])} đơn hàng")
        except Exception as e:
            print(f"⚠️ Lỗi khi lấy dữ liệu từ {source_name}: {e}")

    return checking_maps
def check_sku_by_type(type_):
    if type_:
        if "T-Shirt" in type_:
            return "TX"
        elif "Hoodie" in type_ and "Zip" not in type_:
            return "HDZ72M"
        elif "Zip Hoodie" in type_:
            return "ZIP"
        elif "Sweatshirt" in type_:
            return "WY"
        elif "Long Sleeve" in type_:
            return "CX"
        elif "Tank Top" in type_:
            return "BX"
    return ""

def fetch_product_details(store, product_id,type_):
    """
    Trả về product_url, list_images, sku_workshop, factory dựa trên product_id
    """
    if not product_id:
        return "", "", "", ""

    product_url = f"{store['product_url']}{product_id}?consumer_key={store['consumer_key']}&consumer_secret={store['consumer_secret']}"

    try:
        response = requests.get(product_url)
        if response.status_code == 200:
            product_data = response.json()
            product_permalink = product_data.get("permalink", "")
            list_images = ", ".join(img["src"] for img in product_data.get("images", []))
            
            # 🌟 Lấy SKU workshop và Factory
            categories = product_data.get("categories", [])
            if categories:
                sku_workshop = check_sku_by_type(type_) or categories[0]["name"].split("-")[0].strip()  # Lấy phần đầu trước dấu '-'
                factory = "TP" if sku_workshop == "AODAU" else "MF"
            else:
                sku_workshop = ""
                factory = ""

            return product_permalink, list_images, sku_workshop, factory

    except Exception as e:
        print(f"⚠️ Lỗi khi lấy thông tin sản phẩm {product_id}: {e}")
    
    return "", "", "", ""

# 🌟 Lấy đơn hàng từ WooCommerce
def fetch_orders(store):
    orders = []
    page = 1
    per_page = 100

    while True:
        response = requests.get(
            f"{store['url']}?per_page={per_page}&page={page}&consumer_key={store['consumer_key']}&consumer_secret={store['consumer_secret']}"
        )
        data = response.json()
        if not data or isinstance(data, dict) and "status" in data and data["status"] != 200:
            break
        orders.extend(data)
        page += 1

    return orders


def extract_metadata_value(meta_data, keys, default=""):
    """
    Hàm trích xuất giá trị metadata từ danh sách meta_data.
    - keys: Danh sách các key cần tìm.
    - default: Giá trị mặc định nếu không tìm thấy.
    """
    for item in meta_data:
        if any(item["key"].strip().lower() == key.strip().lower() for key in keys):
            value = item.get("value", "").strip().replace(":", "").replace("|", "")
            return value if value else default
    return default


# 🌟 Xử lý dữ liệu đơn hàng
def clean_value(value):
    """
    Trả về phần đầu tiên trước khoảng trắng (nếu có).
    """
    value = str(value).strip()
    return value.split()[0] if value else ""

def process_orders(orders, existing_orders,store, checking_maps):
    new_orders = []
    updated_orders = []
    
    for order in orders:
        order_id = str(order["id"])
        order_status = order["status"]
        order_date = order["date_created"]
        order_total = order["total"]
        customer_name = f"{order['shipping']['first_name']} {order['shipping']['last_name']}"
        shipping_address_1 = order["shipping"]["address_1"]
        shipping_address_2 = order["shipping"]["address_2"]
        city = order["shipping"]["city"]
        postcode = order["shipping"]["postcode"]
        state = order["shipping"]["state"]
        country = order["shipping"]["country"]
        billing_phone = order["billing"]["phone"]
        shipping_phone = order["shipping"].get("phone", "")
        email = order["billing"]["email"]
        note = order.get("customer_note", "")
        pay_url = order.get("payment_url", "")
        shipping_total = order["shipping_total"]

        checking_number = ""
        if order_id in checking_maps["shoes"]:
            checking_number = checking_maps["shoes"][order_id]
        elif order_id in checking_maps["merchfox"]:
            checking_number = checking_maps["merchfox"][order_id]
        elif order_id in checking_maps["cn"]:
            checking_number = checking_maps["cn"][order_id]
        elif order_id in checking_maps["webbb"]:
            checking_number = checking_maps["webbb"][order_id]



        # 🌟 Kiểm tra đơn hàng đã có trong Google Sheets chưa
        if order_id in existing_orders:
            existing_row = existing_orders[order_id]
            current_status = existing_row["Order Status"]
            existing_checking_number = existing_row["Number Checking"]

            # Nếu trạng thái thay đổi, cập nhật
            if current_status != order_status or existing_checking_number != checking_number:
                updated_orders.append((order_id, order_status, checking_number))
        else:
            is_first_item = True
            # Thêm đơn hàng mới
            for item in order["line_items"]:
                product_name = item["name"]
                product_id = item["product_id"]
                sku = item["sku"]
                quantity = item["quantity"]
                unit_cost = item["price"]
                total_cost = item["total"]
                link_image = item["image"]["src"] if "image" in item else ""

                custom_name = extract_metadata_value(item["meta_data"], ["Custom Name:", "Custom Name (Optional)","Custom Name (Optional):","Custom Name"])
                custom_number = extract_metadata_value(item["meta_data"], ["Custom Number:", "Custom Number (Optional)","Custom Number (Optional):","Custom Number"])
                size = extract_metadata_value(item["meta_data"], ["Size", "SIZE:", "Size Men", "Size Women","Size:","SIZE","Size Men:","Size Women:","pa_size","size"])
                size = clean_value(size)
                color = extract_metadata_value(item["meta_data"], ["Color","Color:","COLOR", "COLOR:","Handle Color"])
                gender = extract_metadata_value(item["meta_data"], ["Gender", "Gender:"])
                gender = clean_value(gender)
                
                # Chỉ giữ lại giá trị hợp lệ cho Gender
                # valid_genders = {"Men", "Women", "Unisex"}
                # gender = gender if gender in valid_genders else ""

                type_ = extract_metadata_value(item["meta_data"], ["Type", "TYPE:", "Style:", "Stype", "STYPE","pa_style","STYPE:"])
                type_ = clean_value(type_)
                # type_ = "" if type_ in valid_genders else type_
                product_url, list_link_image, sku_workshop, factory = fetch_product_details(store, product_id ,type_)

                new_orders.append([
                    order_date, order_id, order_status, pay_url, customer_name, shipping_address_1, shipping_address_2, city, postcode,
                    state, country, billing_phone, shipping_phone, email, note, custom_name, custom_number, size, color, gender, type_,
                    product_name, product_id, sku, quantity, shipping_total, order_total if is_first_item else "", link_image, "", list_link_image,
                    product_url, unit_cost, total_cost, sku_workshop, factory, checking_number
                ])
                is_first_item = False

    return new_orders, updated_orders
def apply_formula_to_cells( sheet, column_letter):
        """
        Gán công thức IMAGE() vào cột column_letter với link ảnh là ô ngay bên phải nó.
        :param sheet: Google Sheet cần chỉnh sửa.
        :param column_letter: Vị trí của cột cần gán công thức (Ví dụ: 'AC').
        """
        try:
            data = sheet.get_all_values()
            num_rows = len(data)  # Tổng số dòng có dữ liệu

            if num_rows <= 1:
                print(f"❌ Không có đủ dữ liệu trong sheet để gán công thức.")
                return
            start_row = 2  # Bắt đầu từ dòng 2  
            end_row = num_rows

            # Xác định cột bên phải chứa link ảnh
            col_index = gspread.utils.a1_to_rowcol(column_letter + "1")[1]  # Lấy chỉ số cột (VD: 'AC' → 29)
            adjacent_col_letter = gspread.utils.rowcol_to_a1(1, col_index  -1).replace("1", "")  # Lấy cột bên trái (VD: 'AD')
            print('adjacent_col_letter' ,adjacent_col_letter , col_index)
            # Xác định phạm vi ô (từ dòng 2 đến num_rows)
            cell_range = f"{column_letter}2:{column_letter}{num_rows}"

            # Tạo danh sách công thức theo từng dòng (VD: =IMAGE(AD2))
            formulas = [[f'=IMAGE({adjacent_col_letter}{i})'] for i in range(start_row, end_row + 1)]

            # Ghi công thức vào Google Sheets
            sheet.update(range_name=cell_range, values=formulas, value_input_option="USER_ENTERED")
            print(f"✅ Công thức đã được gán vào cột {column_letter} ({cell_range}) với link ảnh từ cột {adjacent_col_letter}.")
        except Exception as e:
            print(f"❌ Lỗi khi gán công thức vào {column_letter}: {e}")


# 🌟 Cập nhật dữ liệu trong Google Sheets
def update_google_sheets(google_sheets, new_orders, updated_orders):
    sheet1, sheet2 = google_sheets.get_sheets()
    updates = []
    # Cập nhật trạng thái đơn hàng và Checking Number nếu có thay đổi
    for order_id, new_status, new_checking_number in updated_orders:
        cell = sheet1.find(order_id, in_column=2)  # Order ID là cột 2
        if cell:
            updates.append({"range": f"C{cell.row}", "values": [[new_status]]})  # Cột 3 là Order Status
            updates.append({"range": f"AJ{cell.row}", "values": [[new_checking_number]]})  # Cột 36 (AJ) là Checking Number

    if updates:
        sheet1.batch_update(updates)  # Gửi tất cả yêu cầu cập nhật cùng lúc

    # Thêm đơn hàng mới vào cuối sheet
    if new_orders:
        sheet1.append_rows(new_orders)

    print(f"✅ Đã thêm {len(new_orders)} đơn hàng mới")
    print(f"✅ Đã cập nhật {len(updated_orders)} đơn hàng")

    # 🌟 Sắp xếp lại Sheet1 theo Order Date
    google_sheets.sort_sheet(sheet1, 1)

    print("✅ Đã sắp xếp lại Sheet1 theo Order Date")
    # 🌟 Tô màu cột Order Status
    format_order_status(sheet1)
    apply_formula_to_cells(sheet1, "AC")
    # 🌟 Đặt chiều cao tất cả các hàng thành 100px
    set_row_heights_to_100(sheet1)
    print("✅ Đã tô lại màu Sheet1 theo Order Status")


def format_order_status(sheet):
    """
    Tô màu từng ô trong cột "Order Status" theo giá trị của nó:
    - `processing` → Vàng
    - `failed` → Đỏ
    - `completed` → Xanh
    """

    # Lấy tất cả dữ liệu từ cột "Order Status" (cột C)
    column_values = sheet.col_values(3)  # Cột C (Order Status)

    # Xác định danh sách các ô theo trạng thái
    processing_cells = []
    failed_cells = []
    completed_cells = []

    for row_idx, status in enumerate(column_values[1:], start=2):  # Bỏ qua tiêu đề (hàng 1)
        if status.lower() == "processing":
            processing_cells.append(f"C{row_idx}")
        elif status.lower() == "failed":
            failed_cells.append(f"C{row_idx}")
        elif status.lower() == "completed":
            completed_cells.append(f"C{row_idx}")

    # Tạo định dạng màu sắc
    processing_format = CellFormat(backgroundColor=Color(1, 1, 0))  # Màu vàng 🟡
    failed_format = CellFormat(backgroundColor=Color(1, 0, 0))  # Màu đỏ 🔴
    completed_format = CellFormat(backgroundColor=Color(0, 1, 0))  # Màu xanh 🟢

    # Áp dụng định dạng màu cho từng nhóm trạng thái
    format_cell_ranges(sheet, [(cell, processing_format) for cell in processing_cells])
    format_cell_ranges(sheet, [(cell, failed_format) for cell in failed_cells])
    format_cell_ranges(sheet, [(cell, completed_format) for cell in completed_cells])

    print("✅ Đã tô màu cột Order Status theo trạng thái đơn hàng")

def set_row_heights_to_100(sheet):
    """
    Đặt chiều cao tất cả các hàng từ 1-1000 thành 100px
    """
    set_row_heights(sheet, [('1:1000', 100)])  # Đặt chiều cao tất cả các hàng từ 1 đến 1000 là 100px
    print("✅ Đã đặt chiều cao tất cả các hàng thành 100px")

# 🌟 Chạy chương trình
if __name__ == "__main__":
    print("\n🔄 Đang tải dữ liệu Checking Numbers từ Google Sheets...")
    checking_maps = fetch_checking_numbers()  # Lấy dữ liệu checking numbers trước

    for store in WOOCOMMERCE_STORES:
        print(f"\n🚀 Đang cập nhật dữ liệu cho store: {store['url']}")

        google_sheets = GoogleSheetHandler(store["sheet_id"])
        sheet1, _ = google_sheets.get_sheets()
        existing_data = sheet1.get_all_values()
        existing_orders = {row[1]: {"Order Status": row[2],"Number Checking" : row[35]} for row in existing_data[1:]}
        orders = fetch_orders(store)
        new_orders, updated_orders = process_orders(orders, existing_orders, store, checking_maps)
        update_google_sheets(google_sheets, new_orders, updated_orders)
    print("\n🎉 Hoàn tất cập nhật đơn hàng!") 

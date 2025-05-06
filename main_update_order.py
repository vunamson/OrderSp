import time
import requests
from google_sheets import GoogleSheetHandler  # Import class xử lý Google Sheets
import gspread
from gspread_formatting import (
    format_cell_ranges , set_row_heights,
    CellFormat, Color
)# 🌟 Danh sách WooCommerce Stores & Google Sheets
WOOCOMMERCE_STORES = [
    # {
    #     "url": "https://craftedpod.com/wp-json/wc/v3/orders",
    #     "product_url": "https://craftedpod.com/wp-json/wc/v3/products/",    
    #     "consumer_key": "ck_bdf4bb6a38d558ad042c356346cfd79feddd492f",
    #     "consumer_secret": "cs_08a8e4957bc32840fbd0b0a2cf91522df0b7840c",
    #     "sheet_id": "1u7XQOeP7vegn5u1wR-HZHKv0vjV5FcOzIh1nY92F7jw"
    # },
    # {
    #     "url": "https://onesimpler.com/wp-json/wc/v3/orders",
    #     "product_url": "https://onesimpler.com/wp-json/wc/v3/products/",
    #     "consumer_key": "ck_eb670ea5cee90d559872e5f29386eb3dbbb8f2da",
    #     "consumer_secret": "cs_cffd7acb2e5b6c5629e1a30ae580efdf73411fba",
    #     "sheet_id": "1j5VHpm1g3hlXK-HncynZNybubWLLmlsWt-rK5ws9UFM"
    # },
    # {
    #     "url": "https://lovasuit.com/wp-json/wc/v3/orders",
    #     "product_url": "https://lovasuit.com/wp-json/wc/v3/products/",
    #     "consumer_key": "ck_30eaceb240581910ea9a679053ea7801485cd634",
    #     "consumer_secret": "cs_71f70cbb54b1cf5566867635165fa783482c6919",
    #     "sheet_id": "1oTKNUs_3XRJ7GD4C8q5ay-1JjRub2wKdOF1HDFSXEo8"
    # },
    {
        "url": "https://lobreve.com/wp-json/wc/v3/orders",
        "product_url": "https://lobreve.com/wp-json/wc/v3/products/",    
        "consumer_key": "ck_dfa0a1b6687f6c58ef7b3bb4fc2fcaba1f7e59c4",
        "consumer_secret": "cs_68a0b53f5d1a93d7c4bdb613c6bda038ce8aa807",
        "sheet_id": "1SinUd6nxbowMmwWiZcw16yNJsprOHtEdJl1g0pxb0fM"
    },
    {
        "url": "https://clothguy.com/wp-json/wc/v3/orders",
        "product_url": "https://clothguy.com/wp-json/wc/v3/products/",    
        "consumer_key": "ck_0af4e203af237c0877ad2bb9bfbfa46c9096f85d",
        "consumer_secret": "cs_8c084ab759b120b119713233a9ae043afeebae62",
        "sheet_id": "18Y44B205GJBhgbMrhfOdcc1dcjxsujjjFkHx49cwsU0"
    },
    {
        "url": "https://printpear.com/wp-json/wc/v3/orders",
        "product_url": "https://printpear.com/wp-json/wc/v3/products/",    
        "consumer_key": "ck_a278fd4d4402c9e094c89a87102bbc85b33360a5",
        "consumer_secret": "cs_cb6cdb01d144bd9cced7a8c2061c980cf491c084",
        "sheet_id": "1avty1G04ugUEiS5pwJPKFW0YZr8Rh-ogyro4HajZyRc"
    },
    # {
    #     "url": "https://cracksetup.com/wp-json/wc/v3/orders",
    #     "product_url": "https://cracksetup.com/wp-json/wc/v3/products/",    
    #     "consumer_key": "ck_161b3deffed9f0f1f319b774486b3a2a4ecf4d61",
    #     "consumer_secret": "cs_ccb0ca4e2a9707635b9d64e33e4038b24b252c7d",
    #     "sheet_id": "141M1T0VI6BOrsLokIxKhfzwvzSPrKgVoQKMUAwpw-Bw"
    # },
    {
        "url": "https://clomic.com/wp-json/wc/v3/orders",
        "product_url": "https://clomic.com/wp-json/wc/v3/products/",    
        "consumer_key": "ck_094b18d1bbccb1eb52ce05113a8cce6ffdc2cc06",
        "consumer_secret": "cs_62fe11d4e10536c4bd772308bfff08d40fd00e2b",
        "sheet_id": "1Eh1DQ55AmVQcg0j8q6tFUZ9d8a8V_6ugO3uxU4n9gTw"
    },
    {
        "url": "https://luxinshoes.com/wp-json/wc/v3/orders",
        "product_url": "https://luxinshoes.com/wp-json/wc/v3/products/",    
        "consumer_key": "ck_a7554487cd3d9936118d4f908f9440d06f7c4f54",
        "consumer_secret": "cs_f70a5ae3998e069213bf6bec69f04624a9ceeb6c",
        "sheet_id": "11vRLaxloprMzBe8hwrASOLetiVWZGwjEKBU2p8s11zo"
    },
    {
        "url": "https://davidress.com/wp-json/wc/v3/orders",
        "product_url": "https://davidress.com/wp-json/wc/v3/products/",
        "consumer_key": "ck_e11910c906c2b454aa065e1a240e71a71013396a",
        "consumer_secret": "cs_6565ae4a7da24853b88195eb0abd7754d26bc484",
        "sheet_id": "1SySSJt1i4lHp8Q3SlAE5VmsDfjEJ6oecxTABivAedW0"
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
        "sheet_name": "ORDER JERSEY",
        "order_id_col": 10,  # Cột E (Order ID)
        "checking_number_col": 1  # Cột B (Checking Number)
    },
    "webbb": {
        "sheet_id": "1mCdTlRUw2OlNLBipZWycfP6CDhJ29DZBNF7zv2snoB4",
        "sheet_name": "WEB",
        "order_id_col": 1,  # Cột L (Order ID)
        "checking_number_col": 4  # Cột E (Checking Number)
    },
    "hog" : {
        "sheet_id": "1jDZbTZzUG-_Sw3NXgKMjRa5YD9V3PjMkLlx78-w688Y",
        "sheet_name": "3D(BY SELLER)",
        "order_id_col": 4,  # Cột L (Order ID)
        "checking_number_col": 5  # Cột E (Checking Number)
    }
}

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
}

def fetch_checking_numbers():
    """
    Trả về dictionary ánh xạ Order ID → Checking Number từ nhiều nguồn dữ liệu khác nhau
    """
    checking_maps = {
        "shoes": {},
        "cn": {},
        "merchfox": {},
        "webbb": {},
        "hog" : {}
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
def check_sku_by_type_mf(type_):
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

def check_sku_by_type_hog(type_):
    if type_:
        if "T-Shirt" in type_:
            return "T-Shirt 3D"
        elif "Hoodie" in type_ and "Zip" not in type_:
            return "Hoodie"
        elif "Zip Hoodie" in type_:
            return "Zip Hoodie"
        elif "Sweatshirt" in type_:
            return "WY"
        elif "Long Sleeve" in type_:
            return "Long Sleeve Tee"
        elif "Tank Top" in type_:
            return "Tank Top"
    return ""

def fetch_product_details(store, product_id,type_):
    """
    Trả về product_url, list_images, sku_workshop, factory dựa trên product_id
    """
    if not product_id:
        return "", "", "", ""

    product_url = f"{store['product_url']}{product_id}?consumer_key={store['consumer_key']}&consumer_secret={store['consumer_secret']}"

    try:
        response = requests.get(product_url,headers=headers)
        if response.status_code == 200:
            product_data = response.json()
            product_permalink = product_data.get("permalink", "")
            list_images = ", ".join(img["src"] for img in product_data.get("images", []))
            
            # 🌟 Lấy SKU workshop và Factory
            categories = product_data.get("categories", [])
            if categories:
                name_parts = categories[0]["name"].split("-")
                if len(name_parts) >= 2:
                    categorie = name_parts[0].strip() if name_parts[0].strip() == "AODAU" else name_parts[1].strip()
                else:
                    categorie = name_parts[0].strip()  # hoặc gán giá trị mặc định nếu muốn
                check_sku_type_mf = check_sku_by_type_mf(type_)
                check_sku_type_hog = check_sku_by_type_hog(type_)
                if check_sku_type_mf or check_sku_type_hog :
                    check_sku_type = check_sku_type_hog +'-'+ check_sku_type_mf
                else : 
                    check_sku_type = ''
                sku_workshop = check_sku_type or categorie  # Lấy phần đầu trước dấu '-'
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
            f"{store['url']}?per_page={per_page}&page={page}&consumer_key={store['consumer_key']}&consumer_secret={store['consumer_secret']}",headers=headers
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
        elif order_id in checking_maps["hog"]:
            checking_number = checking_maps["hog"][order_id]



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
                note = extract_metadata_value(item["meta_data"], [
                    "Your Personalized Text: (Name, Number, and more...)",
                    "Customer Note",
                    "Customer Note (Custom Name, Number or etc)"
                ])
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
    order_ids_in_sheet = sheet1.col_values(2)  # Cột B là Order ID
    order_id_to_row = {oid.strip(): idx+1 for idx, oid in enumerate(order_ids_in_sheet)}

    for order_id, new_status, new_checking_number in updated_orders:
        # cell = sheet1.find(order_id, in_column=2)  # Order ID là cột 2
        # order_id_number = int(order_id)
        # if order_id_number % 100 == 0 :
        #     print(f"🕒 Tạm dừng 100 giây vì order_id {order_id} chia hết cho 100...")
        #     time.sleep(100)
        row = order_id_to_row.get(order_id)

        if row:
            updates.append({"range": f"C{row}", "values": [[new_status]]})  # Cột 3 là Order Status
            updates.append({"range": f"AJ{row}", "values": [[new_checking_number]]})  # Cột 36 (AJ) là Checking Number

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
    if processing_cells:
        format_cell_ranges(sheet, [(cell, processing_format) for cell in processing_cells])
    if failed_cells:
        format_cell_ranges(sheet, [(cell, failed_format) for cell in failed_cells])
    if completed_cells:
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
        print('11111111111111')
        sheet1, _ = google_sheets.get_sheets()
        print('2222222')
        existing_data = sheet1.get_all_values()
        print('333333333')
        existing_orders = {row[1]: {"Order Status": row[2],"Number Checking" : row[35]} for row in existing_data[1:]}
        print('444444444')
        orders = fetch_orders(store)
        print('555555555')
        new_orders, updated_orders = process_orders(orders, existing_orders, store, checking_maps)
        print('66666666')
        update_google_sheets(google_sheets, new_orders, updated_orders)
        print('777777777')
    print("\n🎉 Hoàn tất cập nhật đơn hàng!") 

import gspread
from oauth2client.service_account import ServiceAccountCredentials

class GoogleSheetHandler:
    def __init__(self, sheet_id):
        """Khởi tạo với ID Google Sheet"""
        self.sheet_id = sheet_id
        # self.client = gspread.Client(auth=None)  # Không cần xác thực, chỉ truy cập Google Sheet công khai
        self.client = self.authenticate_google_sheets()

    def authenticate_google_sheets(self):
        """Xác thực Google Sheets API"""
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
        return gspread.authorize(creds)

    def get_sheets(self):
        """Truy xuất Sheet1 và Sheet2 từ Google Sheets"""
        sheet = self.client.open_by_key(self.sheet_id)
        return sheet.worksheet("Sheet1"), sheet.worksheet("Sheet2")
    
    
    def update_cell(self, row, col, value):
        """Cập nhật giá trị vào ô (row, col) trong Sheet2"""
        try:
            sheet2 = self.get_sheets()[1]  # Lấy Sheet2
            sheet2.update_cell(row, col, value)  # Cập nhật giá trị
            print(f"✅ Đã cập nhật ô ({row}, {col}) với giá trị: {value}")
        except Exception as e:
            print(f"❌ Lỗi khi cập nhật ô ({row}, {col}): {e}")

    def update_sheet2(self):
        """Cập nhật dữ liệu từ Sheet1 vào Sheet2"""
        sheet1, sheet2 = self.get_sheets()
        data1 = sheet1.get_all_values()
        data2 = sheet2.get_all_values()

        if len(data1) <= 1:
            print("Sheet nguồn không có dữ liệu.")
            return

        # ✅ Xác định vị trí cột cần lấy trong Sheet1
        headers = data1[0]
        try:
            order_date_idx = headers.index("Order Date")
            order_id_idx = headers.index("Order ID")
            email_idx = headers.index("Email")
            name_idx = headers.index("Full Name")
            number_checking_idx = headers.index("Number Checking")
            order_status_idx = headers.index("Order Status")
            pay_url_idx = headers.index("Pay URL")
            shipping_state_idx = headers.index("Shipping State")
        except ValueError:
            print("❌ Không tìm thấy một hoặc nhiều cột cần thiết.")
            return

        print("aaaaaaa",order_date_idx,order_id_idx,email_idx,name_idx,number_checking_idx,order_status_idx,pay_url_idx,shipping_state_idx)
        # ✅ Tạo bản đồ tra cứu Order ID trong Sheet2
        dest_order_map = {}
        for j in range(1, len(data2)):
            dest_order_id = data2[j][1]  # Order ID
            dest_number_checking = data2[j][4]  # Number Checking

            if dest_order_id:
                dest_order_map[dest_order_id] = {"row": j + 1, "Number Checking": dest_number_checking}

        # ✅ Cập nhật hoặc thêm dữ liệu vào Sheet2
        new_rows = []
        updated_rows = []

        for i in range(1, len(data1)):
            row = data1[i]
            order_date = row[order_date_idx]
            order_id = row[order_id_idx]
            email = row[email_idx]
            name = row[name_idx]
            number_checking = row[number_checking_idx]
            order_status = row[order_status_idx]
            pay_url = row[pay_url_idx]
            shipping_state = row[shipping_state_idx]

            if order_id in dest_order_map:
                # Nếu `Order ID` đã tồn tại, kiểm tra `Number Checking`
                dest_row_index = dest_order_map[order_id]["row"]
                dest_number_checking = dest_order_map[order_id]["Number Checking"]

                if number_checking != dest_number_checking:
                    updated_rows.append({"row": dest_row_index, "value": number_checking})
            else:
                # Nếu `Order ID` chưa có, thêm dòng mới vào Sheet2
                new_rows.append([order_date, order_id, email, name, number_checking, "", order_status, pay_url, shipping_state])

        # ✅ Cập nhật các giá trị `Number Checking` bị thay đổi
        for update in updated_rows:
            sheet2.update_cell(update["row"], 5, update["value"])

        # ✅ Thêm các dòng mới vào `Sheet2`
        if new_rows:
            sheet2.append_rows(new_rows)
            print(f"✅ Đã thêm {len(new_rows)} đơn hàng mới vào Sheet2.")

        # ✅ Sắp xếp lại Sheet2 theo `Order Date`
        self.sort_sheet(sheet2,  1)

        num_rows = len(data2)  # Số hàng hiện tại trong Sheet2
        if num_rows > 1:  # Chỉ xóa nếu có dữ liệu
            empty_column = [[""] for _ in range(num_rows - 1)]  # Danh sách rỗng cho từng ô trong cột J
            sheet2.update(f"J2:J{num_rows}", empty_column)  # Xóa dữ liệu từ J2 đến cuối

            print(f"✅ Đã xóa tất cả dữ liệu trong cột J của Sheet2.")

    def sort_sheet(self, sheet, sort_col):
        data = sheet.get_all_values()
        headers = data[0]

        # ✅ Chuyển đổi Order ID thành số (nếu có thể)
        def parse_number(value):
            try:
                return int(value.replace(",", ""))  # Loại bỏ dấu phẩy và chuyển thành số
            except ValueError:
                return 0  # Nếu không thể chuyển đổi, đưa về 0 để xếp cuối

        # ✅ Sắp xếp theo Order ID (cột B)
        sorted_data = sorted(data[1:], key=lambda x: parse_number(x[sort_col]), reverse=True)

        # ✅ Xóa dữ liệu cũ và cập nhật dữ liệu mới
        sheet.clear()
        sheet.append_rows([headers] + sorted_data)



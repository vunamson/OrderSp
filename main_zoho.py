from datetime import datetime, timedelta
import random
from google_sheets import GoogleSheetHandler
from checking_number import Track17Selenium
from send_mail_zoho import EmailSender
from check_date_email import check_date_email, check_date_email_failed,check_date
import time

# ✅ Danh sách ID Google Sheets (Thay thế bằng danh sách của bạn)
SHEET_IDS = [
    "1WYCdD01faFIwJknZSmd_vYur2hwqHVIRKwt8BP_yX1Q",
    "1UiOMmQPkMmq0tewpiCsmyrXx7qMwW6iE21GjqzHVO7c",
    "1GGFnHXapQZNGOh71qmQi5-OdCSYnfgewK1XHDhHu4Fc",
    "14KecG--oRcj5otgvFJ8Kl16D556L9Cz32K4I3TjyBRY",
    "1LnDxYEHkJ5yxLU8KyEZhnivSYoxuYqB4b9TjoAssdSo",
    "1t55QypLzvRFUDh0BchJfU9-Y-wAQPF-06yeJ8XW-ttY",
    "1Eh1DQ55AmVQcg0j8q6tFUZ9d8a8V_6ugO3uxU4n9gTw",
    "1oTKNUs_3XRJ7GD4C8q5ay-1JjRub2wKdOF1HDFSXEo8",
]

nameStor = [
    "Newsongspost",
    "Gardenleap",
    "Vazava",
    "Magliba",
    "BOKOCOKO",
    "Drupid",
    "Clomic",
    "Lovasuit",
]

# nameStor = [ "Davidress","Luxinshoes","Onesimpler","Xanawood","Lovasuit"]
list_mail_support = [
    "support@newsongspost.com",
    "support@gardenleap.com",
    "support@vazava.com",
    "support@magliba.com",
    "support@bokocoko.com",
    "support@drupid.com",
    "supporter@clomic.com",
    "support@lovasuit.com",
]
list_company_logo_URL = [
    "https://newsongspost.com/wp-content/uploads/2025/06/Flux_Dev_Design_a_modern_and_warm_logo_for_a_POD_merchandise_w_2__1_-removebg-preview.png",
    "https://gardenleap.com/wp-content/uploads/2025/06/cropped-Flux_Dev_Create_a_stylized_logo_for_the_Gardenleap_Store_websi_0__1_-removebg-preview.png",
    "https://vazava.com/wp-content/uploads/2025/06/Flux_Dev_Design_a_modern_and_warm_logo_for_a_POD_merchandise_w_0__1_-removebg-preview.png",
    "https://magliba.com/wp-content/uploads/2025/05/Flux_Dev_Design_a_typographical_logo_for_the_magliba_website_w_1__1_-removebg-preview.png",
    "https://bokocoko.com/wp-content/uploads/2025/05/Flux_Dev_Create_a_logo_for_the_website_bokocoko_with_a_hobbies_3-removebg-preview.png",
    "https://drupid.com/wp-content/uploads/2025/05/Flux_Dev_Create_a_modern_logo_for_the_Drupid_Store_website_inc_3__1_-removebg-preview.png",
    "https://clomic.com/wp-content/uploads/2025/05/495186405_4115911682031422_5615525537022283044_n__1_-removebg-preview.png",
    "https://lovasuit.com/wp-content/uploads/2025/06/cropped-snapedit_1741597079142.png",                         
]

# ✅ Thiết lập thông tin API Gmail

zoho_keys = {
    "Newsongspost": {
                    "CLIENT_ID": "1000.LCPONTQJE0INO5UMANPG6NSXW9DW6G",
                    "CLIENT_SECRET": "115f550fa716ef4aecebc2fd7aabc451b6f1375b18",
                    "REFRESH_TOKEN": "1000.f9d892dfccc11abf88cab97e4f482f5f.de6213f5e43053955b5af380333cefe5"
    },
    "Gardenleap": {
                    "CLIENT_ID": "1000.O42S19DWA21R2HP1IAU7D1KFYFW51Y",
                    "CLIENT_SECRET": "2adbce62e067eeca32af43c7c01c1339b5c83830d1",
                    "REFRESH_TOKEN": "1000.484547bee5dd54cc606620fdcb0caf2e.e6af5047406b0f83cad55522aba73c09"
    },
    "Vazava": {
                    "CLIENT_ID": "1000.1V8IV7RE0C5GHOBE1Z8VCDSASTWWTB",
                    "CLIENT_SECRET": "f24909443ecc279754dfd39078ca0d6219928c37b2",
                    "REFRESH_TOKEN": "1000.9760730356448bab6ac92e6394153182.596ff02b9c93b0e1bad53c2a1f28d2b0"
    },
    "Magliba": {
                    "CLIENT_ID": "1000.0LWCZ499LN2L1T0058H137MEBR7SOS",
                    "CLIENT_SECRET": "d790ef54e6f9e8dea190fd4e6f19cfa774c957c0b1",
                    "REFRESH_TOKEN": "1000.96a28187add94e4f7221c03e0779762f.b92208394f66784b45b2bc1bf869d72f"
    },
    "BOKOCOKO": {
                    "CLIENT_ID": "1000.XL34AKYUM47KJQPNSRV14WDZZ3IA4D",
                    "CLIENT_SECRET": "0a68a0c6dff1785a969e6c648e3b98fddfe20bb5fd",
                    "REFRESH_TOKEN": "1000.a8d034c9522ba2bc07690a62134d31ca.08ae67dbef20e22d1ae493826ae15d2d"
    },
    "Drupid": {
                    "CLIENT_ID": "1000.BBOB3I0WXUYLWCI5EC8MGDZYDO4ATW",
                    "CLIENT_SECRET": "6496e426e01380b926469435fa6ea7539355799542",
                    "REFRESH_TOKEN": "1000.c50c117ef6d53e760c0f4db575b75f05.eab47a5c8cdc0f3a1fa86679e9d38b2d"
    },
    "Clomic": {
                    "CLIENT_ID": "1000.H8BE2HVFQ95BE50RWINW0KDO5K2SYG",
                    "CLIENT_SECRET": "001aa0bd4740783f42abd908dd8c716bc0bd9e3698",
                    "REFRESH_TOKEN": "1000.36426c77d1fecda56b731b33f1844df3.9e63f75bfc33867406c83ab76542a74b"
    },
    "Lovasuit": {
                    "CLIENT_ID": "1000.H8BE2HVFQ95BE50RWINW0KDO5K2SYG",
                    "CLIENT_SECRET": "001aa0bd4740783f42abd908dd8c716bc0bd9e3698",
                    "REFRESH_TOKEN": "1000.b4da7ea4e099bcadf141d7f2f6508d47.27a5a2135908ac3250ad503c3f55ff69"
    },
}

# TEST_EMAILS = [
#     "huongtra28922@gmail.com",
#     "huuhoi94.tb@gmail.com",
#     "trungbb.trumpany@gmail.com",
#     "nguyenthidung422003@gmail.com",
#     "minhbinh20081@gmail.com",
# ]

# ✅ Khởi tạo lớp gửi email


MAX_REQUESTS_PER_MINUTE = 80  # Dưới giới hạn Google
request_count = 0

# ✅ Hàm kiểm tra xem có đơn hàng thất bại nhưng vẫn có email không
def check_status_failure(sheet_data, email, index):
    for i in range(index):
        if sheet_data[i][2] == email and sheet_data[i][6] != "failed":
            return True
    return False

# ✅ Hàm kiểm tra nếu OrderID chưa có trạng thái
def check_order_id_no_status(sheet_data, index, status):
    if index >= len(sheet_data):
        return False 
    for i in range(index):
        if len(sheet_data[i]) > 5 and len(sheet_data[index]) > 2:
            if sheet_data[i][2] == sheet_data[index][2] and not sheet_data[i][5] and not status:
                return True
    return False

# ✅ Lặp qua từng Google Sheet ID
for index, sheet_id in enumerate(SHEET_IDS):  # Lấy index tự động
    creds = zoho_keys[nameStor[index]]
    email_sender = EmailSender(creds["CLIENT_ID"], creds["CLIENT_SECRET"], creds["REFRESH_TOKEN"])
    print(f"\n🚀 Bắt đầu xử lý Google Sheet: {sheet_id}")

    # ✅ B1: Cập nhật và sắp xếp lại Sheet2
    google_sheets = GoogleSheetHandler(sheet_id)
    print("🔄 Đang cập nhật Sheet2...")
    google_sheets.update_sheet2()

    # ✅ B2: Lấy dữ liệu từ Sheet2
    print("🔍 Lấy dữ liệu từ Sheet2...")
    sheet2 = google_sheets.get_sheets()[1]
    sheet2_data = sheet2.get_all_values()
    headers = sheet2_data[0]  # Tiêu đề cột
    data_rows = sheet2_data[1:]  # Loại bỏ tiêu đề

    # ✅ Lặp qua từng đơn hàng trong Sheet2
    for i, row in enumerate(data_rows, start=2):  # Bắt đầu từ dòng 2 (sau tiêu đề)

        if request_count >= MAX_REQUESTS_PER_MINUTE:
            print("🛑 Đạt giới hạn request, nghỉ 60 giây...")
            time.sleep(60)
            request_count = 0  # Reset bộ đếm

        row_length = len(row)
        order_date = row[0]
        email = row[2]
        customer_name = row[3]
        tracking_number = row[4]
        current_status = row[5]
        order_status = row[6]
        pay_url = row[7]
        shipping_state = row[8]
        date_status_order = row[11] if row_length > 11 else ""

        # ✅ Kiểm tra nếu đơn hàng thất bại nhưng không thuộc IL hoặc FL và không mua lại sản phẩm
        if order_status == "failed" and not check_status_failure(sheet2_data, email, i):
            if shipping_state not in ["IL", "FL"]:
                email_type = check_date_email(order_date)
                if email_type:
                    if email_type not in ["marketing", "day14"]:
                        email_sender.email_check(list_mail_support[index],email, customer_name, tracking_number, email_type + "Failed", pay_url, nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24)) # tý sửa lại
                        google_sheets.update_cell(i, 10, email_type + "Failed")  # Cập nhật cột J
                    # elif email_type == "marketing":
                    #     email_sender.email_check(list_mail_support[index],email, customer_name, tracking_number, email_type, pay_url, nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24))
                    #     google_sheets.update_cell(i, 10, email_type + "Failed")

                email_type_failed = check_date_email_failed(order_date)
                if email_type_failed:
                    email_sender.email_check(list_mail_support[index],email, customer_name, tracking_number, email_type_failed + "Failed", pay_url, nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24)) # tý sửa lại
                    google_sheets.update_cell(i, 10, email_type_failed + "Failed")

        else: #Nếu đơn hàng không failed 
            if not check_order_id_no_status(sheet2_data, i, current_status): # Không có đơn hàng nào đằng trước cùng ID và cũng chưa có status
                email_type = check_date_email(order_date)
                if tracking_number: # kiểm tra xem đã có number checking chưa
                    tracker = Track17Selenium(tracking_number)
                    new_status = tracker.track()
                    # ✅ Kiểm tra nếu new_status là dict (trả về lỗi)
                    # if isinstance(new_status, dict):
                    #     new_status = new_status.get("error", "Unknown Status")
                    # elif isinstance(new_status, list):
                    #     new_status = new_status[0]  # Nếu là list, lấy phần tử đầu tiên

                    # # ✅ Chuẩn hóa trạng thái tracking
                    # new_status = new_status.lower()
                    print('new_status' , new_status)
                    if("Info received" in new_status) : new_status = "InfoReceived"
                    elif "In transit" in new_status or new_status =="Depart from port" or new_status=="Arrived at port" : new_status = "InTransit"
                    elif("Pick up" in new_status) : new_status = "PickUp"
                    elif("Out for delivery" in new_status) : new_status = "OutForDelivery"
                    elif("Undelivered" in new_status or new_status == "Delivery Exception") : new_status = "Undelivered"
                    elif("Delivered" in new_status) : new_status = "Delivered"
                    elif("Alert" in new_status or new_status =="Package Exception") : new_status = "Alert"
                    elif("Expired" in new_status) : new_status = "Expired"
                    else: new_status=""

                    print('new_status2' , new_status)
                    # ✅ Nếu trạng thái thay đổi -> Cập nhật vào Sheet & Gửi email
                    status_order_false = ['Alert','Undelivered','Expired']
                    if new_status and new_status != current_status:
                        google_sheets.update_cell(i, 6, new_status)  # Cập nhật cột F
                        email_sender.email_check(list_mail_support[index],email, customer_name, tracking_number, new_status, "", nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24)) # tý đổi lại
                        google_sheets.update_cell(i, 10, new_status)  # Cập nhật cột J
                        request_count += 2
                        if new_status not in status_order_false : 
                            formatted_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S") 
                            google_sheets.update_cell(i, 12, formatted_datetime)
                            google_sheets.update_cell(i, 13, '')
                            request_count += 2
                        else : # Nếu order status bị lỗi gửi mail cho chính mình 
                            email_sender.email_check(list_mail_support[index],"poncealine342@gmail.com", customer_name, tracking_number, "LoiOrder", "", nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24))
                            google_sheets.update_cell(i, 13, "Loi " + new_status)
                            request_count +=1
                    else:
                        status_order_true = ['InfoReceived', 'InTransit','PickUp','OutForDelivery']
                        # sau 10 ngày và 13 ngày chưa chuyển trạng thái gửi mail cho chính mình 
                        if date_status_order and new_status : 
                            day_status = check_date(date_status_order)
                            if new_status in status_order_true and day_status == 10 :
                                email_sender.email_check(list_mail_support[index],"poncealine342@gmail.com", customer_name, tracking_number, "DelayOrder10Day", "", nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24))
                                google_sheets.update_cell(i, 13, "ngày 10 chưa chuyển trạng thái ")
                                request_count += 1
                            elif new_status in status_order_true and day_status == 13 :
                                email_sender.email_check(list_mail_support[index],"poncealine342@gmail.com", customer_name, tracking_number, "DelayOrder13Day", "", nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24))
                                google_sheets.update_cell(i, 13, "ngày 13 chưa chuyển trạng thái ")
                                request_count += 1
                elif email_type and email_type not in ["marketing", "day14"]: # nếu chưa có number checking mới bắt đầu gửi mail theo ngày
                    email_sender.email_check(list_mail_support[index],email, customer_name, tracking_number, email_type, "", nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24)) # tý sửa lại
                    google_sheets.update_cell(i, 10, email_type)
                    request_count += 2
                
                if email_type == "marketing" and current_status == "Delivered": # Gửi mail cskh
                    email_sender.email_check(list_mail_support[index],email, customer_name, tracking_number, email_type, "", nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24)) #tý sửa lại
                    google_sheets.update_cell(i, 11, email_type)
                    request_count += 2

        # ✅ Nghỉ giữa các lần chạy để tránh bị Google chặn
        if i % 20 == 0:
            print("🛑 Nghỉ 10 giây để tránh bị chặn...")
            time.sleep(10)

    print(f"✅ Hoàn tất xử lý Google Sheet: {sheet_id}")

print("🎉 Hoàn tất kiểm tra & gửi email cho tất cả Google Sheets!")

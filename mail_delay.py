from datetime import datetime, timedelta
from google_sheets import GoogleSheetHandler
from checking_number import Track17Selenium
from send_mail import EmailSender
from check_date_email import check_date_email, check_date_email_failed,check_date
import time

# ✅ Danh sách ID Google Sheets (Thay thế bằng danh sách của bạn)
SHEET_IDS = [
    "18Y44B205GJBhgbMrhfOdcc1dcjxsujjjFkHx49cwsU0",
    # "1SinUd6nxbowMmwWiZcw16yNJsprOHtEdJl1g0pxb0fM",
    # "1avty1G04ugUEiS5pwJPKFW0YZr8Rh-ogyro4HajZyRc",
    # "141M1T0VI6BOrsLokIxKhfzwvzSPrKgVoQKMUAwpw-Bw",
    # "1Eh1DQ55AmVQcg0j8q6tFUZ9d8a8V_6ugO3uxU4n9gTw",
    # "1SySSJt1i4lHp8Q3SlAE5VmsDfjEJ6oecxTABivAedW0",
    # "11vRLaxloprMzBe8hwrASOLetiVWZGwjEKBU2p8s11zo",
    # "1j5VHpm1g3hlXK-HncynZNybubWLLmlsWt-rK5ws9UFM",
    # "1CmmjO1NVG8hRe6YaurCHT4Co3GhSw39ABIwwTcv4sHw"
    # "1oTKNUs_3XRJ7GD4C8q5ay-1JjRub2wKdOF1HDFSXEo8"
]

nameStor = [
    "Clothguy",
    # "Lobreve",
    # "Printpear",
    # "Cracksetup",
    # "Clomic", 
    # "Davidress",
    # "Luxinshoes"
    ]

# nameStor = [ "Davidress","Luxinshoes","Onesimpler","Xanawood","Lovasuit"]
list_mail_support = [
    "support@clothguy.com",
    # "support@lobreve.com",
    # "support@printpear.com",
    # "support@clomic.com",
    # "support@davidress.com",
    # "support@luxinshoes.com",
    # "support@onesimpler.com",
    # "support@xanawood.com",
    # "support@lovasuit.com",
]
list_company_logo_URL = [
    "https://clothguy.com/wp-content/uploads/2025/03/cropped-Flux_Dev_Create_a_modern_vibrant_logo_for_clothguycom_a_websit_3-removebg-preview.png",
    # "https://lobreve.com/wp-content/uploads/2025/03/Lobreve-removebg-preview.png",
    # "https://printpear.com/wp-content/uploads/2025/03/cropped-Flux_Dev_Design_a_harmonious_logo_for_printpearcom_a_website_s_2-removebg-preview.png",
    # "https://clomic.com/wp-content/uploads/2025/03/Default_Design_a_modern_dynamic_logo_for_clomiccom_a_sportsthe_2_d6ac09dc-a11d-44fc-b2f9-cb4692f503d4_0.png",
    # "https://trumpany.nyc3.digitaloceanspaces.com/davidress/2024/12/12080637/DaviDress_Logo-1.png",
    # "https://trumpany.nyc3.digitaloceanspaces.com/luxinshoes/2024/12/12151154/Luxinshoes_logo.png",
    #  "https://onesimpler.com/wp-content/uploads/2025/01/Chua-co-ten-2000-x-1000-px-1.png",
    #  "https://trumpany.nyc3.digitaloceanspaces.com/xanawood.com/2025/02/24025646/Logo-Xanawood.png",
    #  "https://trumpany.nyc3.digitaloceanspaces.com/lovasuit.com/2025/02/28222122/Favicon.png"                         
]

# ✅ Thiết lập thông tin API Gmail

key_mail = {
    "Clothguy":{
        "CLIENT_ID" : "636468432266-014v2lao9j86hfa1ssgv05nfkaeksi4t.apps.googleusercontent.com",
        "CLIENT_SECRET" : "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",
        "REFRESH_TOKEN" : "1//04u19kzr3sH3TCgYIARAAGAQSNwF-L9IrnX-KddJkZXN4VgKBV9p1RM1R-Ai6L7ILx0svx0V28YqBxnw5iNDpI0TxdU5uJc8JN_c"
    },
    # "Lobreve":{
    #     "CLIENT_ID" : "417705534492-c8u3vh30tp37oav313k9ru72209dkvd5.apps.googleusercontent.com",
    #     "CLIENT_SECRET" : "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",
    #     "REFRESH_TOKEN" : "1//04SGNK28F11KKCgYIARAAGAQSNwF-L9IrcQrSG080Yd5UG2a94KHssbk8Sf2ieX4kVppNU0u-cUU5xZHD3orgdJHJlBhlLTKYLD0"
    # },
    # "Printpear":{
    #     "CLIENT_ID" : "791584904106-p863n6duo8drgtb7f5sam2msu234pqke.apps.googleusercontent.com",
    #     "CLIENT_SECRET" : "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",
    #     "REFRESH_TOKEN" : "1//040Ve4VLX-VcZCgYIARAAGAQSNwF-L9IrImKEuI6gLrOOtHpEwOeCZDdDFyib1KnKITktIqq0Rt6uEPRb1lsREEl8wuJiIje5jZU"
    # },
    # "Clomic":{
    #     "CLIENT_ID" : "208673125837-6rdum7k9fofeoka7u05tmbobivp9a3d2.apps.googleusercontent.com",
    #     "CLIENT_SECRET" : "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",
    #     "REFRESH_TOKEN" : "1//04GtXX9V3tl1tCgYIARAAGAQSNwF-L9Ir1gzksRZIpGL4L-uq523k5k3t89WUi_OIxs_fwQaeqKd3DLRsUW58Yfz889UKJqm8jzo"
    # },
    # "Davidress" : {
    #     "CLIENT_ID" : "815774674800-76rs0q4hr70ihac5e0bkojd4borr33q8.apps.googleusercontent.com",
    #     "CLIENT_SECRET" : "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",
    #     "REFRESH_TOKEN" : "1//04l2PRnjxiWP1CgYIARAAGAQSNwF-L9IrSyYKloXSLIDXPDmKg0AEyExfWHshUGvOuRPtdizbUSBxaDxUke7nQG6xRxXGO3PUgiY"
    # },
    # "Luxinshoes" : {
    #     "CLIENT_ID" : "21574557297-0nhvrl2k8rof50q7fmu4amoleii97sh4.apps.googleusercontent.com",
    #     "CLIENT_SECRET" : "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",
    #     "REFRESH_TOKEN" : "1//04E-tOajsSHlKCgYIARAAGAQSNwF-L9IrnIaaQl5ezgqjEbd2mRpueH7bfESoQcZJx8oYU_67cscyMVhPkBudJjW6PRlWEQbf7ns"
    # },
    # "Onesimpler" : {
    #     "CLIENT_ID" : "802842070292-fdpnac2kp98gcpjb5tspphjbb5obsvr6.apps.googleusercontent.com",
    #     "CLIENT_SECRET" : "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",
    #     "REFRESH_TOKEN" : "1//04hyLCMlWOcWtCgYIARAAGAQSNwF-L9IrS0ofz-gTJklz3CuVAcBPc2yrxrvNagCjmJonNFy5EMu47JGtkyfEuzRGEVsGkwm-ti0"
    # },
    # "Xanawood" : {
    #     "CLIENT_ID" : "33001047069-o5lltvudmh92qnb392ti1h6bj7geccp2.apps.googleusercontent.com",
    #     "CLIENT_SECRET" : "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",
    #     "REFRESH_TOKEN" : "1//04PdMFDR0Cn0YCgYIARAAGAQSNwF-L9Ir2oyc1v-F4XR5eRLDHgC5zlQdDh8lxrsQs2-iXF_EdINjEpMAD65_QzYxxRIS2Nm1DLg"
    # },
}
# CLIENT_ID = "21574557297-0nhvrl2k8rof50q7fmu4amoleii97sh4.apps.googleusercontent.com"
# CLIENT_SECRET = "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"
# REFRESH_TOKEN = "1//04E-tOajsSHlKCgYIARAAGAQSNwF-L9IrnIaaQl5ezgqjEbd2mRpueH7bfESoQcZJx8oYU_67cscyMVhPkBudJjW6PRlWEQbf7ns"


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
    email_sender = EmailSender(key_mail[nameStor[index]]["CLIENT_ID"],key_mail[nameStor[index]]["CLIENT_SECRET"],key_mail[nameStor[index]]["REFRESH_TOKEN"])
    # email_sender = EmailSender(key_mail["Luxinshoes"]["CLIENT_ID"], key_mail["Luxinshoes"]["CLIENT_SECRET"], key_mail["Luxinshoes"]["REFRESH_TOKEN"])
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
        if order_status == "failed" :
            pass

        else: #Nếu đơn hàng không failed 
            if not check_order_id_no_status(sheet2_data, i, current_status): # Không có đơn hàng nào đằng trước cùng ID và cũng chưa có status
                try:
                    order_dt = datetime.fromisoformat(order_date)
                    if order_dt.month != 6 or order_dt.day not in [27,28, 29, 30]:
                        continue  # Bỏ qua đơn không nằm trong ngày 28, 29, 30 tháng 6
                    else : 
                        email_sender.email_check(list_mail_support[index],email, customer_name, "", "Delay", pay_url, nameStor[index],list_company_logo_URL[index],datetime.now() + timedelta(hours=24))
                except Exception as e:
                    print(f"❌ Lỗi khi xử lý ngày '{order_date}' tại dòng {i}: {e}")
                    continue

                

        # ✅ Nghỉ giữa các lần chạy để tránh bị Google chặn
        if i % 20 == 0:
            print("🛑 Nghỉ 10 giây để tránh bị chặn...")
            time.sleep(10)

    print(f"✅ Hoàn tất xử lý Google Sheet: {sheet_id}")

print("🎉 Hoàn tất kiểm tra & gửi email cho tất cả Google Sheets!")

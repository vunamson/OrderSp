import time
import requests
import pytz
from google_sheets import GoogleSheetHandler
from datetime import datetime

# 🚀 Danh sách các Google Sheet ID và cấu hình WooCommerce Store
SHEET_AND_STORES = {
    # "1u7XQOeP7vegn5u1wR-HZHKv0vjV5FcOzIh1nY92F7jw": {
    #     "url": "https://craftedpod.com/",
    #     "consumer_key": "ck_bdf4bb6a38d558ad042c356346cfd79feddd492f",
    #     "consumer_secret": "cs_08a8e4957bc32840fbd0b0a2cf91522df0b7840c",
    #     "type_date": "Etc/GMT+2"
    # },
    # "1j5VHpm1g3hlXK-HncynZNybubWLLmlsWt-rK5ws9UFM": {
    #     "url": "https://onesimpler.com/",
    #     "consumer_key": "ck_eb670ea5cee90d559872e5f29386eb3dbbb8f2da",
    #     "consumer_secret": "cs_cffd7acb2e5b6c5629e1a30ae580efdf73411fba",
    #     "type_date": "Etc/GMT+4"
    # },
    "1oTKNUs_3XRJ7GD4C8q5ay-1JjRub2wKdOF1HDFSXEo8": {
        "url": "https://lovasuit.com/",
        "consumer_key": "ck_6609fc6bd730a925c9fc16e2445e0d433abc323d",
        "consumer_secret": "cs_9265f5db3482e3210fef476cdd85944c1d05f830",
        "type_date": "Etc/GMT+0"
    },
    "1oATa0YEllGkC8aFWiElzWO0nJmp2652mhqyvq3sVnOo": {
        "url": "https://noaweather.com/",
        "consumer_key": "ck_3c4184984f798639b393c9a610a4ca1910013640",
        "consumer_secret": "cs_4c93f7bb12b043b87c7af9685367e73dbfde044d",
        "type_date": "Etc/GMT+0"
    },
    "18Y44B205GJBhgbMrhfOdcc1dcjxsujjjFkHx49cwsU0": {
        "url": "https://clothguy.com/",
        "consumer_key": "ck_0af4e203af237c0877ad2bb9bfbfa46c9096f85d",
        "consumer_secret": "cs_8c084ab759b120b119713233a9ae043afeebae62",
        "type_date": "Etc/GMT+0"
    },
    "1SinUd6nxbowMmwWiZcw16yNJsprOHtEdJl1g0pxb0fM": {
        "url": "https://lobreve.com/",
        "consumer_key": "ck_dfa0a1b6687f6c58ef7b3bb4fc2fcaba1f7e59c4",
        "consumer_secret": "cs_68a0b53f5d1a93d7c4bdb613c6bda038ce8aa807",
        "type_date": "Etc/GMT+0"
    },
    "1avty1G04ugUEiS5pwJPKFW0YZr8Rh-ogyro4HajZyRc": {
        "url": "https://printpear.com/",
        "consumer_key": "ck_be16945fe0444e5e6c9e928f8be6e48e169c8dd3",
        "consumer_secret": "cs_75c0c0fcbcdb7a2975614b1abaa9b35ebe96b1f4",
        "type_date": "Etc/GMT+0"
    },
    "1Eh1DQ55AmVQcg0j8q6tFUZ9d8a8V_6ugO3uxU4n9gTw": {
        "url": "https://clomic.com/",
        "consumer_key": "ck_6650e61b14dcf29b5f8f213d5c2aa83f011582e6",
        "consumer_secret": "cs_6615d190132269f17595881a7dc23ee03d638732",
        "type_date": "Etc/GMT+0"
    },
    "1SySSJt1i4lHp8Q3SlAE5VmsDfjEJ6oecxTABivAedW0": {
        "url": "https://davidress.com/",
        "consumer_key": "ck_e11910c906c2b454aa065e1a240e71a71013396a",
        "consumer_secret": "cs_6565ae4a7da24853b88195eb0abd7754d26bc484",
        "type_date": "Etc/GMT+5"
    },
    "11vRLaxloprMzBe8hwrASOLetiVWZGwjEKBU2p8s11zo": {
        "url": "https://luxinshoes.com/",
        "consumer_key": "ck_a7554487cd3d9936118d4f908f9440d06f7c4f54",
        "consumer_secret": "cs_f70a5ae3998e069213bf6bec69f04624a9ceeb6c",
        "type_date": "Etc/GMT+4"
    },
    # Thêm các store khác nếu có
}

def get_order_tracking(order_id, store_config):
    """
    Lấy mã theo dõi của đơn hàng từ WooCommerce.
    """
    try:
        url = f"{store_config['url']}wp-json/wc-shipment-tracking/v3/orders/{order_id}/shipment-trackings/"
        
        # Gửi yêu cầu GET đến API Shipment Tracking với thông tin xác thực
        response = requests.get(url, auth=(store_config['consumer_key'], store_config['consumer_secret']))
        print('response xxxxx' ,response.json())
        if (response.status_code == 200 or response.status_code == 201) and len(response.json()) != 0 :
           return True
        return None  # Không có mã theo dõi
    except requests.exceptions.RequestException as e:
        print(f"❌ Lỗi khi lấy mã theo dõi cho đơn hàng {order_id}: {e}")
        return None
    
def get_current_time_in_timezone(timezone_str):
    """Trả về thời gian hiện tại theo múi giờ cho trước"""
    try:
        # Lấy múi giờ theo chuỗi (ví dụ "UTC-5", "UTC+7")
        timezone = pytz.timezone(timezone_str)
        # Lấy thời gian hiện tại theo múi giờ đã chọn
        current_time = datetime.now(timezone)
        # Định dạng lại thời gian theo định dạng "YYYY-MM-DD" (chỉ ngày)
        return current_time.strftime("%Y-%m-%d")
    except pytz.UnknownTimeZoneError:
        print("Múi giờ không hợp lệ!")
        return None

def update_order_tracking(order_id, tracking_number, store_config):
    """
    Cập nhật mã theo dõi cho đơn hàng trong WooCommerce.
    """
    time.sleep(2)
    try:
        url = f"{store_config['url']}wp-json/wc-shipment-tracking/v3/orders/{order_id}/shipment-trackings"
        
        # Dữ liệu gửi đi, bao gồm tracking number và provider
        try : 
            payload = {
                "custom_tracking_provider": "Custom Provider",  # Thay provider name theo yêu cầu
                "tracking_number": tracking_number,
                "custom_tracking_link" :  f"https://t.17track.net/en#nums={tracking_number}",
                "date_shipped": get_current_time_in_timezone(store_config["type_date"])
            }
            # Gửi yêu cầu POST tới API Shipment Tracking
            response = requests.post(
                url,
                json=payload,
                auth=(store_config['consumer_key'], store_config['consumer_secret']),  # Thêm auth với consumer_key và consumer_secret
                headers={"Content-Type": "application/json"}
            )
        except Exception as e:
            print(f"❌ Lỗi khi xử lý cập nhật mã theo dõi {order_id}: {e}")
        
        if response.status_code == 200 or response.status_code == 201 :
            print(f"✅ Cập nhật thành công mã theo dõi cho đơn hàng {order_id}.")
            
            # Sau khi cập nhật mã theo dõi, cập nhật trạng thái của đơn hàng
            status_data = {
                "status": "completed"  # Đặt trạng thái thành "completed"
            }
            
            # Thực hiện PUT request để cập nhật trạng thái
            status_url = f"{store_config['url']}/wp-json/wc/v3/orders/{order_id}?consumer_key={store_config['consumer_key']}&consumer_secret={store_config['consumer_secret']}"
            status_response = requests.put(status_url, json=status_data)

            if status_response.status_code == 200:
                print(f"✅ Trạng thái của đơn hàng {order_id} đã được thay đổi thành 'completed'.")
            else:
                print(f"❌ Lỗi khi cập nhật trạng thái cho đơn hàng {order_id}. Mã lỗi: {status_response.status_code}")
        else:
            print(f"❌ Lỗi khi cập nhật mã theo dõi cho đơn hàng {order_id}. Mã lỗi: {response.status_code}")
    except requests.exceptions.RequestException as e:
        print(f"❌ Lỗi khi gửi yêu cầu cập nhật mã theo dõi cho đơn hàng {order_id}: {e}")


def process_orders(sheet_id, store_config ):
    """
    Quản lý và xử lý đơn hàng từ Google Sheets
    """
    try:
        google_sheets = GoogleSheetHandler(sheet_id)
        sheet1, _ = google_sheets.get_sheets()
        data = sheet1.get_all_values()
        
        if len(data) <= 1:
            print("❌ Không có dữ liệu trong Sheet.")
            return
        
        headers = data[0]
        order_id_idx = headers.index("Order ID")
        number_checking_idx = headers.index("Number Checking")

        for row in data[1:]:  # Bỏ qua tiêu đề
            order_id = row[order_id_idx].strip()
            number_checking = row[number_checking_idx].strip()
            if order_id and number_checking and number_checking != "Cancel":  # Kiểm tra nếu có mã theo dõi
                print(f"🔍 Đang kiểm tra mã theo dõi cho đơn hàng: {order_id}")
                # Kiểm tra mã theo dõi
                existing_tracking = get_order_tracking(order_id, store_config)
                if existing_tracking:
                    print(f"✅ Đơn hàng {order_id} đã có mã theo dõi: {existing_tracking}. Bỏ qua.")
                else:
                    print(f"❌ Đơn hàng {order_id} chưa có mã theo dõi. Thêm mã theo dõi mới.")
                    # Nếu chưa có mã theo dõi, thực hiện thêm mã theo dõi vào WooCommerce
                    tracking_number = number_checking  # Tạo mã theo dõi ví dụ: "YT12345"
                    update_order_tracking(order_id, tracking_number, store_config)  # Cập nhật mã theo dõi vào WooCommerce                    
            else:
                print(f"✅ Đơn hàng {order_id} chưa có Checking Number. Bỏ qua.")
    except Exception as e:
        print(f"❌ Lỗi khi xử lý đơn hàng trong Sheet {sheet_id}: {e}")

def main():
    """
    Chạy chương trình với các Sheet ID và store tương ứng.
    """
    for sheet_id, store_config in SHEET_AND_STORES.items():
        print(f"\n🚀 Đang xử lý Sheet với ID: {sheet_id} và WooCommerce Store: {store_config['url']}")
        process_orders(sheet_id, store_config )
    print("\n🎉 Hoàn tất xử lý đơn hàng và cập nhật mã theo dõi!")

if __name__ == "__main__":
    main()

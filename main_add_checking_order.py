import requests
import pytz
from google_sheets import GoogleSheetHandler
from datetime import datetime

# 🚀 Danh sách các Google Sheet ID và cấu hình WooCommerce Store
SHEET_AND_STORES = {
    "1u7XQOeP7vegn5u1wR-HZHKv0vjV5FcOzIh1nY92F7jw": {
        "url": "https://craftedpod.com/",
        "consumer_key": "ck_bdf4bb6a38d558ad042c356346cfd79feddd492f",
        "consumer_secret": "cs_08a8e4957bc32840fbd0b0a2cf91522df0b7840c",
        "type_date": "Etc/GMT+2"
    },
    # "1iU5kAhVSC0pIP2szucrTm4PaplUh501H2oUvLgx0mw8": {
    #     "url": "https://davidress.com/wp-json/wc/v3/orders",
    #     "consumer_key": "ck_140a74832b999d10f1f5b7b6f97ae8ddc25e835a",
    #     "consumer_secret": "cs_d290713d3e1199c51a22dc1e85707bb24bcce769",
    #     "type_date": "UTC-5"
    # },
    # "1cGF0JBFX1dkTq_56-23IblzLKpdqgVkPxNb-ZX5-sQA": {
    #     "url": "https://luxinshoes.com/wp-json/wc/v3/orders",
    #     "consumer_key": "ck_762adb5c45a88080ded28b5259e971f2274bc586",
    #     "consumer_secret": "cs_21df6fc65867df61725e29743f4bd6260f28d2af",
    #     "type_date": "UTC-4"
    # },
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
    try:
        url = f"{store_config['url']}wp-json/wc-shipment-tracking/v3/orders/{order_id}/shipment-trackings"
        
        # Dữ liệu gửi đi, bao gồm tracking number và provider
        payload = {
            "custom_tracking_provider": "Custom Provider",  # Thay provider name theo yêu cầu
            "tracking_number": tracking_number,
            "custom_tracking_link" :  "https://t.17track.net/en#nums={tracking_number}",
            "date_shipped": get_current_time_in_timezone(store_config["type_date"])
        }
        # Gửi yêu cầu POST tới API Shipment Tracking
        response = requests.post(
            url,
            json=payload,
            auth=(store_config['consumer_key'], store_config['consumer_secret']),  # Thêm auth với consumer_key và consumer_secret
            headers={"Content-Type": "application/json"}
        )
        
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
            if order_id and number_checking:  # Kiểm tra nếu có mã theo dõi
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

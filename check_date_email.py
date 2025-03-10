from datetime import datetime

def check_date_email(order_date):
    try:
        # ✅ Kiểm tra nếu order_date chứa thời gian -> Cắt bỏ thời gian, chỉ lấy ngày
        if "/" in order_date:
             order_date = datetime.strptime(order_date, "%m/%d/%Y").date()  # Chuyển đổi từ MM/DD/YYYY
        else : 
            order_date = order_date.split()[0]  # Lấy phần ngày, bỏ thời gian
            order_date = datetime.strptime(order_date, "%Y-%m-%d").date()  # Chuyển thành datetime
        
        today = datetime.today().date()
        diff_days = (today - order_date).days  # Tính số ngày chênh lệch

        if 1 <= diff_days < 2:
            return "day1"
        elif 2 <= diff_days < 3:
            return "day2"
        elif 3 <= diff_days < 4:
            return "day3"
        elif 4 <= diff_days < 5:
            return "day4"
        elif 5 <= diff_days < 6:
            return "day5"
        elif 20 <= diff_days < 21:
            return "marketing"
        else:
            return False
    except Exception as e:
        print(f"⚠️ Lỗi khi xử lý ngày tháng: {e}")
        return False


def check_date_email_failed(order_date):
    """Xác định loại email cần gửi cho đơn hàng thất bại"""
    today = datetime.today().date()
    if "/" in order_date:
        order_date = datetime.strptime(order_date, "%m/%d/%Y").date()  # Chuyển đổi từ MM/DD/YYYY
    else : 
        order_date = order_date.split()[0]  # Lấy phần ngày, bỏ thời gian
        order_date = datetime.strptime(order_date, "%Y-%m-%d").date()  # Chuyển thành datetime
    diff_days = (today - order_date).days

    if 6 <= diff_days < 7:
        return "day6"
    elif 7 <= diff_days < 8:
        return "day7"
    else:
        return False

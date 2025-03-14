from lenful_api import LenfulAPI

# Khai báo thông tin tài khoản Lenful
EMAIL = "duyentb2503@gmail.com"  # Thay bằng email đăng ký Lenful
PASSWORD = "Hung11078"        # Thay bằng mật khẩu

if __name__ == "__main__":
    # Khởi tạo đối tượng API
    lenful = LenfulAPI(EMAIL, PASSWORD)
    
    # Đăng nhập để lấy token
    try:
        lenful.login()
        print("Đăng nhập thành công!")
        
        # Lấy danh sách đơn hàng
        orders = lenful.get_orders()
        print("Danh sách đơn hàng:", orders)
    except Exception as e:
        print("Lỗi:", e)
import requests

class LenfulAPI:
    BASE_URL = "https://api.lenful.com/v1.0"
    
    def __init__(self, email, password):
        self.email = email
        self.password = password
        self.token = None

    def login(self):
        """Đăng nhập vào Lenful và lấy token xác thực"""
        url = f"{self.BASE_URL}/login"
        payload = {
            "email": self.email,
            "password": self.password
        }
        response = requests.post(url, json=payload)
        if response.status_code == 200:
            data = response.json()
            self.token = data.get("data", {}).get("access_token")
            return self.token
        else:
            raise Exception(f"Lỗi đăng nhập: {response.json()}")

    def get_orders(self):
        """Lấy danh sách đơn hàng"""
        if not self.token:
            raise Exception("Bạn chưa đăng nhập. Gọi phương thức login() trước.")

        url = f"{self.BASE_URL}/orders"
        headers = {
            "Authorization": f"Bearer {self.token}"
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json()
        else:
            raise Exception(f"Lỗi lấy danh sách đơn hàng: {response.json()}")
from datetime import datetime
import random
from sys import _xoptions
from bs4 import BeautifulSoup
import undetected_chromedriver as uc
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

def get_random_user_agent():    
    chrome_version = f"{random.randint(2, 200)}.0.0.0"  # Tạo số ngẫu nhiên từ 2.0.0.0 - 200.0.0.0
    return f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{chrome_version} Safari/537.36"


def create_driver(): 
    chrome_options = uc.ChromeOptions()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.set_capability('LT:Options', _xoptions)

    # if proxy:
    #     chrome_options.add_argument(f"--proxy-server={proxy}")
    #     print(f"🔄 Đang sử dụng Proxy: {proxy}")
    # 🔄 Thêm User-Agent ngẫu nhiên
    user_agent = get_random_user_agent()
    chrome_options.add_argument(f"user-agent={user_agent}")
    
    print(f"🆕 Đã thay đổi User-Agent: {user_agent}")
    # return uc.Chrome(headless=False)
    return uc.Chrome(headless=False)

driver = create_driver()

class Track17Selenium:
    BASE_URL = "https://t.17track.net/en"
    # driver = create_driver()
    def __init__(self, tracking_number):
        self.tracking_number = tracking_number
        # self.driver = create_driver()

    def track(self):
        try:
            # driver = self.driver
            driver.get(self.BASE_URL)
            wait = WebDriverWait(driver, 10)

            search_box = wait.until(EC.presence_of_element_located((By.ID, "auto-size-textarea")))
            search_box.clear()
            search_box.send_keys(self.tracking_number)
            print("✅ Đã nhập mã tracking" , self.tracking_number)
            # Lấy dữ liệu tracking
            track_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "batch_track_search-area__9BaOs")))
            track_button.click()
            print("🚀 Đã nhấn nút TRACK")
            tracking_data = []
           

            try:
                status_link = WebDriverWait(driver, 30).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.status-progress-title"))
                )
                print('status_link',status_link)
                
                tracking_data = status_link[0].text.rstrip()
                print(f"📌 Trạng thái vận chuyển: {tracking_data}")
            except:
                print("except")
            return tracking_data if tracking_data else {"error": "No tracking information found"}

        except Exception as e:
            # driver.quit()
            return {"error": f"Failed to fetch tracking data: {str(e)}"}
        # finally:
        #     # Đảm bảo luôn quit driver
        #     try:
        #         driver.quit()
        #     except Exception:
        #         pass


# Sử dụng
# if __name__ == "__main__":
#     tracking_number = "UK049320257YP"  # Thay bằng mã tracking thực tế
#     tracker = Track17Selenium(tracking_number)
#     print(tracker.track())

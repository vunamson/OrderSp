import base64
from datetime import datetime
import requests
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import random

class EmailSender:
    def __init__(self, client_id, client_secret, refresh_token):
        """Khởi tạo với thông tin OAuth2"""
        self.client_id = client_id
        self.client_secret = client_secret
        self.refresh_token = refresh_token

    def load_email_template(self,file_path, replacements):
        """Đọc file HTML và thay thế nội dung động"""
        with open(file_path, "r", encoding="utf-8") as file:
            html_content = file.read()
        
        # Thay thế các giá trị động trong template
        for key, value in replacements.items():
            if isinstance(value, datetime):  # Nếu value là datetime, chuyển sang chuỗi
                value = value.strftime("%Y-%m-%d %H:%M:%S")  # Định dạng ngày giờ
            html_content = html_content.replace(f"{{{key}}}", str(value))  # Chuyển đổi mọi giá trị thành chuỗi
        
        return html_content

    def get_new_access_token(self):
        """Lấy Access Token mới bằng Refresh Token"""
        token_url = "https://oauth2.googleapis.com/token"
        payload = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "refresh_token": self.refresh_token,
            "grant_type": "refresh_token",
        }
        response = requests.post(token_url, data=payload)
        result = response.json()
        
        if "access_token" in result:
            return result["access_token"]
        else:
            raise Exception(f"Lỗi lấy Access Token: {result}")

    def send_email_gmail_api(self,emailSupport, email, subject,replacements,url_html_mail):
        """Gửi email qua Gmail API"""
        access_token = self.get_new_access_token()  # Lấy Access Token mới
        html_body = self.load_email_template(url_html_mail,replacements)
        message = MIMEMultipart()
        message["From"] = emailSupport
        message["To"] = email
        message["Subject"] = subject
        message.attach(MIMEText(html_body, "html"))

        base64_encoded_email = base64.urlsafe_b64encode(message.as_string().encode("utf-8")).decode("utf-8")

        # email_content = f"From: me\nTo: {email}\nSubject: {subject}\n\n{body}"
        # base64_encoded_email = base64.urlsafe_b64encode(email_content.encode("utf-8")).decode("utf-8")

        url = "https://www.googleapis.com/gmail/v1/users/me/messages/send"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        payload = json.dumps({"raw": base64_encoded_email})

        response = requests.post(url, headers=headers, data=payload)
        
        if response.status_code == 200 or response.status_code == 202:
            print(f"✅ Email gửi thành công đến {email}")
        else:
            print(f"❌ Lỗi gửi email: {response.text}")

    def email_check(self,emailSupport, email, customer_name, tracking_number, status, checkout_url, store_name,company_logo_URL,expiration_date):

        testimonials = [
            f"{store_name} is a game-changer for pet lovers! The quality of their personalized clothing is unmatched, and the designs are absolutely adorable.",
            f"I was amazed at the level of detail in my custom t-shirt! {store_name} truly knows how to make pet lovers feel special.",
            f"Finally, a store that understands pet owners! The fabric is soft, the print is durable, and my cat's face on my sweatshirt looks incredibly realistic.",
            f"{store_name} exceeded my expectations! The hoodie I ordered with my pup's image is so cozy, and the design is even better than I imagined.",
            f"As a die-hard dog lover, I couldn't be happier with my personalized tee. The ordering process was seamless, and the result was pure perfection!",
            f"I've never received so many compliments on a hoodie before! The print of my cat is stunning, and the fit is just right. {store_name} has earned a lifelong customer!",
            f"I love how I can showcase my love for my pets through fashion! {store_name} makes it so easy to create something meaningful and stylish.",
            f"The quality of their clothing is truly top-tier. My dog's portrait on my hoodie is vibrant and sharp, and the material is so soft!",
            f"Not only does {store_name} offer fantastic designs, but their customer support is also outstanding. They helped me customize my sweatshirt perfectly!",
            f"I bought a personalized t-shirt for my wife with our two cats on it, and she absolutely loved it! The attention to detail is incredible.",
            f"{store_name} is the best place for pet lovers who want unique and high-quality fashion. Their designs are creative, and their service is excellent!",
            f"I was skeptical about ordering a custom hoodie online, but {store_name} delivered beyond my expectations. The print is vibrant, and the fabric is super comfy!",
            f"The best thing about {store_name} is their passion for pets and quality. I feel proud wearing my hoodie with my dog's face on it!",
            f"I love how easy it was to customize my order. The result? A stunning, personalized hoodie featuring my beloved cat. Simply amazing!",
            f"The colors, the fit, the print—everything about my hoodie is perfect. {store_name} truly delivers on quality!",
            f"I purchased a custom shirt for my best friend, and she was blown away by how beautiful it turned out. A perfect gift for any pet lover!",
            f"The moment I put on my personalized sweatshirt, I knew it was worth every penny. {store_name} creates wearable memories!",
            f"I'm beyond impressed with the print quality and comfort of my hoodie. {store_name} is now my go-to store for pet-themed fashion!",
            f"I have ordered from many custom clothing stores, but {store_name} stands out with its exceptional craftsmanship and attention to detail!",
            f"This is hands down the best personalized apparel I've ever bought. My dog's picture on my hoodie looks stunning, and the fabric is incredibly soft!",
            f"Wearing my customized {store_name} hoodie feels like having my pet with me wherever I go. Thank you for making something so special!",
            f"From order to delivery, everything was seamless. {store_name} is the ultimate destination for pet lovers who want something stylish yet meaningful!",
            f"The print detail is mind-blowing! My cat's face on my sweatshirt looks exactly like her. I couldn't be happier!",
            f"I ordered matching t-shirts for me and my dog, and they turned out incredible! {store_name}'s designs are second to none!",
            f"This brand really understands what pet lovers want—comfortable, high-quality fashion that makes a statement!",
            f"My custom hoodie arrived quickly, and it's everything I wanted and more. Highly recommend {store_name}!",
            f"The best gift for pet lovers! My sister cried happy tears when she saw her personalized sweatshirt with her dog's face on it.",
            f"{store_name} delivers not just a product but an experience! The customization process was smooth, and the final product is amazing!",
            f"Every pet owner needs at least one {store_name} hoodie in their wardrobe. The perfect blend of comfort and personal touch!",
            f"Exceptional quality, beautiful designs, and fast shipping. {store_name} is my new favorite store!",
            f"My friends keep asking where I got my hoodie from! {store_name}, you have a loyal customer now!",
            f"Perfect for gifts or personal use! My personalized t-shirt turned out better than expected.",
            f"{store_name}'s attention to detail is incredible. They truly make pet lovers feel special!",
            f"I can't believe how soft and durable my hoodie is! The print quality is amazing!",
            f"Worth every penny! The customer service was helpful, and the product is outstanding!",
            f"I feel so happy wearing my dog's face on my hoodie! {store_name} makes fashion fun and meaningful!",
            f"The best pet-themed clothing brand out there! Unique, stylish, and so comfortable!",
            f"Every time I wear my {store_name} hoodie, people ask where I got it. I proudly tell them!",
            f"My personalized sweatshirt is a conversation starter! People love the design!",
            f"High-quality, unique, and affordable—{store_name} has it all!",
            f"I ordered two hoodies, and both are amazing! Great fit and perfect prints.",
            f"The best way to express my love for my pet! Thank you, {store_name}!",
            f"The print hasn't faded even after multiple washes. Super durable!",
            f"Stylish and comfortable—{store_name} never disappoints!",
            f"Perfect for pet lovers! I adore my custom hoodie!",
            f"The detail on my cat's portrait is stunning. I love it!",
            f"A must-have for any pet owner. Love my hoodie!",
            f"{store_name} has exceeded my expectations. Amazing quality!",
            f"Fast shipping, great designs, and excellent quality!",
            f"I can't stop wearing my {store_name} hoodie! So comfortable!"
        ]

        customer_testimonial = random.choice(testimonials)  # Chọn 1 lời nhận xét ngẫu nhiên

        email_templates = {
             "InfoReceived": {
                "url_mail": r"C:\17track\file_html\InfoReceived.html",
            },
            "InTransit": {
                "url_mail": r"C:\17track\file_html\InTransit.html",
            },
            "PickUp": {
               "url_mail": r"C:\17track\file_html\PickUp.html",
            },
            "OutForDelivery": {
                "url_mail": r"C:\17track\file_html\OutForDelivery.html",
            },
            "Undelivered": {
                "url_mail": r"C:\17track\file_html\Undelivered.html",
            },
            "Delivered": {
                "url_mail": r"C:\17track\file_html\Delivered.html",
            },
            "Alert": {
                "url_mail": r"C:\17track\file_html\Alert.html",
            },
            "Expired": {
                "url_mail": r"C:\17track\file_html\Expired.html",
            },

            # 📌 EMAIL NGÀY 1-5
            "day1": {
                "url_mail": r"C:\17track\file_html\day1.html",
            },
            "day2": {
                "url_mail": r"C:\17track\file_html\day2.html",
            },
            "day3": {
                "url_mail": r"C:\17track\file_html\day3.html",
            },
            "day4": {
                "url_mail": r"C:\17track\file_html\day4.html",
            },
            "day5": {
                "url_mail": r"C:\17track\file_html\day5.html",
            },

            # 📌 EMAIL FAILED NGÀY 1-7
           "day1Failed": {
               "url_mail": r"C:\17track\file_html\day1Failed.html",
            },
            "day2Failed": {
                "url_mail": r"C:\17track\file_html\day2Failed.html",
            },
            "day3Failed": {
                "url_mail": r"C:\17track\file_html\day3Failed.html",
            },
            "day4Failed": {
                "url_mail": r"C:\17track\file_html\day4Failed.html",
            },
            "day5Failed": {
                "url_mail": r"C:\17track\file_html\day5Failed.html",
            },
            "day6Failed": {
                "url_mail": r"C:\17track\file_html\day6Failed.html",
            },
            "day7Failed": {
                "url_mail": r"C:\17track\file_html\day7Failed.html",
            },
            "marketing": {
                "url_mail": r"C:\17track\file_html\marketing.html",
            }
        }

        if status in email_templates:
            url_html_mail = email_templates[status]["url_mail"]
        else:
            print("❌ Trạng thái không hợp lệ, không có email template phù hợp.")
            return
        replacements = {
            "customer_name": customer_name,
            "store_name": store_name,
            "company_logo_URL": company_logo_URL,
            "email_supports": emailSupport,
            "current_year": "2025",
            "tracking_number": tracking_number,
            "checkout_url": checkout_url,
            "customer_testimonial": customer_testimonial,
            "expiration_date" : expiration_date.strftime("%Y-%m-%d %H:%M:%S"),
        }
        # Gửi email
        list_subject = ["InfoReceived", "InTransit","PickUp","OutForDelivery","Undelivered","Delivered","Alert","Expired"]
        subject = f"Update on Your Order - {status}" if status in list_subject else "Update on Your Order"

        self.send_email_gmail_api(emailSupport, email, subject, replacements, url_html_mail)
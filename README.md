# OrderSp

#### Luồng vận hành

### b1 sẽ lấy các thông tin order sản phẩm (code ở google scrip)

### b2 coppy data từ sheet1 sang sheet2 của từng cửa hàng ( chỉ lấy những trường cần thiết) dối với những sheet mới dùng code coppy sheet ở google scrip

### dối với những cửa hàng đã có data sheet 2 thì cập nhật theo bước sau

### ----- 2.1 cập nhật sheet2

### ----- + kiểm tra xem order id của sheet1 đã có trong sheet2 hay chưa

### ------------ nếu đã có thì kiểm tra number checking xem trong sheet2 có khác sheet1 không , khác nhau thì cập nhật lại value number checking vào sheet2

### ------------ nếu chưa có thêm dòng mới vào sheet1

### ----- + Sắp xếp lại sheet2 theo cột order ID

### ----- 2.2 kiểm tra các điều kiện của sheet2 để gửi mail

### ----- + kiểm tra đơn hàng có status là failed không

### -----------Nếu có : kiểm tra xem nếu không phải từ IL hoặc FL đặt hàng sẽ gửi mail kêu gọi khách hàng quy lại với mua hàng ( gửi 7mail từ ngày 1 đến ngày 7, nếu khách đã mua hàng rồi thì thôi)

### -----------Nếu không failed : gửi mail cskh từ ngày 1-5 , nếu checking number của khách đã có trạng thái thì gửi theo trạng thái luôn

### ----- 2.3 gửi lại mail với những number checking khi kiểm tra checking bị lỗi

### -------------- ở file main_update kiểm tra tất cả các hàng của sheet2 cột E có value mà cột J hoặc cột L(cột gửi mail cho mình) chưa có value thì sẽ kiểm tra lại stauts để gửi mail

### ------------file main_update_order : cập nhật order của tất cả các cửa hàng

#### ------------- th1 nếu order id chưa có thong sheet : thêm hàng dữ liệu đó vào sheet

#### ------------- th2 nếu đã có order id thì cập nhật lại 2 trường order status + number checking của các order id đó ( phần này có thể cải thiện là nếu order id có sự khác biệt về number checking hoặc order status thì mới cập nhât -> giảm số lượng request lên google sheet nhưng tăng số phép tính thực hiện trong code) hiện tại đang cập nhật lại toàn bộ chứ k só sánh

### ------------- file main_add_checking_order : kiểm tra tất cả các hàng ở sheet1

##### -------------------có order_id and number_checking kiểm tra xem order_id đó đã nhập checking trong web chưa - nếu đã nhập rồi thì bỏ qua nếu chưa thì sẽ call api để nhập number checking và order status vào trong web

#### luồng chạy : chạy file main_update_order để cập nhật tất cả các order(cập nhật chứ không tạo mới)

#### chạy file main để gửi mail tự động cho khách hàng (hàm này cập nhật cả sheet2 để gửi mail)

#### chạy file main_update để kiểm tra lại các number checking để gửi lại mail khi file main kiểm tra number checking bị sót

#### chạy file main_add_checking_order để cập nhật number checking và status order trong web admin

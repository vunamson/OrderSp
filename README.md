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

### -------------- ở file main_update kiểm tra tất cả các hàng của sheet2 cột E có value mà cột J chưa có value thì sẽ kiểm tra lại stauts để gửi mail

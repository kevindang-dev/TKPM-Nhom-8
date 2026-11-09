Feature: Tối ưu hóa thể tích logistics bằng nén chân không và tích hợp dịch vụ đóng gói quà tặng (NV12)
  In order to chủ động giảm chi phí vận chuyển gấu bông cồng kềnh và chuẩn bị trọn gói quà tặng gửi đến người nhận
  As a Khách hàng mua quà tặng thú bông trực tuyến
  I want to hệ thống tái cấu trúc giảm 50% thể tích đóng gói khi hút chân không và tổng hợp chính xác phụ phí các dịch vụ đi kèm

  Background:
    Given công thức quy đổi thể tích bưu chính quốc tế bằng chiều dài nhân chiều rộng nhân chiều cao chia cho 5000
    And dịch vụ hút chân không giảm 50% thể tích đóng gói với phụ phí cố định 15.000 VNĐ
    And phụ phí gói quà sinh nhật là 20.000 VNĐ và viết thiệp chúc mừng là 10.000 VNĐ

  Rule: Rule 1: Tối ưu chi phí qua Dịch vụ hút chân không (Tái cấu trúc hình học kiện hàng)
    Example: Kiện thú bông cồng kềnh giảm một nửa thể tích khi chọn hút chân không
      Given kiện hàng có cân nặng thực tế 2.0 kg và kích thước 100.0 cm x 60.0 cm x 50.0 cm
      When hệ thống tính cước có áp dụng dịch vụ hút chân không
      Then khối lượng thể tích ban đầu là 60.0 kg
      And khối lượng thể tích sau hút chân không là 30.0 kg
      And trọng lượng tính cước sau hút chân không là 30.0 kg
      And phụ phí hút chân không là 15.000 VNĐ

  Rule: Rule 2: Kết hợp đa dịch vụ gia tăng quà tặng và lưu vết hệ thống
    Example: Khách hàng chọn kết hợp hút chân không, gói quà sinh nhật và viết thiệp chúc mừng
      Given kiện hàng có cân nặng thực tế 2.0 kg và kích thước 100.0 cm x 60.0 cm x 50.0 cm
      And khách hàng chọn các dịch vụ gia tăng "HUT_CHAN_KHONG, GOI_QUA, VIET_THIEP"
      When hệ thống tính cước và lưu kết quả vào hệ thống
      Then phụ phí hút chân không là 15.000 VNĐ
      And phụ phí gói quà là 20.000 VNĐ
      And phụ phí thiệp chúc mừng là 10.000 VNĐ
      And tổng phụ phí là 45.000 VNĐ
      And hệ thống tạo mã phiếu tính cước hợp lệ và lưu trữ kết quả thành công

  Rule: Rule 3: Ma trận kết hợp đa dịch vụ đóng gói và tính tổng chi phí thanh toán
    Scenario Outline: Tính toán tổng tiền cước khi khách hàng kết hợp các dịch vụ gia tăng khác nhau
      Given kiện hàng có cân nặng thực tế <can_nang_thuc> kg và kích thước <dai> cm x <rong> cm x <cao> cm
      And dịch vụ gia tăng được chọn là "<dich_vu>"
      When hệ thống thực hiện tính toán cước phí vận chuyển
      Then trọng lượng tính cước là <trong_luong_tinh_cuoc> kg
      And cước cơ bản là <cuoc_co_ban> VNĐ
      And tổng phụ phí là <tong_phu_phi> VNĐ
      And tổng cước thanh toán là <tong_cuoc> VNĐ

      Examples:
        | can_nang_thuc | dai   | rong | cao  | dich_vu                             | trong_luong_tinh_cuoc | cuoc_co_ban | tong_phu_phi | tong_cuoc |
        | 2.0           | 30.0  | 20.0 | 15.0 | GOI_QUA                             | 2.0                   | 30.000      | 20.000       | 50.000    |
        | 2.0           | 100.0 | 60.0 | 50.0 | HUT_CHAN_KHONG                      | 30.0                  | 310.000     | 15.000       | 325.000   |
        | 2.0           | 100.0 | 60.0 | 50.0 | HUT_CHAN_KHONG, GOI_QUA             | 30.0                  | 310.000     | 35.000       | 345.000   |
        | 2.0           | 100.0 | 60.0 | 50.0 | HUT_CHAN_KHONG, GOI_QUA, VIET_THIEP | 30.0                  | 310.000     | 45.000       | 355.000   |

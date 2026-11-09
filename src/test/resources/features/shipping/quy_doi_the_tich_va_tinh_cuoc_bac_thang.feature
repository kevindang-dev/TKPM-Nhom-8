Feature: Quy đổi thể tích bưu chính và định giá cước vận chuyển bậc thang cho kiện hàng cồng kềnh (NV11)
  In order to giải quyết điểm đau phí vận chuyển gấu bông cồng kềnh quá cao và định giá cước giao hàng minh bạch, chuẩn xác
  As a Khách hàng mua thú bông trực tuyến
  I want to hệ thống tự động quy đổi kích thước hình học 3 chiều theo thể tích bưu chính quốc tế và áp dụng biểu phí lũy tiến bậc thang

  Background:
    Given biểu phí vận chuyển tiêu chuẩn áp dụng nấc cơ bản 20.000 VNĐ cho 1.0 kg đầu tiên
    And nấc lũy tiến là 5.000 VNĐ cho mỗi 0.5 kg tiếp theo sau 1.0 kg đầu
    And công thức quy đổi thể tích bưu chính quốc tế bằng chiều dài nhân chiều rộng nhân chiều cao chia cho 5000

  Rule: Rule 1: Trọng lượng tính cước theo chuẩn bưu chính quốc tế
    Example: Kiện hàng nhỏ gọn tính theo cân nặng thực tế khi thể tích quy đổi nhỏ hơn
      Given kiện hàng có cân nặng thực tế 1.5 kg và kích thước 20.0 cm x 15.0 cm x 10.0 cm
      When hệ thống tính cước vận chuyển không áp dụng hút chân không
      Then khối lượng thể tích là 0.6 kg
      And trọng lượng tính cước là 1.5 kg

    Example: Thú bông khổng lồ cồng kềnh tính theo khối lượng thể tích khi thể tích lớn hơn
      Given kiện hàng có cân nặng thực tế 2.0 kg và kích thước 100.0 cm x 60.0 cm x 50.0 cm
      When hệ thống tính cước vận chuyển không áp dụng hút chân không
      Then khối lượng thể tích là 60.0 kg
      And trọng lượng tính cước là 60.0 kg

  Rule: Rule 2: Biểu phí cước vận chuyển bậc thang
    Scenario Outline: Tính toán cước cơ bản bậc thang và tổng cước thanh toán cho nhiều kiện hàng
      Given kiện hàng có cân nặng thực tế <can_nang_thuc> kg và kích thước <dai> cm x <rong> cm x <cao> cm
      And dịch vụ gia tăng được chọn là "<dich_vu>"
      When hệ thống thực hiện tính toán cước phí vận chuyển
      Then trọng lượng tính cước là <trong_luong_tinh_cuoc> kg
      And cước cơ bản là <cuoc_co_ban> VNĐ
      And tổng phụ phí là <tong_phu_phi> VNĐ
      And tổng cước thanh toán là <tong_cuoc> VNĐ

      Examples:
        | can_nang_thuc | dai   | rong  | cao  | dich_vu | trong_luong_tinh_cuoc | cuoc_co_ban | tong_phu_phi | tong_cuoc |
        | 0.8           | 15.0  | 10.0  | 10.0 | KHONG   | 0.8                   | 20.000      | 0            | 20.000    |
        | 1.5           | 20.0  | 15.0  | 10.0 | KHONG   | 1.5                   | 25.000      | 0            | 25.000    |
        | 2.0           | 100.0 | 60.0  | 50.0 | KHONG   | 60.0                  | 610.000     | 0            | 610.000   |

  Rule: Rule 3: Kiểm tra tính hợp lệ dữ liệu đầu vào và các trường hợp biên
    Scenario Outline: Từ chối tính cước khi kích thước kiện hàng không hợp lệ
      Given kiện hàng có thông tin kích thước <dai> cm x <rong> cm x <cao> cm và cân nặng 2.0 kg
      When hệ thống kiểm tra và tính cước kiện hàng
      Then hệ thống từ chối tính cước với thông báo lỗi "<thong_bao_loi>"

      Examples:
        | dai   | rong  | cao  | thong_bao_loi                      |
        | 0.0   | 50.0  | 40.0 | Kích thước kiện hàng phải lớn hơn 0 |
        | -10.0 | 50.0  | 40.0 | Kích thước kiện hàng phải lớn hơn 0 |
        | 50.0  | 0.0   | 40.0 | Kích thước kiện hàng phải lớn hơn 0 |
        | 50.0  | -5.0  | 40.0 | Kích thước kiện hàng phải lớn hơn 0 |
        | 50.0  | 40.0  | 0.0  | Kích thước kiện hàng phải lớn hơn 0 |
        | 50.0  | 40.0  | -2.0 | Kích thước kiện hàng phải lớn hơn 0 |

    Scenario Outline: Từ chối tính cước khi cân nặng thực tế không hợp lệ
      Given kiện hàng có thông tin kích thước 50.0 cm x 40.0 cm x 30.0 cm và cân nặng <can_nang> kg
      When hệ thống kiểm tra và tính cước kiện hàng
      Then hệ thống từ chối tính cước với thông báo lỗi "Cân nặng thực tế phải lớn hơn 0"

      Examples:
        | can_nang |
        | 0.0      |
        | -1.5     |

    Scenario Outline: Gắn cờ cảnh báo khi kiện hàng vượt trần kích thước vận chuyển tiêu chuẩn
      Given kiện hàng có kích thước <dai> cm x <rong> cm x <cao> cm và cân nặng <can_nang> kg
      When hệ thống thực hiện tính toán cước phí vận chuyển
      Then hệ thống hiển thị cảnh báo "Vượt quá kích thước tối đa của vận chuyển tiêu chuẩn, chuyển sang đường vận chuyển đặc biệt"

      Examples:
        | dai   | rong  | cao   | can_nang | ghi_chu                                                |
        | 210.0 | 50.0  | 40.0  | 5.0      | Chiều dài 210 cm vượt quá trần 200 cm                 |
        | 150.0 | 120.0 | 100.0 | 5.0      | Khối lượng quy đổi 360 kg vượt quá trần 100 kg         |

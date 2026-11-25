# Security Policy

## Supported Versions

Hiện tại chúng tôi hỗ trợ cập nhật bảo mật cho phiên bản mới nhất:

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |

## Reporting a Vulnerability

Nếu bạn phát hiện lỗ hổng bảo mật trong dự án này (ví dụ: rò rỉ API Key, lỗi xử lý XML/PDF độc hại, hoặc lỗ hổng từ thư viện phụ thuộc), vui lòng **không** tạo Issue công khai.

Hãy gửi email trực tiếp về địa chỉ: `lequoccuong1412004@gmail.com` hoặc sử dụng tính năng **Report a vulnerability** (GitHub Private Vulnerability Reporting) trong tab **Security** của repository.

Chúng tôi sẽ xác nhận và phản hồi trong vòng **48 giờ**.

## Security Best Practices cho người dùng

- **Không bao giờ** commit file `.env` hoặc `GEMINI_API_KEY` lên GitHub (file `.env` đã được chặn mặc định trong `.gitignore`).
- Hạn chế tải lên hóa đơn chứa thông tin nhạy cảm lên các môi trường chia sẻ công cộng khi chưa bật xác thực.
- Thường xuyên cập nhật các thư viện trong `requirements.txt` theo cảnh báo từ **Dependabot**.

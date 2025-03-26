# Convert PDF file or image to excel using Gemini API
📌 **Tổng quan**

Ứng dụng này cho phép chuyển đổi tài liệu PDF/hình ảnh sang file Excel tự động bằng cách sử dụng Gemini API của Google. Người dùng có thể tương tác qua giao diện đồ họa đơn giản. **Phần mềm này được lập trình bằng AI**.

![Demo Ứng dụng](https://github.com/user-attachments/assets/421dfb16-b899-4398-8c28-edad67acea39) 

🚀 **Tính năng chính**

- ✅ Chọn file PDF hoặc ảnh đầu vào

- ✅ Chỉ định thư mục lưu file Excel

- ✅ Tương tác với Gemini API qua giao diện

- ✅ Tự động sinh code Python từ AI

- ✅ Xem trước và chỉnh sửa code

- ✅ Thực thi code trực tiếp trong ứng dụng

- ✅ Tự động mở file Excel sau khi tạo

⚙️ **Cài đặt**

  Yêu cầu hệ thống
  
    Python 3.7+
    
    Hệ điều hành: Windows/macOS/Linux
  
  Cài đặt thư viện
  
    pip install -r requirements.txt
  
  Nội dung file requirements.txt:
  
    tkinter
    openpyxl>=3.1.2
    requests>=2.31.0
    python-dotenv>=1.0.0
    Pillow>=10.0.0
  🔑 **Cấu hình API**
  
  1. Lấy API Key từ Google AI Studio
  
  2. Nhập API Key vào ô tương ứng trong ứng dụng
  
  3. API Key sẽ được lưu tự động ở:
  
    Windows: C:\Users$$Tên_người_dùng]\.excel_converter\config.json
  
    macOS/Linux: ~/.excel_converter/config.json

🖥️ **Cách sử dụng**

1. Khởi chạy ứng dụng:

       python gemini_excel_converter.py
  
2. Thao tác với giao diện:

  - Nhập API Key của bạn

  - Chọn file PDF/ảnh cần xử lý

  - Chọn thư mục lưu file Excel

  - Nhập yêu cầu xử lý (ví dụ: "Giữ nguyên định dạng bảng")

  - Nhấn "Chạy Prompt" để sinh code

  - Xem và kiểm tra code

  - Nhấn "Chạy Code" để tạo file Excel

  - File kết quả sẽ được lưu tại:


        [Thư_mục_đã_chọn]/[Tên_file_gốc].xlsx

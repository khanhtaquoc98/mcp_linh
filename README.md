# Office Automation MCP Server (`Office-Automation-MCP`)

Dự án Local MCP Server viết bằng Python (sử dụng `FastMCP`) giúp trợ lý AI tự động hóa các tác vụ văn phòng trên **Microsoft Word (`.docx`)**, **Microsoft Excel (`.xlsx`)**, **Google Sheets**, và **Google Docs**.

Thiết kế hoàn toàn dùng đường dẫn tương đối (`os.path.dirname(os.path.abspath(__file__))`), dễ dàng `git clone` sang bất kỳ máy tính nào (Windows, macOS, Linux) mà không gặp lỗi đường dẫn cứng.

---

## Danh sách các công cụ (MCP Tools)

1. **`read_word_file(file_path)`**: Đọc toàn bộ văn bản và bảng biểu (ngăn cách bởi dấu `|`) từ file `.docx`.
2. **`create_word_file(file_path, title, paragraphs, table_data=None)`**: Tạo file `.docx` mới với tiêu đề, đoạn văn và bảng biểu.
3. **`read_excel_sheet(file_path, sheet_name=None, max_rows=200)`**: Đọc dữ liệu từ file `.xlsx`.
4. **`write_excel_data(file_path, data, sheet_name="Sheet1")`**: Ghi dữ liệu mảng 2 chiều vào file Excel mới hoặc nối tiếp vào file có sẵn.
5. **`fill_excel_column(file_path, values, column_letter="B", start_row=2, sheet_name=None)`**: Điền danh sách giá trị dọc theo một cột cố định từ dòng chỉ định, bảo toàn định dạng và công thức các ô khác của file mẫu.
6. **`read_google_sheet(url_or_id, worksheet_name=None)`**: Tự động tách ID từ URL Google Sheets và đọc dữ liệu qua `gspread`.
7. **`append_google_sheet(url_or_id, rows_data, worksheet_name=None)`**: Thêm các dòng mới vào cuối Google Sheet với chế độ `USER_ENTERED`.
8. **`read_google_doc(url_or_id)`** & **`append_google_doc(url_or_id, text_to_append)`**: Tự động tách ID từ URL Google Docs để đọc nội dung và chèn thêm văn bản vào cuối tài liệu.

---

## Hướng dẫn cài đặt từng bước

### Bước 1: Clone dự án về máy
```bash
git clone git@github.com:khanhtaquoc98/mcp_linh.git
cd mcp_linh
```

### Bước 2: Khởi tạo môi trường ảo `venv` và cài đặt thư viện
- **Trên macOS / Linux:**
  ```bash
  python3 -m venv venv
  source venv/bin/activate
  pip install -r requirements.txt
  ```
- **Trên Windows (PowerShell / CMD):**
  ```powershell
  python -m venv venv
  .\venv\Scripts\activate
  pip install -r requirements.txt
  ```

### Bước 3: Thiết lập Google Service Account (Chỉ cần khi dùng Google Sheets / Docs)
1. Truy cập [Google Cloud Console](https://console.cloud.google.com/), tạo Project và bật **Google Sheets API**, **Google Docs API**, **Google Drive API**.
2. Vào **IAM & Admin > Service Accounts**, tạo một Service Account mới và tải khóa JSON về máy.
3. Đổi tên file JSON vừa tải thành `service_account.json` và đặt vào **cùng thư mục** với `office_server.py`.
4. Mở file Google Sheets hoặc Google Docs bạn muốn tự động hóa, bấm **Share (Chia sẻ)** và cấp quyền **Editor (Người chỉnh sửa)** cho địa chỉ email (`client_email`) nằm trong file `service_account.json`.

### Bước 4: Cấu hình vào Antigravity / Claude Desktop
Xem file mẫu `mcp_config.example.json` và dán cấu hình tương ứng vào phần `mcpServers` của Antigravity:

- **Cách 1: Dùng `venv` truyền thống (macOS/Linux)**
  ```json
  {
    "mcpServers": {
      "office-automation": {
        "command": "/đường/dẫn/tới/mcp_linh/venv/bin/python",
        "args": ["/đường/dẫn/tới/mcp_linh/office_server.py"]
      }
    }
  }
  ```

- **Cách 2: Dùng `uv run` (Tiện lợi khi di chuyển giữa các máy)**
  ```json
  {
    "mcpServers": {
      "office-automation": {
        "command": "uv",
        "args": [
          "run",
          "--with-requirements",
          "/đường/dẫn/tới/mcp_linh/requirements.txt",
          "python",
          "/đường/dẫn/tới/mcp_linh/office_server.py"
        ]
      }
    }
  }
  ```

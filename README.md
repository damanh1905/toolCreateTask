# Tool Import Task/Sub-task lên Jira từ Excel

## Lệnh chạy nhanh

### 1. Kiểm tra Python

```powershell
python --version
pip --version
```

Nếu chưa có Python, tải và cài tại: https://www.python.org/downloads/

Khi cài nhớ tick **Add python.exe to PATH**, sau đó mở lại PowerShell.

### 2. Mở thư mục dự án

```powershell
cd C:\Users\Admin\Documents\vnpt\toolCreateTask
```

### 3. Cài thư viện

```powershell
pip install -r requirements.txt
```

### 4. Tạo file cấu hình `.env`

```powershell
notepad .env
```

Dán cấu hình mẫu:

```env
JIRA_URL=https://dvs.vnptmedia.vn
JIRA_TOKEN=<token_cua_ban>
JIRA_USERNAME=<username_jira>
JIRA_COOKIE=<cookie_neu_can>
JIRA_VERIFY_SSL=true

JIRA_DEFAULT_PROJECT_KEY=MYTVB2C
JIRA_DEFAULT_ISSUE_TYPE=Task
JIRA_SUBTASK_ISSUE_TYPE=Sub-Task
JIRA_FIELD_TARGET_START=customfield_11002
JIRA_FIELD_TARGET_END=customfield_11003
JIRA_FIELD_NOTE=customfield_12200
JIRA_REFERENCE_ISSUE=MYTVB2C-51653
```

### 5. Test kết nối Jira

```powershell
python main.py test
```

Test với issue cụ thể:

```powershell
python main.py test -i MYTVB2C-51653
```

### 6. Tạo file Excel mẫu

```powershell
python main.py template -o file_mau.xlsx
```

Các cột trong Excel:

```text
Mã | Loại | Task cha | Tiêu đề | Người xử lý | Mô tả | Ghi chú | Ngày bắt đầu | Ngày kết thúc | Deadline
```

`Người xử lý` để trống sẽ tự lấy `JIRA_USERNAME` trong `.env`.

### 7. Xem cây task trong sprint

```powershell
python main.py sprint-tree -k MYTVB2C-51653
```

Tìm task theo tiêu đề:

```powershell
python main.py sprint-tree -k MYTVB2C-51653 --lookup "tên task cần tìm"
```

### 8. Kiểm tra file Excel

```powershell
python main.py validate file_mau.xlsx
```

### 9. Chạy thử, không tạo issue thật

```powershell
python main.py import file_mau.xlsx --dry-run
```

### 10. Import thật lên Jira

```powershell
python main.py import file_mau.xlsx
```

### 11. Một số lệnh thêm

Chỉ định issue tham chiếu sprint khác:

```powershell
python main.py import file_mau.xlsx -k MYTVB2C-50784
```

Bỏ qua dòng lỗi và chạy tiếp:

```powershell
python main.py import file_mau.xlsx --continue-on-error
```

Tạo file mẫu không có dòng ví dụ:

```powershell
python main.py template -o file_mau.xlsx --no-sample
```

## Thứ tự chạy thường dùng

```powershell
cd C:\Users\Admin\Documents\vnpt\toolCreateTask
pip install -r requirements.txt
python main.py test
python main.py template -o file_mau.xlsx 
python main.py validate file_mau.xlsx
python main.py import file_mau.xlsx --dry-run
python main.py import file_mau.xlsx
```

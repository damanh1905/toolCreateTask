# Tool Import Task/Sub-task lên Jira từ Excel

Tool import Excel lên Jira, logic bám theo hàm `create_and_update_issue` trong dự án.

## Cài đặt

```bash
pip install -r requirements.txt
copy .env.example .env
```

## Cấu hình `.env`

```env
JIRA_URL=https://dvs.vnptmedia.vn
JIRA_TOKEN=<token>
JIRA_USERNAME=<username>

JIRA_DEFAULT_PROJECT_KEY=MYTVB2C
JIRA_FIELD_TARGET_START=customfield_11002
JIRA_FIELD_TARGET_END=customfield_11003
JIRA_FIELD_NOTE=customfield_xxxxx
JIRA_REFERENCE_ISSUE=MYTVB2C-50784
```

## Tạo file mẫu

```bash
python main.py template -o file_mau_don_gian.xlsx
```

## Các cột Excel

| Cột Excel | Field trong dự án | Gửi lên Jira |
|-----------|-------------------|--------------|
| `Loại` | `issuetype` / `is_task` | `Task` hoặc `Sub-task` |
| `Task cha` | `parent_key` | Bắt buộc với Sub-task |
| `Tiêu đề` | `title` / `summary` | Bắt buộc |
| `Người xử lý` | `assignee` | Bắt buộc |
| `Mô tả` | `description` | Có |
| `Ghi chú` | `note` | Update `JIRA_FIELD_NOTE` sau khi tạo |
| `Ngày bắt đầu` | `start_date` | `JIRA_FIELD_TARGET_START` |
| `Ngày kết thúc` | `end_date` | `JIRA_FIELD_TARGET_END` |
| `Deadline` | `deadline` | `duedate` |
| `Mã` | — | Mã tạm nội bộ, liên kết subtask trong file |

Các cột `Người tạo`, `Ngày tạo`, `Ngày update`, `Trạng thái`, `Thao tác` chỉ để theo dõi, không gửi lên Jira.

## Logic giống dự án

- **Task**: `create_task` — labels rỗng, có start/end
- **Sub-task**: `create_subtask` — cần `Task cha`
- Sau khi tạo: update `note` → **add vào sprint** (`add_issue_to_sprint`) — bắt buộc để web lấy được
- Nếu add sprint thất bại → xóa issue vừa tạo (giống dự án)

## Chạy

```bash
các lệnh kiểm tra:
python main.py test
Kết nối Jira

python main.py validate file_mau.xlsx
File Excel hợp lệ

python main.py import file_mau.xlsx --dry-run
Xem trước, không tạo

python main.py sprint-tree --issue-key ...
Xem task trong sprint

python main.py import file_mau.xlsx
```
# Tự động cập nhật dashboard BQLDT

Bộ file này dùng cho repo GitHub Pages `daklak-tech/bqldt`.

## Cách dùng

1. Chép toàn bộ các file/thư mục trong thư mục này vào repo `daklak-tech/bqldt`.
2. Bảo đảm repo có file `index.html` ở thư mục gốc.
3. Vào GitHub repo > Settings > Actions > General, bật quyền **Read and write permissions** cho workflow.
4. Workflow `.github/workflows/daily-dashboard-update.yml` sẽ chạy hằng ngày lúc 18:00 giờ Việt Nam.
5. Có thể chạy thử bằng nút **Run workflow** trong tab Actions.

## Nguồn dữ liệu

Script đang lấy từ Google Sheet:
https://docs.google.com/spreadsheets/d/1vpDSRjzNjJL3gNsAOdIxFh0Aad16yoD4/edit?gid=631738297#gid=631738297

Google Sheet cần cho phép tải/xuất file Excel để GitHub Actions đọc được.

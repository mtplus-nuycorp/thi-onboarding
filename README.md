# Tool thi Onboarding TP/TN

Web thi trắc nghiệm và tự luận. Mỗi người nhận một mã đề riêng, câu hỏi và đáp án được trộn ngẫu nhiên. Nộp bài xong chỉ hiện điểm trắc nghiệm, không hiện đáp án. Tự luận chấm chéo bằng tay trong Sheet.

- Link phần 1–3: `.../thi-onboarding/?phan=P1,P2,P3`
- Link phần 4–5: `.../thi-onboarding/?phan=P4,P5`

- `index.html`: trang thi, gửi link này cho anh em.
- `admin.html`: trang quản lý (bảng điểm, xếp hạng, câu sai nhiều nhất, ai sai cái gì, ai chưa thi). Cần mật khẩu.
- `config.js`: file duy nhất cần sửa (link API, tên chương trình).
- `apps-script/Code.gs`: phần chấm bài, dán vào Google Sheet. Không cần đưa file này lên GitHub.

## Cài đặt lần đầu (khoảng 15 phút)

### Bước 1. Gắn phần chấm bài vào Google Sheet
1. Mở Google Sheet "THI ONBOARDING TP/TN - Ngân hàng câu hỏi & Kết quả".
2. Menu **Tiện ích mở rộng → Apps Script**.
3. Xoá hết nội dung có sẵn trong `Code.gs`, dán toàn bộ nội dung file `apps-script/Code.gs`, bấm **Lưu** (biểu tượng đĩa).
4. Bấm **Triển khai → Tùy chọn triển khai mới**. Ở "Chọn loại", chọn **Ứng dụng web**.
   - Thực thi với tư cách: **Tôi**
   - Người có quyền truy cập: **Bất kỳ ai**
5. Bấm **Triển khai**, cấp quyền khi Google hỏi (chọn tài khoản → Nâng cao → Đi tới dự án → Cho phép).
6. Sao chép **URL ứng dụng web** (dạng `https://script.google.com/macros/s/.../exec`).

### Bước 3. Đưa web lên GitHub Pages
1. Mở `config.js`, thay `DAN_LINK_APPS_SCRIPT_VAO_DAY` bằng URL ở Bước 1. Sửa tên chương trình nếu muốn.
2. Trên GitHub bấm **New repository**, đặt tên (vd `thi-onboarding`), chọn **Public**, bấm **Create repository**.
3. Bấm **uploading an existing file**, kéo thả 4 file: `index.html`, `admin.html`, `style.css`, `config.js`. Bấm **Commit changes**.
4. Vào **Settings → Pages**. Ở "Branch" chọn `main` và `/ (root)`, bấm **Save**.
5. Đợi 1–2 phút, link web là `https://<tên-github>.github.io/thi-onboarding/`.
   - Link thi: `.../thi-onboarding/`
   - Link quản lý: `.../thi-onboarding/admin.html`

Repo để Public vẫn an toàn: trong 4 file này không có đáp án hay hướng dẫn chấm.

## Vận hành hằng ngày (không cần mở Claude)

| Việc | Làm ở đâu |
|---|---|
| Thêm, sửa câu hỏi, đáp án | Tab `CAU_HOI`. Đáp án ghi chữ cái, nhiều đáp án cách nhau dấu phẩy (vd `B,D`). |
| Tạm ẩn một câu | Cột "Đang dùng" chọn `KHÔNG` |
| Bối cảnh doanh nghiệp | Tab `PHAN_THI` cột I (mỗi dòng 1 bối cảnh). Câu tự luận nào dùng bối cảnh thì cột M tab `CAU_HOI` = `CÓ` |
| Đổi thời gian, số câu, điểm đạt | Tab `PHAN_THI` |
| Mở/đóng phần thi | Tab `PHAN_THI`, cột "Mở thi" |
| Thêm người thi | Tab `NHAN_VIEN` (Mã NV + Tên, "Đang hoạt động" = `CÓ`) |
| Chấm chéo, sửa điểm tự luận | Tab `CHI_TIET`, cột "Điểm chấm chéo", rồi menu **Tool thi → Tính lại điểm** |

Sửa Sheet xong là web cập nhật ngay, không cần triển khai lại.
Chỉ khi sửa `Code.gs` mới cần: **Triển khai → Quản lý các bản triển khai → Sửa (bút chì) → Phiên bản mới → Triển khai**. URL giữ nguyên.

## Cách tính điểm
- Trắc nghiệm: chọn đủ và đúng tất cả đáp án mới được điểm câu đó. Điểm TN = số câu đúng / tổng câu × 10.
- Tự luận: người chấm chéo điền điểm 0–10 vào cột "Điểm chấm chéo" tab `CHI_TIET`. Điểm TL = trung bình các câu.
- Tổng = TN × tỷ trọng + TL × (1 − tỷ trọng). Tỷ trọng đặt ở tab `PHAN_THI`.
- Tổng ≥ "Điểm đạt" là ĐẠT. Nộp quá giờ hơn 3 phút sẽ ghi chú "(nộp trễ)".

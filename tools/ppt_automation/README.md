# PPT Automation Starter

CLI khởi động nhanh để tự động tạo PowerPoint từ Word theo bố cục báo cáo cố định.

## Cài đặt

```bash
pip install python-docx python-pptx
```

## Chạy nhanh

```bash
python tools/ppt_automation/generate_ppt.py \
  --docx report.docx \
  --out output/report.pptx \
  --spec tools/ppt_automation/spec.example.json
```

## Cách hoạt động

- Parse các `Heading 1/2` trong Word thành tiêu đề slide.
- Parse paragraph/list thành bullet.
- Gán ảnh theo `slide_index` từ file spec JSON.
- Nhận animation timeline theo `slide_index` và lưu vào Notes để native runner (Office JS/COM) apply chính xác trong bước sau.

## Lưu ý quan trọng

Bản CLI này là **fallback renderer** (python-pptx). Nếu cần giống PowerPoint 100% cho animation/transition/phần tử nâng cao, dùng native PowerPoint execution layer như đã mô tả trong `TOOL_AUTO_PPT_FROM_WORD.md`.

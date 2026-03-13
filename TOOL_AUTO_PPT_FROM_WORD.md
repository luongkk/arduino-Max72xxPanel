# Thiết kế tool tạo PowerPoint tự động từ Word (template cố định ~80% nhưng vẫn linh hoạt)

## 1) Mục tiêu sản phẩm
- Tự động chuyển nội dung từ file Word có bố cục tương đối cố định (khoảng 80%) thành slide PowerPoint.
- Mục tiêu cao nhất: output hành vi và khả năng chỉnh sửa **giống PowerPoint 100%** ở file cuối.
- Vẫn giữ được năng lực xử lý “mạnh như PowerPoint”:
  - bố cục chuẩn,
  - định dạng phong phú,
  - animation/transition,
  - biểu đồ/bảng,
  - media,
  - chỉnh sửa thủ công hậu kỳ.
- Hạn chế thao tác tay lặp đi lặp lại khi làm báo cáo định kỳ.

## 2) Điều kiện để “giống PowerPoint 100%”

Muốn đạt mức 100%, nguyên tắc kỹ thuật là: **để chính PowerPoint render và lưu file cuối** (native engine), thay vì chỉ sinh OpenXML bằng thư viện bên thứ ba.

- Khuyến nghị kiến trúc 2 lớp:
  1. **Planning layer** (server): đọc Word, chuẩn hóa schema, lên kế hoạch bố cục/tài nguyên/animation.
  2. **Native PowerPoint execution layer** (client Windows hoặc Office Add-in): mở template, apply toàn bộ lệnh lên object model của PowerPoint, rồi `SaveAs .pptx`.
- Vì sao cần native engine:
  - đảm bảo fidelity cho animation nâng cao, morph/transition, media timing, smart art/chart behavior, group/ungroup, theme effect.
  - tránh sai khác hiển thị giữa các thư viện tạo file và ứng dụng PowerPoint thật.
- Chính sách triển khai:
  - **Gold path (100%)**: Office JS Add-in hoặc COM/VSTO automation.
  - **Fallback path (không cam kết 100%)**: `python-pptx`/OpenXML khi môi trường không có PowerPoint.

## 3) Kiến trúc đề xuất (Hybrid: Rule-based + AI)

### Luồng tổng quát
1. **Ingest Word**
   - Đọc `.docx` bằng parser có cấu trúc (thay vì OCR nếu có thể).
   - Trích xuất heading, paragraph, table, image, caption, danh sách.
2. **Chuẩn hóa dữ liệu**
   - Chuyển thành JSON trung gian (Document AST / Report Schema).
   - Gắn metadata: mức heading, vai trò đoạn (summary/detail), độ ưu tiên nội dung.
3. **Slide Planning Engine**
   - Rule cho 80% trường hợp chuẩn (mapping heading → layout).
   - AI chỉ xử lý 20% ngoại lệ: tóm tắt dài, gợi ý chia slide, viết lại bullet.
4. **Slide Rendering Engine**
   - Render vào `.pptx` từ template master (Slide Master + custom layouts).
   - Điền text, chart, table, hình ảnh, icon, notes.
5. **Post-process & QA**
   - Kiểm tra tràn chữ, overlap, font fallback, tỉ lệ hình.
   - Chấm điểm chất lượng và flag slide cần duyệt tay.
6. **Human-in-the-loop**
   - UI cho phép “accept / regenerate / edit” trước khi xuất bản.

## 4) Cách giữ “mạnh như PowerPoint” nhưng tự động hóa cao

### A. Chuẩn hóa thiết kế bằng template hệ thống
- Xây **Design System cho slide**:
  - token màu, font, spacing,
  - bộ layout chuẩn (Title, Section, Two-column, KPI, Timeline, Table-heavy…),
  - thư viện component (chart card, comparison block, risk matrix).
- Dùng duy nhất template `.potx`/`.pptx` chuẩn của công ty.
- Quy tắc: nội dung đổ vào component, không “vẽ tự do” khi auto-generate.

### B. Tách “nội dung” khỏi “trình bày”
- Word chỉ là nguồn nội dung.
- JSON trung gian mới là input chính cho engine render.
- Nhờ vậy dễ đổi template mà không đổi pipeline parser.

### C. Rule-first để ổn định, AI-second để linh hoạt
- Rule-based cho phần cố định:
  - `Heading 1` => trang section,
  - `Heading 2 + bullet` => content slide,
  - bảng > N cột => auto chia 2 slide hoặc dùng appendix.
- AI cho phần khó:
  - rút gọn đoạn dài thành bullet,
  - gợi ý tiêu đề ngắn,
  - phát hiện nội dung nên thành biểu đồ.

### D. Khả năng edit hậu kỳ như PowerPoint thật
- Xuất file `.pptx` chuẩn để đội ngũ vẫn mở bằng PowerPoint chỉnh tay.
- Gắn `Notes` và `Tags` để truy vết nội dung đã sinh tự động.
- Giữ naming convention cho shape để script có thể cập nhật lần sau (round-trip).

## 5) Định nghĩa schema trung gian (gợi ý)
```json
{
  "report_meta": {
    "title": "Q2 Business Review",
    "author": "Team A",
    "language": "vi"
  },
  "sections": [
    {
      "id": "sec_1",
      "title": "Tổng quan",
      "slides": [
        {
          "type": "kpi_summary",
          "title": "Kết quả chính",
          "bullets": ["..."],
          "kpis": [{"name": "Revenue", "value": "120B", "trend": "+12%"}],
          "assets": []
        }
      ]
    }
  ]
}
```

## 6) Công nghệ triển khai thực tế

### Parser & NLP
- `python-docx` hoặc `mammoth` để đọc Word.
- Optional: LLM API cho summarize/rewrite/classify.
- Vietnamese NLP: tách ý, chuẩn hóa thuật ngữ domain.

### Tạo PowerPoint
- Để đạt parity 100%: ưu tiên **Office JS Add-in** hoặc **COM/VSTO automation** để thao tác trực tiếp trên PowerPoint thật.
- `python-pptx` dùng cho môi trường server/fallback, nhưng không nên là đường chính nếu yêu cầu “giống 100%” cho animation/transition nâng cao.
- Cơ chế run-time đề xuất:
  - server tạo “execution plan” (JSON),
  - PowerPoint runner đọc plan và apply tuần tự trên object model,
  - PowerPoint native `SaveAs` file đầu ra.

### Backend service
- API kiểu:
  - `POST /ingest-word`
  - `POST /plan-slides`
  - `POST /render-pptx`
  - `POST /qa`
- Lưu bản nháp + versioning để regenerate theo từng section.

### Frontend/UX
- Giao diện duyệt từng slide:
  - compare “source text vs generated slide”,
  - nút Regenerate by style (ngắn hơn, formal hơn, data-first...).


## 7) Gán ảnh theo chỉ định (place image / fill shape như PowerPoint)

Để đáp ứng nhu cầu “tôi cung cấp ảnh + vị trí muốn gán”, nên thêm lớp **Asset Mapping Engine**:

- Input ảnh từ người dùng (upload hoặc URL) kèm chỉ định:
  - gán vào **slide cụ thể**,
  - gán vào **shape placeholder cụ thể** (`shape_name`),
  - hoặc gán kiểu **Fill shape** (Picture/Texture fill).
- Hỗ trợ nhiều chế độ fit giống PowerPoint:
  - `contain` (vừa khung, giữ đủ ảnh),
  - `cover` (lấp đầy khung, có thể crop),
  - `stretch` (co giãn),
  - `crop_focus` (crop theo trọng tâm/face/object).
- Thứ tự ưu tiên khi render:
  1) mapping theo `shape_id/shape_name`,
  2) mapping theo `role` (hero_image, product_image...),
  3) fallback vào vùng ảnh mặc định của layout.
- QA riêng cho ảnh:
  - cảnh báo ảnh quá nhỏ (mờ),
  - cảnh báo crop quá ngưỡng,
  - cảnh báo lệch tỉ lệ thương hiệu.

### Ví dụ chỉ định ảnh trong schema
```json
{
  "asset_bindings": [
    {
      "asset_id": "img_factory_01",
      "slide_ref": "sec_2_slide_1",
      "target": {
        "type": "shape_fill",
        "shape_name": "PHOTO_FRAME_MAIN",
        "fit": "cover",
        "crop_focus": "center"
      }
    },
    {
      "asset_id": "img_chart_bg",
      "slide_ref": "sec_2_slide_2",
      "target": {
        "type": "picture_placeholder",
        "shape_name": "RIGHT_IMAGE",
        "fit": "contain"
      }
    }
  ]
}
```

## 8) Chỉ định animation theo từng phần tử trong slide

Để animation “mạnh như PowerPoint”, nên tách thành **Animation Timeline Spec** độc lập với nội dung:

- Mỗi phần tử (title, bullet group, chart, image, icon...) có `object_ref` riêng.
- Khai báo cho từng object:
  - `effect` (fade, wipe, fly-in, zoom...),
  - `trigger` (on_click, with_previous, after_previous),
  - `duration`, `delay`, `order`,
  - `direction` nếu hiệu ứng hỗ trợ.
- Cho phép preset theo style báo cáo:
  - `formal_minimal`, `sales_pitch`, `training_step_by_step`.
- Rule an toàn:
  - không quá N animation/slide,
  - ưu tiên nhấn mạnh insight chính,
  - tự tắt animation khi xuất bản in/PDF.

### Ví dụ animation spec
```json
{
  "animations": [
    {
      "slide_ref": "sec_2_slide_1",
      "timeline": [
        {
          "object_ref": "TITLE",
          "effect": "fade",
          "trigger": "with_previous",
          "duration_ms": 400,
          "order": 1
        },
        {
          "object_ref": "BULLET_GROUP_1",
          "effect": "wipe",
          "direction": "from_left",
          "trigger": "on_click",
          "duration_ms": 350,
          "order": 2
        },
        {
          "object_ref": "PHOTO_FRAME_MAIN",
          "effect": "zoom",
          "trigger": "after_previous",
          "delay_ms": 120,
          "duration_ms": 450,
          "order": 3
        }
      ]
    }
  ]
}
```

## 9) Chiến lược xử lý 80/20
- **80% cố định**: pipeline deterministic để tốc độ nhanh và output nhất quán.
- **20% linh hoạt**:
  - fallback nhiều mức:
    1) layout mặc định,
    2) layout thay thế,
    3) đẩy vào appendix nếu quá tải.
- Đo chất lượng bằng KPI:
  - % slide không cần sửa,
  - thời gian chỉnh tay trung bình,
  - tỉ lệ lỗi tràn chữ/đè hình.

## 10) Lộ trình MVP (4–6 tuần)
1. Tuần 1: Chốt template + schema + quy ước Word input.
2. Tuần 2: Parser Word + mapping rule cơ bản.
3. Tuần 3: Render pptx + QA kiểm lỗi layout.
4. Tuần 4: Thêm AI summarize/title rewrite + UI duyệt.
5. Tuần 5-6: Pilot với dữ liệu thật, tối ưu rule, logging, metrics.

## 11) Nguyên tắc để scale lâu dài
- Template versioning (v1, v2...) và migration rule.
- Prompt/version control cho AI step.
- Telemetry chi tiết từng slide generation step.
- Human feedback loop để “học” pattern chỉnh tay phổ biến.
- Thiết lập **Parity Test Suite**: mở cùng file trên PowerPoint, export ảnh từng slide, so sánh pixel/layout/animation timeline với baseline.
- Chỉ đánh dấu “100% mode” khi chạy qua native PowerPoint execution layer và pass parity suite.

## 12) Trả lời ngắn cho câu hỏi của bạn
Nếu mục tiêu là **giống PowerPoint 100%**, bạn nên dùng kiến trúc:
- **Word -> JSON schema -> Rule planner -> Asset mapping -> Animation timeline -> Native PowerPoint runner -> QA parity -> Human review**.
- Điểm mấu chốt: file cuối phải do **PowerPoint native engine** render/lưu.
- AI vẫn dùng để tối ưu nội dung, nhưng phần “vẽ và hiệu ứng” giao cho PowerPoint object model để đảm bảo độ giống tuyệt đối.

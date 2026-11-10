# MR Office — Tiếng Việt

[English](../README.md) · [English (UK)](../i18n/README.en-GB.md) · [简体中文](../i18n/README.zh.md) · [繁體中文](../i18n/README.zh-TW.md) · [日本語](../i18n/README.ja.md) · [한국어](../i18n/README.ko.md) · [Deutsch](../i18n/README.de.md) · [Français](../i18n/README.fr.md) · [Español](../i18n/README.es.md) · [Español (Latinoamérica)](../i18n/README.es-419.md) · [Português (Brasil)](../i18n/README.pt-BR.md) · [Italiano](../i18n/README.it.md) · [Nederlands](../i18n/README.nl.md) · [Svenska](../i18n/README.sv.md) · [Dansk](../i18n/README.da.md) · [Norsk](../i18n/README.no.md) · [Suomi](../i18n/README.fi.md) · [Bahasa Indonesia](../i18n/README.id.md) · [ไทย](../i18n/README.th.md) · [Tiếng Việt](../i18n/README.vi.md)

> ### ⬇ Để dùng đầy đủ tính năng, hãy cài ứng dụng máy tính MR Office (miễn phí)
>
> Ứng dụng máy tính MR Office miễn phí, có các trình soạn thảo này cùng toàn bộ dịch vụ và chạy ngoại tuyến trên máy tính của bạn. Ứng dụng bổ sung những gì các trình soạn thảo mã nguồn mở còn thiếu:
>
> - Mở và lưu tệp Word, Excel và PowerPoint (.docx, .xlsx, .pptx)
> - Tính công thức có hàm (SUM, IF, VLOOKUP …)
> - Biểu đồ
> - Chỉnh sửa SmartArt
> - Ghi các chỉnh sửa PDF trở lại tệp
>
> **[Tải MR Office từ Releases →](https://github.com/dongmingyang2007-pixel/MROffice/releases)** · [hoặc từ mroffice.ai](https://mroffice.ai)

<p align="center"><a href="https://mroffice.ai"><b>🌐 mroffice.ai</b></a> · <a href="https://mroffice.ai">Dùng thử trong MR Office →</a></p>

> Các trình soạn thảo mã nguồn mở của MR Office — tài liệu, bảng tính, bản trình bày, PDF, viết tay và mã — chạy trong trình duyệt.

MR Office là không gian làm việc cho ghi chú, AI và tệp thực. Kho này công bố mã nguồn các trình soạn thảo của nó: chính những trình soạn thảo MR Office đang chạy, bạn có thể tự dựng, chạy và nhúng. Đây là các trình soạn thảo mã nguồn mở, không phải một bộ ứng dụng văn phòng mã nguồn mở hoàn chỉnh: việc đọc và ghi tệp Microsoft Office vẫn nằm trong MR Office (xem bên dưới).

## Các trình soạn thảo

Những gì mỗi trình soạn thảo làm được trong kho này:

- **Tài liệu** — trình xử lý văn bản theo trang: kiểu, phông chữ, danh sách, bảng, đầu trang và chân trang, phần và cột, hình dạng và phương trình, trang bìa và mục lục, chú thích và theo dõi thay đổi, cùng nét mực trên trang. Lưu ở định dạng riêng (.mrdoc); mở HTML, Markdown và văn bản thuần.
- **Bảng tính** — ô và công thức, định dạng số, định dạng có điều kiện, sắp xếp và lọc, cố định ngăn và nhiều trang tính, kèm tổng và trung bình của vùng chọn. Phép tính giữa các ô (=A1*B1+10) được tính trong trình duyệt khi gõ; công thức có hàm (SUM, IF, VLOOKUP …) được giữ nguyên như đã viết. Lưu ở định dạng riêng (.mrsheet); mở và lưu CSV, TSV.
- **Bản trình bày** — trang chiếu với tám chủ đề riêng, bố cục, dạng xem trang chiếu cái và bố cục, văn bản, hình dạng, hộp văn bản, bảng, WordArt và phương trình, chú thích, phần, chuyển tiếp và hoạt ảnh có xem trước, cùng trình chiếu. Lưu ở định dạng riêng (.mrslides).
- **PDF** — đọc, tìm kiếm, trình chiếu và in tệp PDF, sao chép và đọc to văn bản, và đánh dấu văn bản đã chọn (tô sáng, gạch chân, gạch ngang, gạch lượn sóng) kèm ghi chú và ngăn chú thích khi tệp đang mở; khi lưu sẽ trả lại tệp như lúc mở.
- **Viết tay** — ghi chú viết tay với bút cảm ứng lực nhấn, bút đánh dấu, tẩy, lasso và thước, và chuyển nét mực thành hình dạng. Lưu ở định dạng riêng (.mrink).
- **Mã** — trình soạn thảo mã gọn nhẹ với tô sáng cú pháp cho nhiều ngôn ngữ, dàn ý, tìm và thay thế; JavaScript/TypeScript, Python, SQL, Lua và shell chạy ngay trong trình duyệt.

Việc tính công thức có hàm, vẽ biểu đồ và dàn trang SmartArt do bộ máy riêng của MR Office thực hiện, trong ứng dụng máy tính MR Office: các trình soạn thảo mã nguồn mở giữ nguyên công thức, biểu đồ và SmartArt trong tệp, và hiển thị SmartArt như khi được lưu.

## Tệp Word, Excel, PowerPoint và PDF

Mở và lưu tệp Word, Excel và PowerPoint (.docx, .xlsx, .pptx), cũng như ghi các chỉnh sửa PDF trở lại tệp, do ứng dụng máy tính MR Office (miễn phí) cung cấp; ứng dụng có tất cả các trình soạn thảo này và chạy ngoại tuyến trên máy tính của bạn:

- **MR Office** (ứng dụng máy tính miễn phí): https://github.com/dongmingyang2007-pixel/MROffice/releases
- **Trang web**: https://mroffice.ai

[Tất cả phiên bản →](https://github.com/dongmingyang2007-pixel/MROffice/releases)

## Chạy các trình soạn thảo

Sáu trình soạn thảo nằm trong `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, dùng chung `@mrnote/core`) và chạy cùng nhau trong `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` khởi động máy chủ phát triển tại cùng địa chỉ, và `npm run smoke` kiểm tra một bản dựng. README của mỗi gói mô tả chi tiết gói làm được gì và cách nhúng.

## Giấy phép

Mã nguồn mở theo GNU Affero General Public License v3.0 (LICENSE), kèm các điều khoản bổ sung trong NOTICE.md: giữ ghi công "MR Office" và liên kết đến https://mroffice.ai, và không cấp quyền đối với tên hay biểu trưng MR Office. Giấy phép thương mại có tại support@mr-note.com. Phần mềm và dữ liệu của bên thứ ba được liệt kê trong THIRD_PARTY_NOTICES.md. Đóng góp: xem CONTRIBUTING.md.

## Thêm

Trang web và sản phẩm đầy đủ: **https://mroffice.ai** · Câu hỏi và góp ý: support@mr-note.com

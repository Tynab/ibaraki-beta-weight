# Ibaraki Beta Weight

Công cụ VB.NET WinForms/console hỗ trợ đội 西山 thuộc nhóm エマール nhập và chuyển dữ liệu nhanh hơn cho mẫu Excel 茨城 (ベタ) 重量 của đối tác 文化シャッター.

## Màn hình mẫu

<p align="center">
  <img src="pic/0.png" alt="Màn hình nhập liệu Ibaraki Beta Weight" />
</p>

## Tóm tắt mã nguồn

- `Script/Main.vb`: entrypoint của ứng dụng, cấu hình console UTF-8, kiểm tra license và khởi chạy workflow chính.
- `Script/Common.vb`: helper dùng chung cho update, nhập liệu console, quản lý Excel Interop, ghi/xóa ô Excel và định dạng console.
- `Script/Util.vb`: ánh xạ từng nhóm vật tư/thép vào đúng ô trong mẫu Excel Ibaraki Beta.
- `Script/Service.vb`: điều phối thứ tự câu hỏi nhập liệu và gọi các helper ghi Excel tương ứng.
- `Script/Constant.vb`: tập trung tên tiến trình, tên file MSI và đường dẫn lưu bộ cài update.
- `Control/FrmUpdate.vb`: form tải bộ cài mới, hiển thị tiến trình download và chạy MSI sau khi tải xong.
- `*.Designer.vb`, `*.resx`, `My Project/*.Designer.vb`: file do Visual Studio sinh tự động, không chỉnh tay khi cleanup/refactor.

## Luồng xử lý chính

1. Ứng dụng kiểm tra license trong user settings hoặc yêu cầu nhập serial.
2. Ứng dụng kiểm tra phiên bản mới và mở form update nếu server báo có bản khác.
3. Người dùng chọn workbook Excel cần xử lý.
4. Workflow console lần lượt hỏi từng nhóm vật tư/thép, chỉ ghi những số lượng lớn hơn 0 vào mẫu Excel.
5. Workbook được lưu, Excel Interop được đóng và file kết quả được mở lại cho người dùng.

## Ví dụ mã

```vb
''' <summary>
''' Nhập số lượng ngoại/nội chu vi GL-150 theo cấp G.
''' </summary>
''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
Friend Sub Unit150(xlApp As Application)
    PubSequentialInputs(xlApp, 18, GradePrompts)
End Sub
```

## Package

<img src="pic/1.png" align="left" width="3%" height="3%" alt="Package icon" />

- Microsoft.Office.Interop.Excel `15.0.4795.1001`

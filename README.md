# TOUHOKU (BETA) WEIGHT

Ứng dụng hỗ trợ đội 西山 thuộc エマールグループ nhập và chuyển dữ liệu nhanh hơn cho mẫu trọng lượng 茨城 (ベタ) từ đối tác 文化シャッター.

## Icon

<p align='center'>
<img src='pic/0.png'></img>
</p>

## Thành Phần Chính

| Thành phần | Tệp | Mô tả |
|---|---|---|
| `Main` | `Ibaraki Beta Weight/Script/Main.vb` | Điểm vào ứng dụng: thiết lập console UTF-8, kiểm tra serial bản quyền và gọi `RunApp`. |
| `Common` | `Ibaraki Beta Weight/Script/Common.vb` | Hàm dùng chung: kiểm tra cập nhật, thao tác Excel COM, định dạng console, hiệu ứng biểu mẫu, quản lý tiến trình và file. |
| `Constant` | `Ibaraki Beta Weight/Script/Constant.vb` | Hằng số toàn ứng dụng: tên tiến trình Excel, tên file cài đặt và đường dẫn trong AppData. |
| `Service` | `Ibaraki Beta Weight/Script/Service.vb` | Điều phối luồng `WtIbarakiBeta`, chạy lần lượt toàn bộ nhóm câu hỏi nhập liệu. |
| `Util` | `Ibaraki Beta Weight/Script/Util.vb` | Logic nghiệp vụ: mỗi nhóm nhập liệu được ánh xạ tới các ô Excel tương ứng. |
| `FrmUpdate` | `Ibaraki Beta Weight/Control/FrmUpdate.vb` | Biểu mẫu cập nhật: tải file MSI, hiển thị tiến độ, chặn Alt+F4 và chạy trình cài đặt khi tải xong. |

## Luồng Chạy

```text
Main()
 ├─ TryValidateLicense()       ← kiểm tra serial trong tài nguyên
 └─ RunApp()
     ├─ ChkUpd()               ← kiểm tra phiên bản qua mạng; mở FrmUpdate nếu có bản mới
     ├─ KillXl()               ← kết thúc các tiến trình excel.exe đang mở
     ├─ OpenFileDialog         ← người dùng chọn file *.xlsx / *.xls
     └─ Excel.Application      ← xử lý workbook bằng COM
         ├─ Workbooks.Open()
         ├─ WtIbarakiBeta()    ← ghi toàn bộ nhóm dữ liệu vào sổ tính
         ├─ Workbook.Close(SaveChanges:=True)
         └─ Process.Start()    ← mở lại file đã lưu bằng ứng dụng mặc định
```

## Nhóm Nhập Liệu

| # | Câu hỏi | Ô Excel |
|---|---|---|
| 1 | 運賃 (2トン車), kèm số lượng mặc định D16/D13/D10 | BA243, BA157 - BA159 |
| 2 | 外周/内周GL-150 | BA18 - BA25 |
| 3 | 外周深GL-300 | BA27 - BA34 |
| 4 | 外周深GL-300/+30 | BA36 - BA43 |
| 5 | 外周深GL-400 | BA45 - BA52 |
| 6 | 外周深GL-400/+30 | BA54 - BA61 |
| 7 | 玄関・勝手口 | BA63 - BA70 |
| 8 | ガレージ外周GL-300 | BA72 - BA79 |
| 9 | スラブユニット | BA99 - BA106 |
| 10 | 電気温水器 | BA107 |
| 11 | コーナー, theo D16/D13/D10 | BA163 - BA165 |
| 12 | ストレート, theo D16/D13/D10 | BA160 - BA162 |
| 13 | キャップタイヤ (320) | BA181, BA186 |
| 14 | 端部 (700×350) | BA180 |
| 15 | ロングコーナー (D16) | BA167, BA177 - BA179 |
| 16 | クランク | BA171 - BA174 |
| 17 | 島 (D16) | BA175 - BA176 |
| 18 | ストレート (D16) | BA169 - BA170, BA182 - BA184 |
| 19 | コーナー3 (D16) | BA166, BA168, BA187, BA190 |
| 20 | クランク3 (D16) | BA188 - BA189 |
| 21 | コ型3 (D16) | BA191, BA196 |
| 22 | Ｍ型 (D16[350×460×460×350]) | BA195 |
| 23 | フック (D10) | BA185, BA192 - BA194 |
| 24 | 主筋補強 (D10) | BA202 - BA203 |
| 25 | スラブ曲 (D13) | BA115 - BA124 |
| 26 | スラブ直 (D13) | BA125 - BA136 |
| 27 | スラブ補強曲 (D10) | BA137 - BA146 |
| 28 | スラブ補強直 (D10) | BA147 - BA156 |
| 29 | スリーブ | BA197 - BA201 |
| 30 | 副資材リスト, gồm vật tư phụ và thông tin công trình | BA214 - BA242, BA244, AD5, BJ12, BJ13, BO2 |

## Minh Họa Mã Nguồn

Trích nguyên văn từ `Ibaraki Beta Weight/Script/Util.vb`:

```vb
    ''' <summary>
    ''' Ghi phí vận chuyển xe 2 tấn và số lượng thép mặc định.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Fare(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            DctVal(xlApp, "BA243", choosen)
        End If
        DctVal(xlApp, "BA157", 2) ' D16
        DctVal(xlApp, "BA158", 3) ' D13
        DctVal(xlApp, "BA159", 3) ' D10
    End Sub
```

## Gói Phụ Thuộc

<img src='pic/1.png' align='left' width='3%' height='3%'></img>
<div style='display:flex;'>

- Microsoft.Office.Interop.Excel » 15.0.4795.1001

</div>

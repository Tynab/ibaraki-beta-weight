Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Tập hợp các thao tác nhập số lượng vật tư/thép và ánh xạ chúng vào đúng ô của mẫu Excel Ibaraki Beta.
''' </summary>
Friend Module Util
    Private ReadOnly GradePrompts As String() = {"  4G", "3.5G", "  3G", "2.5G", "  2G", "1.5G", "  1G", "0.5G"}
    Private ReadOnly SlabBendingPrompts As String() = {"250×5250", "250×4750", "250×4250", "250×3750", "250×3250", "250×2750", "250×2250", "250×1750", "250×1250", "250× 750"}
    Private ReadOnly SlabStraightPrompts As String() = {"5500", "5000", "4500", "4000", "3500", "3000", "2500", "2000", "1500", "1200", "1000", " 900"}
    Private ReadOnly SlabReinfStraightPrompts As String() = {"5500", "5000", "4500", "4000", "3500", "3000", "2500", "2000", "1500", "1000"}

    ''' <summary>
    ''' Kiểm tra lựa chọn 1/0 từ người dùng cho các nhóm nhập liệu tùy chọn.
    ''' </summary>
    ''' <param name="choosen">Giá trị người dùng chọn.</param>
    ''' <returns>True nếu người dùng chọn 1.</returns>
    Private Function IsSelected(choosen As Double) As Boolean
        Return choosen = 1
    End Function

    ''' <summary>
    ''' Ghi lần lượt các prompt vào các ô BA liên tiếp, bắt đầu từ dòng được truyền vào.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="firstRow">Dòng BA đầu tiên cần ghi.</param>
    ''' <param name="prompts">Danh sách nhãn nhập liệu.</param>
    Private Sub PubSequentialInputs(xlApp As Application, firstRow As Integer, prompts As IEnumerable(Of String))
        Dim row As Integer = firstRow
        For Each prompt As String In prompts
            PubDVal(xlApp, $"BA{row}", DtlDInp(vbTab & prompt & ": "))
            row += 1
        Next
    End Sub

    ''' <summary>
    ''' Ghi một nhóm prompt tuần tự chỉ khi người dùng đã bật nhóm đó.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="firstRow">Dòng BA đầu tiên cần ghi.</param>
    ''' <param name="prompts">Danh sách nhãn nhập liệu.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Private Sub PubConditionalSequentialInputs(xlApp As Application, firstRow As Integer, prompts As IEnumerable(Of String), choosen As Double)
        If IsSelected(choosen) Then
            PubSequentialInputs(xlApp, firstRow, prompts)
        End If
    End Sub

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

    ''' <summary>
    ''' Nhập số lượng ngoại/nội chu vi GL-150 theo cấp G.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    Friend Sub Unit150(xlApp As Application)
        PubSequentialInputs(xlApp, 18, GradePrompts)
    End Sub

    ''' <summary>
    ''' Nhập số lượng ngoại chu vi sâu GL-300 theo cấp G khi người dùng chọn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Unit300(xlApp As Application, choosen As Double)
        PubConditionalSequentialInputs(xlApp, 27, GradePrompts, choosen)
    End Sub

    ''' <summary>
    ''' Nhập số lượng ngoại chu vi sâu GL-300/+30 theo cấp G khi người dùng chọn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Unit300Cut(xlApp As Application, choosen As Double)
        PubConditionalSequentialInputs(xlApp, 36, GradePrompts, choosen)
    End Sub

    ''' <summary>
    ''' Nhập số lượng ngoại chu vi sâu GL-400 theo cấp G khi người dùng chọn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Unit400(xlApp As Application, choosen As Double)
        PubConditionalSequentialInputs(xlApp, 45, GradePrompts, choosen)
    End Sub

    ''' <summary>
    ''' Nhập số lượng ngoại chu vi sâu GL-400/+30 theo cấp G khi người dùng chọn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Unit400Cut(xlApp As Application, choosen As Double)
        PubConditionalSequentialInputs(xlApp, 54, GradePrompts, choosen)
    End Sub

    ''' <summary>
    ''' Nhập số lượng khu vực cửa vào/cửa sau theo cấp G khi người dùng chọn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub EntrBackDoor(xlApp As Application, choosen As Double)
        PubConditionalSequentialInputs(xlApp, 63, GradePrompts, choosen)
    End Sub

    ''' <summary>
    ''' Nhập số lượng ngoại chu vi garage GL-300 theo cấp G khi người dùng chọn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Unit300Gar(xlApp As Application, choosen As Double)
        PubConditionalSequentialInputs(xlApp, 72, GradePrompts, choosen)
    End Sub

    ''' <summary>
    ''' Nhập số lượng slab unit theo cấp G.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    Friend Sub SlabUnit(xlApp As Application)
        PubSequentialInputs(xlApp, 99, GradePrompts)
    End Sub

    ''' <summary>
    ''' Ghi số lượng bình nước nóng điện hoặc xóa ô khi không có giá trị.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="value">Số lượng cần ghi.</param>
    Friend Sub ElecWtrHtr(xlApp As Application, value As Double)
        If value > 0 Then
            DctVal(xlApp, "BA107", value)
        Else
            ClrVal(xlApp, "BA107")
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng joint góc theo cỡ thép D16/D13/D10.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    Friend Sub JtCor(xlApp As Application)
        PubDVal(xlApp, "BA165", DtlDInp(vbTab & "D16: "))
        PubDVal(xlApp, "BA164", DtlDInp(vbTab & "D13: "))
        PubDVal(xlApp, "BA163", DtlDInp(vbTab & "D10: "))
    End Sub

    ''' <summary>
    ''' Nhập số lượng joint thẳng theo cỡ thép D16/D13/D10.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    Friend Sub JtStr(xlApp As Application)
        PubDVal(xlApp, "BA162", DtlDInp(vbTab & "D16: "))
        PubDVal(xlApp, "BA161", DtlDInp(vbTab & "D13: "))
        PubDVal(xlApp, "BA160", DtlDInp(vbTab & "D10: "))
    End Sub

    ''' <summary>
    ''' Nhập số lượng cap tire 320 khi người dùng chọn nhóm này.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub CapTire(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            PubDModVal(xlApp, "186", "（コノ字型）", "680×320×680", 2.7, DtlDInp(vbTab & "D16: "))
            PubDVal(xlApp, "BA181", DtlDInp(vbTab & "D10: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng long corner D16 và dòng tùy biến 1250×1250.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub LongCor(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            PubDVal(xlApp, "BA178", DtlDInp(vbTab & " 750×1250: "))
            PubDVal(xlApp, "BA179", DtlDInp(vbTab & " 750×2250: "))
            PubDVal(xlApp, "BA177", DtlDInp(vbTab & " 750×1750: "))
            PubDModVal(xlApp, "167", "1250×1250", 4.1, DtlDInp(vbTab & "1250×1250: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng crank theo các quy cách D16/D10.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Crank(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            PubDVal(xlApp, "BA171", DtlDInp(vbTab & "D16 (750×920×750): "))
            PubDVal(xlApp, "BA172", DtlDInp(vbTab & "D10 (500×920×500): "))
            PubDVal(xlApp, "BA173", DtlDInp(vbTab & "D16 (750×460×750): "))
            PubDVal(xlApp, "BA174", DtlDInp(vbTab & "D10 (500×460×500): "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng island D16 khi người dùng chọn nhóm này.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Island(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            PubDVal(xlApp, "BA176", DtlDInp(vbTab & "350×930×350: "))
            PubDVal(xlApp, "BA175", DtlDInp(vbTab & "350×470×350: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng thép thẳng D16, gồm hai dòng có đơn giá tùy chỉnh.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Straight(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            PubDModVal(xlApp, "169", "4000", 6.3, My.Settings.Pr_D16, DtlDInp(vbTab & "4000: "))
            PubDModVal(xlApp, "170", "3500", 5.5, My.Settings.Pr_D16, DtlDInp(vbTab & "3500: "))
            PubDVal(xlApp, "BA182", DtlDInp(vbTab & "3000: "))
            PubDVal(xlApp, "BA183", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA184", DtlDInp(vbTab & "2000: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng corner 3D D16 cho bên phải/trái.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Corner3d(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            PubDVal(xlApp, "BA166", DtlDInp(vbTab & "右 (750×460×350): "))
            PubDVal(xlApp, "BA168", DtlDInp(vbTab & "左 (750×460×350): "))
            PubDModVal(xlApp, "187", "750×240×350", 2.2, DtlDInp(vbTab & "右 (750×240×350): "))
            PubDModVal(xlApp, "190", "750×240×350", 2.2, DtlDInp(vbTab & "左 (750×240×350): "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng crank 3D D16 cho bên phải/trái.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub Crank3d(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            PubDModVal(xlApp, "189", "（クランク３右）", "750×460×460×350", 3.3, DtlDInp(vbTab & "右 (750×460×460×350): "))
            PubDModVal(xlApp, "188", "（クランク３左）", "750×460×460×350", 3.3, DtlDInp(vbTab & "左 (750×460×460×350): "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng U-type 3D D16 cho bên phải/trái.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub UType3d(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            PubDModVal(xlApp, "191", "（コノ字３右）", "750×460×460×350", 3.3, DtlDInp(vbTab & "右 (750×460×460×350): "))
            PubDModVal(xlApp, "196", "（コノ字３左）", "750×460×460×350", 3.3, DtlDInp(vbTab & "左 (750×460×460×350): "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng hook D10 theo các quy cách cố định.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    Friend Sub Hook(xlApp As Application)
        PubDModVal(xlApp, "192", "695×160　　フック付", 0.6, DtlDInp(vbTab & "695×160: "))
        PubDModVal(xlApp, "193", "595×160　　フック付", 0.5, DtlDInp(vbTab & "595×160: "))
        PubDModVal(xlApp, "194", "160×160　　フック付", 0.3, DtlDInp(vbTab & "160×160: "))
        PubDVal(xlApp, "BA185", DtlDInp(vbTab & "435×250: "))
    End Sub

    ''' <summary>
    ''' Nhập số lượng thép gia cường chính D10 khi người dùng chọn nhóm này.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub MainReinf(xlApp As Application, choosen As Double)
        If IsSelected(choosen) Then
            PubDVal(xlApp, "BA202", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA203", DtlDInp(vbTab & "2000: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng slab uốn D13 theo các chiều dài chuẩn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub SlabBndg(xlApp As Application, choosen As Double)
        PubConditionalSequentialInputs(xlApp, 115, SlabBendingPrompts, choosen)
    End Sub

    ''' <summary>
    ''' Nhập số lượng slab thẳng D13 theo các chiều dài chuẩn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    Friend Sub SlabStr(xlApp As Application)
        PubSequentialInputs(xlApp, 125, SlabStraightPrompts)
    End Sub

    ''' <summary>
    ''' Nhập số lượng slab gia cường uốn D10 theo các chiều dài chuẩn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub SlabReinfBndg(xlApp As Application, choosen As Double)
        PubConditionalSequentialInputs(xlApp, 137, SlabBendingPrompts, choosen)
    End Sub

    ''' <summary>
    ''' Nhập số lượng slab gia cường thẳng D10 theo các chiều dài chuẩn.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="choosen">Giá trị chọn 1/0.</param>
    Friend Sub SlabReinfStr(xlApp As Application, choosen As Double)
        PubConditionalSequentialInputs(xlApp, 147, SlabReinfStraightPrompts, choosen)
    End Sub

    ''' <summary>
    ''' Ghi số lượng sleeve vào các dòng liên quan, riêng BA200 dùng gấp đôi số lượng.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="value">Số lượng sleeve.</param>
    Friend Sub Sleeve(xlApp As Application, value As Double)
        If value > 0 Then
            DctVal(xlApp, "BA198", value)
            DctVal(xlApp, "BA197", value)
            DctVal(xlApp, "BA199", value)
            DctVal(xlApp, "BA200", value * 2)
            DctVal(xlApp, "BA201", value)
        End If
    End Sub

    ''' <summary>
    ''' Nhập thông tin công trình và danh sách phụ tư vào sheet phụ tư.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    Friend Sub Parts(xlApp As Application)
        Dim name As String = $"{DtlSInp(vbTab & "邸名" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ")}様邸"
        DctVal(xlApp, "BJ12", name)
        CType(xlApp.ActiveSheet, Worksheet).Name = name
        PubSVal(xlApp, vbTab & "住所" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BJ13")
        PubSVal(xlApp, vbTab & "邸名コード" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "AD5")
        PubSVal(xlApp, vbTab & "納品日" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BO2")
        Dim ipp As Double = DtlYNQ(vbTab & "運賃 (分納)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ")
        If IsSelected(ipp) Then
            DctVal(xlApp, "BA244", ipp)
        End If
        PubDVal(xlApp, "BA214", DtlDInpDesc(vbTab & "フラットアンカーボルト (本)", vbTab & vbTab & "[M12×350]" & vbTab))
        PubDVal(xlApp, "BA215", DtlDInpDesc(vbTab & "カットスクリュー・Ⅱ (袋)", vbTab & vbTab & "[M12用]" & vbTab & vbTab))
        PubDVal(xlApp, "BA216", DtlDInp(vbTab & "カットスクリュー・Ⅱ専用ピット (個)" & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA217", DtlDInpDesc(vbTab & "ホールダウンアンカーボルト (本)", vbTab & vbTab & "[M12×700]" & vbTab))
        PubDVal(xlApp, "BA218", DtlDInpDesc(vbTab & "アンカーグリッパーM12用 (箱)", vbTab & vbTab & "[D10 TG1210D]" & vbTab))
        PubDVal(xlApp, "BA219", DtlDInpDesc(vbTab & "アンカーグリッパーM12用 (箱)", vbTab & vbTab & "[D13 TG1213D]" & vbTab))
        PubDVal(xlApp, "BA220", DtlDInpDesc(vbTab & "アンカーグリッパーM12用 (箱)", vbTab & vbTab & "[D16 TG1216D]" & vbTab))
        PubDVal(xlApp, "BA237", DtlDInpDesc(vbTab & "マグネット差し筋アンカー (ｾｯﾄ)", vbTab & vbTab & "[直]" & vbTab & vbTab))
        PubDVal(xlApp, "BA236", DtlDInpDesc(vbTab & "マグネット差し筋アンカー (ｾｯﾄ)", vbTab & vbTab & "[曲]" & vbTab & vbTab))
        PubDVal(xlApp, "BA221", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & "[60ﾐﾘ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA222", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & "[80ﾐﾘ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA223", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & "[60×70×80]" & vbTab))
        PubDVal(xlApp, "BA225", DtlDInpDesc(vbTab & "排水用スリーブホルダー D10用 (袋)", vbTab & "[50Φ・75Φ用]" & vbTab))
        PubDVal(xlApp, "BA226", DtlDInpDesc(vbTab & "給水用スリーブホルダー D10用 (袋)", vbTab & "[50Φ]" & vbTab & vbTab))
        Dim curingShRingTree As Double = DtlDInpDesc(vbTab & "養生シート輪木 (ｾｯﾄ)", vbTab & vbTab & vbTab & "[3.6×5.4]" & vbTab)
        If curingShRingTree > 0 Then
            DctVal(xlApp, "BA227", curingShRingTree)
        Else
            DctVal(xlApp, "BA227", 1)
            ClrVal(xlApp, "BF227")
            ClrVal(xlApp, "CB227")
        End If
        PubDVal(xlApp, "BA228", DtlDInp(vbTab & "Ｍ型鉄筋ベース (個)" & vbTab & vbTab & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA229", DtlDInpDesc(vbTab & "樹脂スペーサー改 (ｹｰｽ)", vbTab & vbTab & vbTab & "[300ヶ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA232", DtlDInpDesc(vbTab & "アンカーボルトセット (ｾｯﾄ)", vbTab & vbTab & "[M18×380]" & vbTab))
        PubDVal(xlApp, "BA234", DtlDInpDesc(vbTab & "NSP吊巾止 W160用 (本)", vbTab & vbTab & vbTab & "[200本]" & vbTab & vbTab))
        PubDVal(xlApp, "BA238", DtlDInpDesc(vbTab & "アンカーボルト (本)", vbTab & vbTab & vbTab & "[M16×415]" & vbTab))
        ' Các vật tư bổ sung nằm ở phần mở rộng của mẫu Excel.
        PubDVal(xlApp, "BA224", DtlDInpDesc(vbTab & "樹脂スペーサー (個)", vbTab & vbTab & vbTab & "[70×80]" & vbTab & vbTab))
        PubDVal(xlApp, "BA230", DtlDInpDesc(vbTab & "鉄筋スペーサー (個)", vbTab & vbTab & vbTab & "[60ﾖｳ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA231", DtlDInpDesc(vbTab & "鉄筋スペーサー (個)", vbTab & vbTab & vbTab & "[80ﾖｳ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA233", DtlDInpDesc(vbTab & "偏心用鉄筋ベース (個)", vbTab & vbTab & vbTab & "[280×160×60]" & vbTab))
        PubDVal(xlApp, "BA235", DtlDInpDesc(vbTab & "防錆巾止め金具 (本)", vbTab & vbTab & vbTab & "[Fﾊﾟﾈﾙ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA240", DtlDInpDesc(vbTab & "アンカーボルトセット (本)", vbTab & vbTab & "[M12×498]" & vbTab))
        PubDVal(xlApp, "BA241", DtlDInpDesc(vbTab & "アンカーボルトセット軸柱用 (本)", vbTab & vbTab & "[M12×498]" & vbTab))
        PubDVal(xlApp, "BA242", DtlDInpDesc(vbTab & "Ｕボルト (ｾｯﾄ)", vbTab & vbTab & vbTab & vbTab & "[M8]" & vbTab & vbTab))
        PubDVal(xlApp, "BA239", DtlDInpDesc(vbTab & "アンカーボルト (本)", vbTab & vbTab & vbTab & "[M16×417]" & vbTab))
    End Sub
End Module

Imports System.Console
Imports System.ConsoleColor
Imports System.IO
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Threading.Thread
Imports System.Windows.Forms

''' <summary>
''' Cung cấp helper chung cho cập nhật ứng dụng, vòng đời Excel, nhập liệu console,
''' định dạng console và thao tác ghi dữ liệu vào mẫu Excel.
''' </summary>
Friend Module Common
#Region "Helper"
    ''' <summary>
    ''' Kiểm tra nhanh endpoint update có phản hồi hay không.
    ''' </summary>
    ''' <returns>True nếu kết nối tới endpoint cơ sở thành công.</returns>
    Private Function IsNetAvail() As Boolean
        Try
            Dim request = DirectCast(WebRequest.Create(New Uri(My.Resources.link_base)), HttpWebRequest)
            request.Timeout = 5000
            request.ReadWriteTimeout = 5000

            Using response As WebResponse = request.GetResponse()
            End Using

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Kiểm tra phiên bản mới và mở form cập nhật khi server báo khác phiên bản hiện tại.
    ''' </summary>
    Private Sub ChkUpd()
        HdrSty("アップデートの確認...")

        If Not IsNetAvail() Then
            Return
        End If

        Try
            Using webClient As New WebClient()
                If Not webClient.DownloadString(My.Resources.link_ver).Contains(My.Resources.app_ver) Then
                    MsgBox($"「{My.Resources.app_true_name}」新しいバージョンが利用可能！", 262144, Title:="更新")
                    Using frmUpd As New FrmUpdate()
                        frmUpd.ShowDialog()
                    End Using
                End If
            End Using
        Catch ex As Exception
            ErrSty("アップデートの確認に失敗しました。既存バージョンで続行します..." & vbCrLf)
        End Try
    End Sub

    ''' <summary>
    ''' Lưu trạng thái license hợp lệ vào user settings.
    ''' </summary>
    Friend Sub UpdVldLic()
        My.Settings.Chk_Key = True
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Tăng dần độ trong suốt để form hiện ra mượt hơn.
    ''' </summary>
    <Extension()>
    Friend Sub FIFrm(frm As Form)
        While frm.Opacity < 1
            frm.Opacity += 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub

    ''' <summary>
    ''' Giảm dần độ trong suốt để form đóng mượt hơn.
    ''' </summary>
    <Extension()>
    Friend Sub FOFrm(frm As Form)
        While frm.Opacity > 0
            frm.Opacity -= 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub
#End Region

#Region "Master"
    ''' <summary>
    ''' Kết thúc tất cả tiến trình theo tên, dùng cho Excel và tiến trình app cũ sau khi update.
    ''' </summary>
    ''' <param name="name">Tên tiến trình không kèm phần mở rộng.</param>
    Friend Sub KillPrcs(name As String)
        For Each item As Process In Process.GetProcessesByName(name)
            Try
                item.Kill()
                item.WaitForExit(5000)
            Catch ex As Exception
                ' Bỏ qua tiến trình đã thoát hoặc không có quyền, vì đây chỉ là bước dọn dẹp.
            Finally
                item.Dispose()
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Nhắc người dùng đóng Excel rồi dọn các tiến trình Excel còn sót để tránh khóa workbook.
    ''' </summary>
    Private Sub KillXl()
        Clear()
        HdrSty("警告：このアプリケーションを使用する前に、すべての「エクセル」を閉じてください。「エンター」キーを押して続行します...")
        ReadLine()
        KillPrcs(XL_NAME)
    End Sub

    ''' <summary>
    ''' Nhả COM object của Excel nếu đã được tạo để hạn chế Excel.exe bị giữ lại.
    ''' </summary>
    ''' <param name="value">COM object cần nhả.</param>
    Private Sub ReleaseComObjectSafe(value As Object)
        If value IsNot Nothing AndAlso Marshal.IsComObject(value) Then
            Marshal.FinalReleaseComObject(value)
        End If
    End Sub

    ''' <summary>
    ''' Chạy luồng chính: kiểm tra update, chọn file Excel, ghi dữ liệu và mở lại file kết quả.
    ''' </summary>
    Friend Sub RunApp()
        ChkUpd()
        KillXl()

        Using ofd As New OpenFileDialog With {
            .Multiselect = False,
            .Title = "「エクセル」ドキュメントを開く",
            .Filter = "「エクセル」ドキュメント|*.xlsx;*.xls"
        }
            If ofd.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            Dim filePath As String = ofd.FileName
            Dim xlApp As Microsoft.Office.Interop.Excel.Application = Nothing
            Dim workbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Dim completed As Boolean = False

            Try
                xlApp = New Microsoft.Office.Interop.Excel.Application With {
                    .DisplayAlerts = False
                }
                workbook = xlApp.Workbooks.Open(filePath)
                WtIbarakiBeta(xlApp)
                completed = True
            Catch ex As Exception
                ErrSty($"処理中にエラーが発生しました: {ex.Message}{vbCrLf}")
            Finally
                If workbook IsNot Nothing Then
                    Try
                        workbook.Close(SaveChanges:=completed)
                    Catch ex As Exception
                        ErrSty($"ワークブックを閉じる時にエラーが発生しました: {ex.Message}{vbCrLf}")
                    Finally
                        ReleaseComObjectSafe(workbook)
                    End Try
                End If

                If xlApp IsNot Nothing Then
                    Try
                        xlApp.Quit()
                    Catch ex As Exception
                        ErrSty($"Excelを終了する時にエラーが発生しました: {ex.Message}{vbCrLf}")
                    Finally
                        ReleaseComObjectSafe(xlApp)
                    End Try
                End If
            End Try

            If completed Then
                Process.Start(filePath)
            End If
        End Using
    End Sub
#End Region

#Region "Main"
    ''' <summary>
    ''' Tạo thư mục nếu chưa tồn tại.
    ''' </summary>
    ''' <param name="path">Đường dẫn thư mục.</param>
    Friend Sub CrtDirAdv(path As String)
        If Not Directory.Exists(path) Then
            Directory.CreateDirectory(path)
        End If
    End Sub

    ''' <summary>
    ''' Xóa file nếu đang tồn tại.
    ''' </summary>
    ''' <param name="path">Đường dẫn file.</param>
    Friend Sub DelFileAdv(path As String)
        If File.Exists(path) Then
            File.Delete(path)
        End If
    End Sub

    ''' <summary>
    ''' Kiểm tra giá trị chọn 0/1.
    ''' </summary>
    ''' <param name="value">Giá trị người dùng nhập.</param>
    ''' <returns>True nếu giá trị là 0 hoặc 1.</returns>
    Private Function IsBinaryChoice(value As Double) As Boolean
        Return value = 0 OrElse value = 1
    End Function

    ''' <summary>
    ''' Hiển thị câu hỏi Yes/No ở cấp header và chỉ nhận giá trị 1/0.
    ''' </summary>
    ''' <param name="caption">Nhãn câu hỏi.</param>
    ''' <returns>Giá trị chọn của người dùng.</returns>
    Friend Function HdrYNQ(caption As String) As Double
        Dim value As Double = HdrDWrng(caption)
        Do Until IsBinaryChoice(value)
            value = HdrDErr(caption)
        Loop
        Return value
    End Function

    ''' <summary>
    ''' Hiển thị câu hỏi Yes/No ở cấp detail và chỉ nhận giá trị 1/0.
    ''' </summary>
    ''' <param name="caption">Nhãn câu hỏi.</param>
    ''' <returns>Giá trị chọn của người dùng.</returns>
    Friend Function DtlYNQ(caption As String) As Double
        PrefSel(caption)
        Dim value As Double = Val(ReadLine)
        Do Until IsBinaryChoice(value)
            PrefWrng(caption)
            value = Val(ReadLine)
        Loop
        Return value
    End Function

    ''' <summary>
    ''' Ghi trực tiếp một giá trị vào ô Excel.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    ''' <param name="value">Giá trị cần ghi.</param>
    Friend Sub DctVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, value As Object)
        Dim target As Microsoft.Office.Interop.Excel.Range = xlApp.Range(cell)
        Try
            target.FormulaR1C1 = value
        Finally
            ReleaseComObjectSafe(target)
        End Try
    End Sub

    ''' <summary>
    ''' Ghi giá trị vào ô Excel và tô màu để đánh dấu dòng được tùy biến.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    ''' <param name="value">Giá trị cần ghi.</param>
    Private Sub ModVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, value As Object)
        Dim target As Microsoft.Office.Interop.Excel.Range = xlApp.Range(cell)
        Try
            target.FormulaR1C1 = value
            target.Interior.Color = RGB(0, 176, 240)
        Finally
            ReleaseComObjectSafe(target)
        End Try
    End Sub

    ''' <summary>
    ''' Xóa nội dung của ô hoặc vùng merge chứa ô đó.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    Friend Sub ClrVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String)
        Dim target As Microsoft.Office.Interop.Excel.Range = xlApp.Range(cell)
        Try
            target.MergeArea.ClearContents()
        Finally
            ReleaseComObjectSafe(target)
        End Try
    End Sub

    ''' <summary>
    ''' Nhập chuỗi từ console rồi ghi vào ô Excel.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    Friend Sub PubSVal(xlApp As Microsoft.Office.Interop.Excel.Application, caption As String, cell As String)
        DctVal(xlApp, cell, DtlSInp(caption))
    End Sub

    ''' <summary>
    ''' Ghi số dương vào ô Excel, bỏ qua giá trị 0 hoặc âm.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    ''' <param name="value">Giá trị số cần ghi.</param>
    Friend Sub PubDVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, value As Double)
        If value > 0 Then
            DctVal(xlApp, cell, value)
        End If
    End Sub

    ''' <summary>
    ''' Ghi một dòng thép tùy biến gồm tên, trọng lượng và số lượng.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="row">Số dòng trong mẫu Excel.</param>
    ''' <param name="name">Tên hoặc quy cách thép.</param>
    ''' <param name="weight">Trọng lượng đơn vị.</param>
    ''' <param name="number">Số lượng.</param>
    Friend Sub PubDModVal(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Ghi một dòng thép tùy biến có thêm tiêu đề/ghi chú phân loại.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="row">Số dòng trong mẫu Excel.</param>
    ''' <param name="title">Tiêu đề hoặc loại thép.</param>
    ''' <param name="name">Tên hoặc quy cách thép.</param>
    ''' <param name="weight">Trọng lượng đơn vị.</param>
    ''' <param name="number">Số lượng.</param>
    Friend Sub PubDModVal(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, title As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"X{row}", title)
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Ghi một dòng thép tùy biến có cả trọng lượng và đơn giá.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang mở.</param>
    ''' <param name="row">Số dòng trong mẫu Excel.</param>
    ''' <param name="name">Tên hoặc quy cách thép.</param>
    ''' <param name="weight">Trọng lượng đơn vị.</param>
    ''' <param name="price">Đơn giá.</param>
    ''' <param name="number">Số lượng.</param>
    Friend Sub PubDModVal(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, name As String, weight As Double, price As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            ModVal(xlApp, $"CQ{row}", price)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub
#End Region

#Region "Timer"
    ''' <summary>
    ''' Bật timer nếu timer chưa chạy.
    ''' </summary>
    <Extension()>
    Friend Sub StrtAdv(tmr As Timer)
        If Not tmr.Enabled Then
            tmr.Start()
        End If
    End Sub

    ''' <summary>
    ''' Tắt timer nếu timer đang chạy.
    ''' </summary>
    <Extension()>
    Friend Sub StopAdv(tmr As Timer)
        If tmr.Enabled Then
            tmr.Stop()
        End If
    End Sub
#End Region

#Region "Actor"
    ''' <summary>
    ''' In nội dung cấp header bằng màu cảnh báo nhẹ.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub HdrSty(caption As String)
        ForegroundColor = DarkYellow
        Write(caption)
    End Sub

    ''' <summary>
    ''' In nội dung giới thiệu bằng màu xanh dương.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub IntroSty(caption As String)
        ForegroundColor = Blue
        Write(caption)
    End Sub

    ''' <summary>
    ''' In tiêu đề bằng màu xanh lá.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub TitSty(caption As String)
        ForegroundColor = Green
        Write(caption)
    End Sub

    ''' <summary>
    ''' In nhãn nhập liệu bằng màu cyan.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub InpSty(caption As String)
        ForegroundColor = Cyan
        Write(caption)
    End Sub

    ''' <summary>
    ''' In mô tả bổ sung bằng màu magenta.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub DescSty(caption As String)
        ForegroundColor = Magenta
        Write(caption)
    End Sub

    ''' <summary>
    ''' In cảnh báo hoặc câu hỏi lựa chọn bằng màu vàng.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub WrngSty(caption As String)
        ForegroundColor = Yellow
        Write(caption)
    End Sub

    ''' <summary>
    ''' In lỗi bằng màu đỏ.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Friend Sub ErrSty(caption As String)
        ForegroundColor = Red
        Write(caption)
    End Sub

    ''' <summary>
    ''' In prefix cho dòng nhập liệu rồi trả màu chữ về trắng.
    ''' </summary>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    Private Sub PrefInp(caption As String)
        InpSty(caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' In prefix cho dòng chọn 0/1 rồi trả màu chữ về trắng.
    ''' </summary>
    ''' <param name="caption">Nhãn lựa chọn.</param>
    Private Sub PrefSel(caption As String)
        WrngSty(caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' In lại prefix khi giá trị nhập không hợp lệ.
    ''' </summary>
    ''' <param name="caption">Nhãn lựa chọn.</param>
    Private Sub PrefWrng(caption As String)
        WrngSty(caption)
        ForegroundColor = Red
    End Sub

    ''' <summary>
    ''' In mô tả nằm sau nhãn nhập liệu.
    ''' </summary>
    ''' <param name="description">Mô tả bổ sung.</param>
    Private Sub SfxDesc(description As String)
        DescSty(description)
        PrefInp(": ")
    End Sub

    ''' <summary>
    ''' Xóa màn hình và in banner ứng dụng trước mỗi cụm nhập liệu lớn.
    ''' </summary>
    Private Sub Intro()
        Clear()
        IntroSty(My.Resources.gr_name & vbCrLf)
        IntroSty(My.Resources.cc_text & vbCrLf)
        TitSty(vbCrLf & My.Resources.app_true_name & vbCrLf & vbCrLf)
    End Sub

    ''' <summary>
    ''' Hiển thị banner rồi đọc một số từ người dùng.
    ''' </summary>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function HdrDInp(caption As String) As Double
        Intro()
        Return DtlDInp(caption)
    End Function

    ''' <summary>
    ''' Hiển thị banner rồi in cảnh báo/chủ đề của nhóm nhập liệu.
    ''' </summary>
    ''' <param name="caption">Nội dung cảnh báo hoặc tiêu đề nhóm.</param>
    Friend Sub HdrWrng(caption As String)
        Intro()
        WrngSty(caption)
    End Sub

    ''' <summary>
    ''' Hiển thị banner, đọc lựa chọn số lần đầu với màu cảnh báo.
    ''' </summary>
    ''' <param name="caption">Nhãn lựa chọn.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function HdrDWrng(caption As String) As Double
        Intro()
        PrefSel(caption)
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Hiển thị banner và đọc lại lựa chọn số khi lần nhập trước không hợp lệ.
    ''' </summary>
    ''' <param name="caption">Nhãn lựa chọn.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function HdrDErr(caption As String) As Double
        Intro()
        PrefWrng(caption)
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Đọc một giá trị số ở dòng nhập liệu chi tiết.
    ''' </summary>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function DtlDInp(caption As String) As Double
        PrefInp(caption)
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Đọc một chuỗi ở dòng nhập liệu chi tiết.
    ''' </summary>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    ''' <returns>Chuỗi người dùng nhập.</returns>
    Friend Function DtlSInp(caption As String) As String
        PrefInp(caption)
        Return If(ReadLine(), String.Empty)
    End Function

    ''' <summary>
    ''' Đọc một giá trị số ở dòng nhập liệu chi tiết có mô tả phụ.
    ''' </summary>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    ''' <param name="description">Mô tả phụ.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function DtlDInpDesc(caption As String, description As String) As Double
        InpSty(caption)
        SfxDesc(description)
        Return Val(ReadLine)
    End Function
#End Region
End Module

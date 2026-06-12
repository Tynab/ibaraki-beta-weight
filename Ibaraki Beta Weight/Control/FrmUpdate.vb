Imports System.ComponentModel
Imports System.IO
Imports System.Math
Imports System.Net
Imports System.Windows.Forms
Imports System.Windows.Forms.Keys

''' <summary>
''' Form toàn màn hình nhỏ dùng để tải bộ cài mới và khởi chạy MSI sau khi tải xong.
''' </summary>
Public Class FrmUpdate
#Region "Fields"
    Private ReadOnly _wc As New WebClient
#End Region

#Region "Overridden"
    ''' <summary>
    ''' Ẩn form khỏi danh sách Alt+Tab trong lúc cập nhật.
    ''' </summary>
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H80
            Return cp
        End Get
    End Property

    ''' <summary>
    ''' Chặn Alt+F4 để tránh đóng dở lúc đang tải bộ cài.
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Return keyData = (Alt Or F4) OrElse MyBase.ProcessCmdKey(msg, keyData)
    End Function
#End Region

#Region "Events"
    ''' <summary>
    ''' Khởi tạo trạng thái hiển thị trước khi bắt đầu download.
    ''' </summary>
    Private Sub FrmUpdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblCapacity.Text = ""
        lblPercent.Text = ""
        pnlProgressBar.Width = 1
        tmrMain.StopAdv()
    End Sub

    ''' <summary>
    ''' Chuẩn bị thư mục update và bắt đầu tải bộ cài mới.
    ''' </summary>
    Private Sub FrmUpdate_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        FIFrm()
        CrtDirAdv(FRNT_PATH)
        DelFileAdv(FILE_SETUP_ADR)
        AddHandler _wc.DownloadProgressChanged, AddressOf Upd_DownloadProgressChanged
        AddHandler _wc.DownloadFileCompleted, AddressOf Upd_DownloadFileCompleted

        Try
            Dim setupUrl As String = _wc.DownloadString(My.Resources.link_app).Trim()
            _wc.DownloadFileAsync(New Uri(setupUrl), FILE_SETUP_ADR)
        Catch ex As Exception
            MessageBox.Show("アップデートを開始できませんでした。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Close()
        End Try
    End Sub

    ''' <summary>
    ''' Cập nhật dung lượng, phần trăm và độ rộng thanh progress trong lúc download.
    ''' </summary>
    Private Sub Upd_DownloadProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs)
        Dim receivedMb As String = (e.BytesReceived / 1024D / 1024D).ToString("0.00")
        If e.TotalBytesToReceive > 0 Then
            Dim totalMb As String = (e.TotalBytesToReceive / 1024D / 1024D).ToString("0.00")
            lblCapacity.Text = $"{receivedMb} MB / {totalMb} MB"
        Else
            lblCapacity.Text = $"{receivedMb} MB"
        End If

        lblPercent.Text = $"{e.ProgressPercentage}%"
        pnlProgressBar.Width = CInt(Ceiling(e.ProgressPercentage * pnlMain.Width / 100D))
    End Sub

    ''' <summary>
    ''' Đóng form khi download hoàn tất, hoặc báo lỗi nếu file không tải được.
    ''' </summary>
    Private Sub Upd_DownloadFileCompleted(sender As Object, e As AsyncCompletedEventArgs)
        If e.Cancelled OrElse e.Error IsNot Nothing Then
            MessageBox.Show("アップデートのダウンロードに失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Close()
            Return
        End If

        lblPercent.Text = "100%"
        pnlProgressBar.Width = pnlMain.Width
        Close()
    End Sub

    ''' <summary>
    ''' Timer dự phòng cho designer cũ; workflow hiện đóng bằng DownloadFileCompleted.
    ''' </summary>
    Private Sub TmrMain_Tick(sender As Object, e As EventArgs) Handles tmrMain.Tick
        If lblPercent.Text = "100%" Then
            tmrMain.StopAdv()
            Close()
        End If
    End Sub

    ''' <summary>
    ''' Fade out form trước khi đóng.
    ''' </summary>
    Private Sub FrmUpdate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        FOFrm()
    End Sub

    ''' <summary>
    ''' Gỡ event, giải phóng WebClient và mở MSI đã tải nếu file tồn tại.
    ''' </summary>
    Private Sub FrmUpdate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        RemoveHandler _wc.DownloadProgressChanged, AddressOf Upd_DownloadProgressChanged
        RemoveHandler _wc.DownloadFileCompleted, AddressOf Upd_DownloadFileCompleted
        _wc.Dispose()

        If File.Exists(FILE_SETUP_ADR) Then
            Process.Start(FILE_SETUP_ADR)
            KillPrcs(My.Resources.app_name)
        End If
    End Sub
#End Region
End Class

Imports System.Console
Imports System.Text.Encoding
Imports System.Windows.Forms
Imports System.Windows.Forms.DialogResult
Imports System.Windows.Forms.MessageBox
Imports System.Windows.Forms.MessageBoxButtons

''' <summary>
''' Module entrypoint chịu trách nhiệm xác thực license và khởi động workflow chính.
''' </summary>
Public Module Main
    ''' <summary>
    ''' Điểm vào của ứng dụng: cấu hình console UTF-8, xác thực license nếu cần,
    ''' rồi chạy luồng nhập liệu và ghi dữ liệu vào Excel.
    ''' </summary>
    Public Sub Main()
        OutputEncoding = UTF8

        If My.Settings.Chk_Key OrElse TryValidateLicense() Then
            RunApp()
        Else
            ErrSty("終了するには、任意のキーを押してください...")
            ReadKey()
        End If
    End Sub

    ''' <summary>
    ''' Hỏi serial cho tới khi người dùng nhập đúng hoặc chọn hủy.
    ''' </summary>
    ''' <returns>True nếu license hợp lệ, ngược lại là False.</returns>
    Private Function TryValidateLicense() As Boolean
        Do
            If InputBox("シリアルを入力", "ライセンスキー") = My.Resources.key_ser Then
                UpdVldLic()
                Return True
            End If

            If Show("ライセンスが間違っています！", "エラー", RetryCancel, MessageBoxIcon.Error) <> Retry Then
                Return False
            End If
        Loop
    End Function
End Module

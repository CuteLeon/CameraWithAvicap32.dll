Option Explicit On
Public Class Form1
    Dim Video_Handle As Long

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Disconnect(Video_Handle)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Video_Handle = CreateCaptureWindow(PicCapture.Handle, , , PicCapture.Width, PicCapture.Height, 0)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Jpeg As Image
        Jpeg = CapturePicture(Video_Handle)
        If Jpeg Is Nothing Then MsgBox("无法拍摄") : Exit Sub
        Jpeg.Save("D:\" & My.Computer.Clock.LocalTime.Millisecond.ToString & ".jpg", Imaging.ImageFormat.Jpeg)
    End Sub

End Class

Option Explicit On
Module DealCamera
    Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hWndParent As Integer, ByVal nID As Integer) As Integer

    Private Const WS_CHILD = &H40000000
    Private Const WS_VISIBLE = &H10000000
    Private Const WM_USER = &H400
    Private Const WM_CAP_START = &H400
    Private Const WM_CAP_EDIT_COPY = (WM_CAP_START + 30)
    Private Const WM_CAP_DRIVER_CONNECT = (WM_CAP_START + 10)
    Private Const WM_CAP_SET_PREVIEWRATE = (WM_CAP_START + 52)
    Private Const WM_CAP_SET_OVERLAY = (WM_CAP_START + 51)
    Private Const WM_CAP_SET_PREVIEW = (WM_CAP_START + 50)
    Private Const WM_CAP_DRIVER_DISCONNECT = (WM_CAP_START + 11)

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As VariantType) As Integer

    Private Preview_Handle As Integer

    Public Function CreateCaptureWindow(ByVal hWndParent As Integer, Optional ByVal x As Integer = 0, Optional ByVal y As Integer = 0, Optional ByVal nWidth As Integer = 320, Optional ByVal nHeight As Integer = 240, Optional ByVal nCameraID As Integer = 0) As Integer
        Preview_Handle = capCreateCaptureWindow("Video", WS_CHILD + WS_VISIBLE, x, y, nWidth, nHeight, hWndParent, 1)

        SendMessage(Preview_Handle, WM_CAP_DRIVER_CONNECT, nCameraID, 0)
        SendMessage(Preview_Handle, WM_CAP_SET_PREVIEWRATE, 30, 0)
        SendMessage(Preview_Handle, WM_CAP_SET_OVERLAY, 1, 0)
        SendMessage(Preview_Handle, WM_CAP_SET_PREVIEW, 1, 0)

        CreateCaptureWindow = Preview_Handle
    End Function

    Public Function CapturePicture(ByVal nCaptureHandle As Long) As Object
        Clipboard.Clear()
        SendMessage(nCaptureHandle, WM_CAP_EDIT_COPY, 0, 0)
        CapturePicture = CType(Clipboard.GetDataObject().GetData(GetType(System.Drawing.Bitmap)), Image)
    End Function

    Public Sub Disconnect(ByVal nCaptureHandle As Long, Optional ByVal nCameraID As Long = 0)
        SendMessage(nCaptureHandle, WM_CAP_DRIVER_DISCONNECT, nCameraID, 0)
    End Sub

End Module

Attribute VB_Name = "Module1"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046

Public Sub get_desktop()
DeskhWnd& = GetDesktopWindow()  'We need the hwnd of the desktop
DeskDC& = GetDC(DeskhWnd&)       'We got the hwnd, lets get the dc
BitBlt FrmMain.Picbg.hdc, 0&, 0&, Screen.Width / _
15, Screen.Height / 15, DeskDC&, 0&, 0&, SRCCOPY 'Put it in a picturebox for later use
jackh = False
Call BitBlt(FrmMain.hdc, 0, 0, FrmMain.ScaleWidth, _
FrmMain.ScaleHeight, FrmMain.Picbg.hdc, 0, 0, vbSrcCopy) 'Paint the dc on the form

End Sub


Attribute VB_Name = "Module2"
' Type of one point in our Shaped form
Type POINTAPI
    X As Long
    Y As Long
End Type

' Array of points that will create the desired polygon
Public Dots() As POINTAPI

' The three API functions needed to create the polygon
' and set it as the form's region
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CreatePolygonRgn& Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long)
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
  ByVal cy As Long, ByVal wFlags As Long) As Long

' The two API Functions needed to allow the form to be
' moved by pressing anywhere on it, and not only on
' the caption
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As _
  Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Global Const WM_NCLBUTTONDOWN = &HA1
Global Const WM_LBUTTONCLK = &H202
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_RBUTTONUP = &H205
Global Const WM_MOUSEMOVE = &H200
Global Const HTCAPTION = 2

' The function that set the form's shape:
' Params: Pnt - the first point in an array of POINTAPI
'         Frm - the form that will be changed
Public Sub SetForm(Pnt As POINTAPI, frm As Form)
    l = CreatePolygonRgn(Pnt, UBound(Dots), 2)
    SetWindowRgn frm.hwnd, l, True
End Sub

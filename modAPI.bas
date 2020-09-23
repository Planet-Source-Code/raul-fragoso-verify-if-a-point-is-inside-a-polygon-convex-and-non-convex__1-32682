Attribute VB_Name = "modAPI"
'Declares PI
Public Const PI As Double = 3.14159265358979

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


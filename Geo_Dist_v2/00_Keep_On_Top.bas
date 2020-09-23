Attribute VB_Name = "Keep_On_Top"
  Option Explicit

' This function is called to keep the display on top of all other windows.

  Public Declare Function SetWindowPos Lib "user32" _
 (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
  ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' USAGE: SetWindowPos FormName.hwnd, -1,0,0,0,0,3 ' Bring on top
' To restore to normal, change -1 to -2

  Public Sub Keep_On_Top()
  SetWindowPos Geodesic_Distance_Form.hwnd, -1, 0, 0, 0, 0, 3
  End Sub


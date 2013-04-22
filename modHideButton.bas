Attribute VB_Name = "modHideButton"
Option Compare Database
Option Explicit

' ********** Code Start *************

'   Call Buttons(False)
'To turn them off and
'    Call Buttons(True)
'to turn them on

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_SYSMENU = &H80000


Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20

Private Declare Function GetWindowLong _
Lib "user32" Alias "GetWindowLongA" ( _
  ByVal hwnd As Long, _
  ByVal nIndex As Long _
) As Long

Private Declare Function SetWindowLong _
Lib "user32" Alias "SetWindowLongA" ( _
  ByVal hwnd As Long, _
  ByVal nIndex As Long, _
  ByVal dwNewLong As Long _
) As Long

Private Declare Function SetWindowPos _
Lib "user32" ( _
  ByVal hwnd As Long, _
  ByVal hWndInsertAfter As Long, _
  ByVal X As Long, _
  ByVal Y As Long, _
  ByVal cx As Long, _
  ByVal cy As Long, _
  ByVal wFlags As Long _
) As Long
' **************************************************

Function AccessTitleBar(Show As Boolean) As Long
  Dim hwnd As Long
  Dim nIndex As Long
  Dim dwNewLong As Long
  Dim dwLong As Long
  Dim wFlags As Long

  hwnd = hWndAccessApp
  nIndex = GWL_STYLE
  wFlags = SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED + SWP_NOMOVE

  dwLong = GetWindowLong(hwnd, nIndex)

  If Show Then
    dwNewLong = (dwLong Or WS_CAPTION)
  Else
    dwNewLong = (dwLong And Not WS_CAPTION)
  End If

  Call SetWindowLong(hwnd, nIndex, dwNewLong)
  Call SetWindowPos(hwnd, 0&, 0&, 0&, 0&, 0&, wFlags)
End Function


Function Buttons(Show As Boolean) As Long
  Dim hwnd As Long
  Dim nIndex As Long
  Dim dwNewLong As Long
  Dim dwLong As Long

  hwnd = hWndAccessApp
  nIndex = GWL_STYLE

  Const wFlags = SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED + SWP_NOMOVE
  Const FLAGS_COMBI = WS_MINIMIZEBOX Or WS_MAXIMIZEBOX Or WS_SYSMENU

  dwLong = GetWindowLong(hwnd, nIndex)

  If Show Then
    dwNewLong = (dwLong Or FLAGS_COMBI)
  Else
    dwNewLong = (dwLong And Not FLAGS_COMBI)
  End If

  Call SetWindowLong(hwnd, nIndex, dwNewLong)
  Call SetWindowPos(hwnd, 0&, 0&, 0&, 0&, 0&, wFlags)
End Function
' ********** Code End *************


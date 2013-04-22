Attribute VB_Name = "modWinSizer"
Option Compare Database
Option Explicit

'************ Code Start **********
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3


Private Declare Function apiShowWindow Lib "user32" _
    Alias "ShowWindow" (ByVal hwnd As Long, _
          ByVal nCmdShow As Long) As Long

Function fSetAccessWindow(nCmdShow As Long)
'Usage Examples
'Maximize window:
'       ?fSetAccessWindow(SW_SHOWMAXIMIZED)
'Minimize window:
'       ?fSetAccessWindow(SW_SHOWMINIMIZED)
'Hide window:
'       ?fSetAccessWindow(SW_HIDE)
'Normal window:
'       ?fSetAccessWindow(SW_SHOWNORMAL)
'
Dim loX  As Long
Dim loForm As Form
    On Error Resume Next
    Set loForm = Screen.ActiveForm
    If Err <> 0 Then 'no Activeform
      If nCmdShow = SW_HIDE Then
        MsgBox "Cannot hide Access unless " _
                    & "a form is on screen"
      Else
        loX = apiShowWindow(hWndAccessApp, nCmdShow)
        Err.Clear
      End If
    Else
        If nCmdShow = SW_SHOWMINIMIZED And loForm.Modal = True Then
            MsgBox "Cannot minimize Access with " _
                    & (loForm.Caption + " ") _
                    & "form on screen"
        ElseIf nCmdShow = SW_HIDE And loForm.PopUp <> True Then
            MsgBox "Cannot hide Access with " _
                    & (loForm.Caption + " ") _
                    & "form on screen"
        Else
            loX = apiShowWindow(hWndAccessApp, nCmdShow)
        End If
    End If
    fSetAccessWindow = (loX <> 0)
End Function

'************ Code End **********


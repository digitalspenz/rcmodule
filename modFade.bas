Attribute VB_Name = "modFade"
Option Compare Database
Option Explicit

'Attribute VB_Name = "basFadeForm"
'Author: © Copyright 2005 Graham R Seach
' gseach@accessmvp.com
'
'Description: This function changes a form's opacity, so it
' selectively fades in or out, or is set to a specific
' opacity level.
'
'Dependencies: Microsoft Access 2000/2002/2003 only.
'
'Inputs: frm:           The form object to fade.
'        Direction:     Integer value (from the FadeDirection enum), indicating
'                       the direction of fade.
'        iDelay:        (Optional) Integer value indicating the delay (in milliseconds)
'                       between each fade step.
'                       - Defaults to zero.
'        StartOpacity:  (Optional) Long value indicating the opacity level at
'                       which to start the fade (between 1 and 255).
'                       - Defaults to 5.
'
'Outputs: None.
'Call example:
'        Private Sub Form_Close()
'        'Fade the form out
'            FadeForm Me, Fadeout, 2, 255
'
'        End Sub
'
'        Private Sub Form_Open(Cancel As Integer)
'        'Fade the form in
'            FadeForm Me, FadeIn, 20, 5
'
'        End Sub

Public Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    
Public Declare Function SetWindowLong Lib "user32" Alias _
    "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
    
Public Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, _
     ByVal dwFlags As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&

Public Enum FadeDirection
    FadeIn = -1
    Fadeout = 0
    SetOpacity = 1
End Enum

Public Sub FadeForm(frm As Form, _
        Optional Direction As FadeDirection = FadeDirection.FadeIn, _
        Optional iDelay As Integer = 0, _
        Optional StartOpacity As Long = 5)
    
    Dim lOriginalStyle As Long
    Dim ictr As Integer
    Const OFFSET As Integer = 20
    
    'You can only set a form's opacity if it's Popup property = True
    If (frm.PopUp = True) Then
        'Remember the original form (window) style
        lOriginalStyle = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
        SetWindowLong frm.hwnd, GWL_EXSTYLE, lOriginalStyle Or WS_EX_LAYERED
        
        If (lOriginalStyle = 0) And (Direction <> FadeDirection.SetOpacity) Then
            FadeForm frm, SetOpacity, , StartOpacity
        Else
            frm.Visible = True
        End If
        
        Select Case Direction
            Case FadeDirection.FadeIn
                If StartOpacity < 1 Then StartOpacity = 1
                
                For ictr = StartOpacity To 255 Step OFFSET
                    SetLayeredWindowAttributes frm.hwnd, 0, CByte(ictr), LWA_ALPHA
                    DoEvents
                    Sleep iDelay
                Next
                
                SetLayeredWindowAttributes frm.hwnd, 0, CByte(255), LWA_ALPHA
            Case FadeDirection.Fadeout
                If StartOpacity < 6 Then StartOpacity = 255
                
                For ictr = StartOpacity To 1 Step -OFFSET
                    SetLayeredWindowAttributes frm.hwnd, 0, CByte(ictr), LWA_ALPHA
                    DoEvents
                    Sleep iDelay
                Next
            Case Else 'FadeDirection.SetOpacity
                Select Case StartOpacity
                    Case Is < OFFSET: StartOpacity = 1
                    Case Is > 255: StartOpacity = 255
                End Select
                                
                SetLayeredWindowAttributes frm.hwnd, 0, CByte(StartOpacity), LWA_ALPHA
                DoEvents
                'frm.SetFocus
        End Select
    Else
        DoCmd.Beep
        MsgBox "The form's Popup property must be set to True.", _
            vbOKOnly + vbInformation, "Cannot fade form"
    End If
End Sub




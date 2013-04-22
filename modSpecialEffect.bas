Attribute VB_Name = "modSpecialEffect"
Option Compare Database
Option Explicit

Private Const conWhite = 16777215
Private Const conGray = -12632256
Private Const conIndent = 2
Private Const conFlat = 0

Public Function SpecialEffectEnter()
On Error GoTo HandleErr

    ' Set the current control to be indented.
    Screen.ActiveControl.SpecialEffect = conIndent
    
    ' Set the current control's background color to be white.
    Screen.ActiveControl.BackColor = conWhite
    

ExitHere:
    Exit Function

HandleErr:
    MsgBox Err & ": " & Err.Description
    Resume ExitHere
End Function



Public Function SpecialEffectExit()
On Error GoTo HandleErr

    ' Set the current control to be flat.
    Screen.ActiveControl.SpecialEffect = conFlat
    
    ' Set the current control's background color to be gray.
    Screen.ActiveControl.BackColor = conGray

ExitHere:
    Exit Function

HandleErr:
    MsgBox Err & ": " & Err.Description
    Resume ExitHere
End Function

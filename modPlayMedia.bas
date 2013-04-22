Attribute VB_Name = "modPlayMedia"
Option Compare Database
Option Explicit

'****************** Code Start *********************
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
Public Const pcsSYNC = 0      ' wait until sound is finished playing
Public Const pcsASYNC = 1     ' don't wait for finish
Public Const pcsNODEFAULT = 2 ' play no default sound if sound doesn't exist
Public Const pcsLOOP = 8      ' play sound in an infinite loop (until next apiPlaySound)
Public Const pcsNOSTOP = 16   ' don't interrupt a playing sound

'Sound APIs
Private Declare Function apiPlaySound Lib "Winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'AVI APIs
Private Declare Function apimciSendString Lib "Winmm.dll" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function apimciGetErrorString Lib "Winmm.dll" _
    Alias "mciGetErrorStringA" (ByVal dwError As Long, _
    ByVal lpstrBuffer As String, ByVal uLength As Long) As Long


Function fPlayStuff(ByVal strFileName As String, _
            Optional intPlayMode As Integer) As Long
'MUST pass a filename _with_ extension
'Supports Wav, AVI, MID type files
Dim lngRet As Long
Dim strTemp As String

    Select Case LCase(fGetFileExt(strFileName))
        Case "wav":
            If Not IsMissing(intPlayMode) Then
                lngRet = apiPlaySound(strFileName, intPlayMode)
            Else
                MsgBox "Must specify play mode."
                Exit Function
            End If
        Case "avi", "mid":
            strTemp = String$(256, 0)
            lngRet = apimciSendString("play " & strFileName, strTemp, 255, 0)
    End Select
    fPlayStuff = lngRet
End Function
Function fStopStuff(ByVal strFileName As String)
'Stops a multimedia playback
Dim lngRet As Long
Dim strTemp As String
    Select Case LCase(fGetFileExt(strFileName))
        Case "Wav":
            lngRet = apiPlaySound(0, pcsASYNC)
        Case "avi", "mid":
            strTemp = String$(256, 0)
            lngRet = apimciSendString("stop " & strFileName, strTemp, 255, 0)
    End Select
    fStopStuff = lngRet
End Function

Private Function fGetFileExt(ByVal strFullPath As String) As String
Dim intPos As Integer, intLen As Integer
    intLen = Len(strFullPath)
    If intLen Then
        For intPos = intLen To 1 Step -1
            'Find the last \
            If Mid$(strFullPath, intPos, 1) = "." Then
                fGetFileExt = Mid$(strFullPath, intPos + 1)
                Exit Function
            End If
        Next intPos
    End If
End Function

Function fGetError(ByVal lngErrNum As Long) As String
    ' Translate the error code to a string
    Dim lngX As Long
    Dim strErr As String

    strErr = String$(256, 0)
    lngX = apimciGetErrorString(lngErrNum, strErr, 255)
    strErr = Left$(strErr, Len(strErr) - 1)
    fGetError = strErr
End Function
Function fatest()
Dim a As Long
    a = fPlayStuff("C:\winnt\clock.avi")
    'a = fStopStuff("C:\winnt\clock.avi")
End Function

'****************** Code End *********************


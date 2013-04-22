Attribute VB_Name = "modGetFile"
Option Compare Database
Option Explicit

'***************** Code Start ***************
'This code was originally written by Dev Ashish.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish

'Now from your code, call the function as
'strFilePath = fReturnFilePath("network.wri", "C:")
'(Note: Depending on whether you have Microsoft Office's FindFast turned on or not, the search might take a few seconds.)

Private Declare Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
        "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath _
        As String, ByVal lpOutPath As String) As Long

Function fSearchFile(ByVal strFileName As String, _
            ByVal strSearchPath As String) As String
'Returns the first match found
    Dim lpBuffer As String
    Dim lngResult As Long
    fSearchFile = ""
    lpBuffer = String$(1024, 0)
    lngResult = apiSearchTreeForFile(strSearchPath, strFileName, lpBuffer)
    If lngResult <> 0 Then
        If InStr(lpBuffer, vbNullChar) > 0 Then
            fSearchFile = Left$(lpBuffer, InStr(lpBuffer, vbNullChar) - 1)
        End If
    End If
End Function
'******************** Code End ****************

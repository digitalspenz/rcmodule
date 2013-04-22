Attribute VB_Name = "modCopyFile"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : CopyFile
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Copy a file
'             Overwrites existing copy without prompting
'             Cannot copy locked files (currently in use)
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strSource - Path/Name of the file to be copied
' strDest - Path/Name for copying the file to
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 1         2007-Apr-01             Initial Release
'---------------------------------------------------------------------------------------
Function CopyFile(strSource As String, strDest As String) As Boolean
On Error GoTo CopyFile_Error
 
    FileCopy strSource, strDest
    CopyFile = True
    Exit Function
 
CopyFile_Error:
    If Err.Number = 0 Then
    ElseIf Err.Number = 70 Then
        MsgBox "The file is currently in use and therfore is locked and cannot be copied at this" & _
               " time.  Please ensure that no one is using the file and try again.", vbOKOnly, _
               "File Currently in Use"
    ElseIf Err.Number = 53 Then
        MsgBox "The Source File '" & strSource & "' could not be found.  Please validate the" & _
               " location and name of the specifed Source File and try again", vbOKOnly, _
               "File Currently in Use"
    Else
        MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
               Err.Number & vbCrLf & "Error Source: ModExtFiles / CopyFile" & vbCrLf & _
               "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    End If
    Exit Function
End Function


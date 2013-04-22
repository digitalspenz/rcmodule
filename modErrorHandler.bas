Attribute VB_Name = "modErrorHandler"
Option Compare Database
Option Explicit

Public Sub DisplayUnexpectedError(ErrorNumber As String, ErrorDescription As String)

   MsgBox "Sorry, nagkaroon po ng hindi inaasahang error." _
   & vbCrLf & "Ang impormasyon tungkol sa error ay:" _
   & vbCrLf & vbCrLf & "Error Number: " & ErrorNumber & ", " & ErrorDescription, vbCritical, gstrTitle
   
'   MsgBox "An error has occured. Please contact your technical support and tell them this information:" _
'   & vbCrLf & vbCrLf & "Error Number: " & ErrorNumber & ", " & ErrorDescription, vbCritical, "Unexpected Error"
   
End Sub


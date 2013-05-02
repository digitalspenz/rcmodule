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


'--this will be the updated error handling routine that will be implemented in the future
'On Error GoTo Error_Handler
    'Your code will go here
    
'Error_Handler_Exit:
 ''   On Error Resume Next
   '' Exit {PROCEDURE_TYPE}
 
'Error_Handler:
    'MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
    ''       "Error Number: " & Err.Number & vbCrLf & _
     ''      "Error Source: {PROCEDURE_NAME}/{MODULE_NAME}" & vbCrLf & _
     ''      "Error Description: " & Err.Description, vbCritical, _
     ''      "An Error has Occured!"
  ''  Resume Error_Handler_Exit
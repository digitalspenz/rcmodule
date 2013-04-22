Attribute VB_Name = "modFindFieldUsedWhere"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : FindFieldUsedWhere
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Locate where a field is used within queries, forms and reports
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFieldName    : Field Name to search for in the various Access objects
'
' Usage:
' ~~~~~~
' FindFieldUsedWhere("Type A")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2011-04-08                 Initial Release
'---------------------------------------------------------------------------------------
Function FindFieldUsedWhere(sFieldName As String)
On Error GoTo Error_Handler
    Dim db               As DAO.Database
    Dim qdf              As DAO.QueryDef
    Dim sSQL             As String
    Dim ctl              As Control
    Dim frm              As AccessObject
    Dim DbO              As AccessObject
    Dim DbP              As Object
 
    Set db = CurrentDb
    Debug.Print "FindFieldUsedWhere Begin"
    Debug.Print "Searching for '" & sFieldName & "'"
    Debug.Print "================================================================================"
 
    'Check Queries
    For Each qdf In db.QueryDefs
        'qdf.Name    'The current query's name
        'qdf.SQL     'The current query's SQL statement
        sSQL = qdf.SQL
        If InStr(sSQL, sFieldName) Then
            'The Query is a Make Table Query and has our TableName we are looking for
            Debug.Print "Query: " & qdf.Name
        End If
    Next
 
    'Check Forms
    For Each frm In CurrentProject.AllForms
        DoCmd.OpenForm frm.Name, acDesign
        If InStr(Forms(frm.Name).RecordSource, sFieldName) Then
            'The Query is a Make Table Query and has our TableName we are looking for
            Debug.Print "Form: " & frm.Name
        End If
        'Loop throught the Form Controls
        For Each ctl In Forms(frm.Name).Form.Controls
            Select Case ctl.ControlType
                Case acComboBox
                    If Len(ctl.Tag) > 0 Then
                        If InStr(ctl.Tag, sFieldName) Then
                            'The Query is a Make Table Query and has our TableName we are looking for
                            Debug.Print "Form: " & frm.Name & " :: Control: " & ctl.Name
                        End If
                        If InStr(ctl.ControlSource, sFieldName) Then
                            'The Query is a Make Table Query and has our TableName we are looking for
                            Debug.Print "Form: " & frm.Name & " :: Control: " & ctl.Name
                        End If
                    End If
                Case acTextBox, acCheckBox
                    If InStr(ctl.ControlSource, sFieldName) Then
                        'The Query is a Make Table Query and has our TableName we are looking for
                        Debug.Print "Form: " & frm.Name & " :: Control: " & ctl.Name
                    End If
            End Select
        Next ctl
        DoCmd.Close acForm, frm.Name, acSaveNo
    Next frm
 
    'Check Reports
    Set DbP = Application.CurrentProject
    For Each DbO In DbP.AllReports
        DoCmd.OpenReport DbO.Name, acDesign
        If InStr(Reports(DbO.Name).RecordSource, sFieldName) Then
            'The Query is a Make Table Query and has our TableName we are looking for
            Debug.Print "Report: " & DbO.Name
        End If
        'Loop throught the Report Controls
        For Each ctl In Reports(DbO.Name).Report.Controls
            Select Case ctl.ControlType
                Case acComboBox
                    If Len(ctl.Tag) > 0 Then
                        If InStr(ctl.Tag, sFieldName) Then
                            'The Query is a Make Table Query and has our TableName we are looking for
                            Debug.Print "Report: " & DbO.Name & " :: Control: " & ctl.Name
                        End If
                        If InStr(ctl.ControlSource, sFieldName) Then
                            'The Query is a Make Table Query and has our TableName we are looking for
                            Debug.Print "Report: " & DbO.Name & " :: Control: " & ctl.Name
                        End If
                    End If
                Case acTextBox, acCheckBox
                    If InStr(ctl.ControlSource, sFieldName) Then
                        'The Query is a Make Table Query and has our TableName we are looking for
                        Debug.Print "Report: " & DbO.Name & " :: Control: " & ctl.Name
                    End If
            End Select
        Next ctl
        DoCmd.Close acReport, DbO.Name, acSaveNo
    Next DbO
 
    Debug.Print "================================================================================"
    Debug.Print "FindFieldUsedWhere End"
 
Error_Handler_Exit:
    On Error Resume Next
    Set qdf = Nothing
    Set db = Nothing
    Set ctl = Nothing
    Set frm = Nothing
    Set DbP = Nothing
    Set DbO = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: FindFieldUsedWhere" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occured"
    Resume Error_Handler_Exit
End Function


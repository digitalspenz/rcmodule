Attribute VB_Name = "tmodExportAllObj"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : ExpObj2ExtDb
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Export all the database object to another database
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sExtDb    : Fully qualified path and filename of the database to export the objects
'             to.
'
' Usage:
' ~~~~~~
' ExpObj2ExtDb "c:\databases\dbtest.accdb"
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2008-Sep-27                 Initial Release
'---------------------------------------------------------------------------------------
Public Sub ExpObj2ExtDb(sExtDb As String)
    On Error GoTo Error_Handler
    Dim qdf             As QueryDef
    Dim tdf             As TableDef
    Dim obj             As AccessObject
 
    ' Forms.
    For Each obj In CurrentProject.AllForms
        DoCmd.TransferDatabase acExport, "Microsoft Access", sExtDb, _
                               acForm, obj.Name, obj.Name, False
    Next obj
 
    ' Macros.
    For Each obj In CurrentProject.AllMacros
        DoCmd.TransferDatabase acExport, "Microsoft Access", sExtDb, _
                               acMacro, obj.Name, obj.Name, False
    Next obj
 
    ' Modules.
    For Each obj In CurrentProject.AllModules
        DoCmd.TransferDatabase acExport, "Microsoft Access", sExtDb, _
                               acModule, obj.Name, obj.Name, False
    Next obj
 
    ' Queries.
    For Each qdf In CurrentDb.QueryDefs
        If Left(qdf.Name, 1) <> "~" Then    'Ignore/Skip system generated queries
            DoCmd.TransferDatabase acExport, "Microsoft Access", sExtDb, _
                                   acQuery, qdf.Name, qdf.Name, False
        End If
    Next qdf
 
    ' Reports.
    For Each obj In CurrentProject.AllReports
        DoCmd.TransferDatabase acExport, "Microsoft Access", sExtDb, _
                               acReport, obj.Name, obj.Name, False
    Next obj
 
    ' Tables.
    For Each tdf In CurrentDb.TableDefs
        If Left(tdf.Name, 4) <> "MSys" Then    'Ignore/Skip system tables
            DoCmd.TransferDatabase acExport, "Microsoft Access", sExtDb, _
                                   acTable, tdf.Name, tdf.Name, False
        End If
    Next tdf
 
Error_Handler_Exit:
    On Error Resume Next
    Set qdf = Nothing
    Set tdf = Nothing
    Set obj = Nothing
    Exit Sub
 
Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Sub


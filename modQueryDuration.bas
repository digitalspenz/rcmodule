Attribute VB_Name = "modQueryDuration"
Option Compare Database
Option Explicit

Private Declare Function apiGetTickCount Lib "kernel32" _
 Alias "GetTickCount" () As Long

Public Function QueryDuration(ByVal pQueryName As String) As Long
    Dim db As DAO.Database
    Dim lngStart As Long
    Dim lngDone As Long
    Dim rs As DAO.Recordset

    Set db = CurrentDb()
    lngStart = apiGetTickCount() ' milliseconds '
    Set rs = db.OpenRecordset(pQueryName, dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
    End If
    lngDone = apiGetTickCount()
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    QueryDuration = lngDone - lngStart
End Function


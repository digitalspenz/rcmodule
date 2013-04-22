Attribute VB_Name = "modRunSum"
Option Compare Database
Option Explicit

Public Function frmRunSum(frm As Form, pkName As String, sumName As String)
Dim rst As DAO.Recordset, fld As DAO.Field, subTotal

Set rst = frm.RecordsetClone
Set fld = rst(sumName)

'Set starting point.
rst.FindFirst "[" & pkName & "] = " & frm(pkName)

'Running Sum (subTotal) for each record group occurs here.
'After the starting point is set, we sum backwards to record 1.
If Not rst.EOF Then
   Do Until rst.EOF
      subTotal = subTotal + Nz(fld, 0)
      rst.MoveNext
'      rst.MovePrevious
   Loop
Else
   subTotal = 0
End If

frmRunSum = subTotal
Set fld = Nothing
Set rst = Nothing


End Function

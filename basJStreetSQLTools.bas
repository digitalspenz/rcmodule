Attribute VB_Name = "basJStreetSQLTools"
'--------------------------------------------------------------------
'
' Copyright 1995-2009 J Street Technology, Inc.
' www.JStreetTech.com
'
' This code may be used and distributed as part of your application
' provided that all comments remain intact.
'
' J Street Technology offers this code "as is" and does not assume
' any liability for bugs or problems with any of the code.  In
' addition, we do not provide free technical support for this code.
'
'--------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Sub ParseSQL(varSQL As Variant, varInsert As Variant, varSELECT As Variant, varWhere As Variant, varOrderBy As Variant, varGROUPBY As Variant, varHAVING As Variant)
    On Error GoTo Error_Handler
    '12/4/1995  CM  Created
    '2/15/1996  AS  Converted to Sub from Function
    '10/27/1997 AS  Added GroupBy capability
    '04/22/2008 EI  Replaced CR and LF characters with spaces in input SQL
    '
    'This subroutine accepts a valid SQL string and passes back separated INSERT, SELECT, WHERE, ORDER BY and GROUP BY clauses.
    '
    'INPUT:
    '    varSQL      valid SQL string to parse
    'OUTPUT (BYREF):
    '    varInsert   INSERT portion of SQL
    '    varSELECT   SELECT portion of SQL (includes JOIN info)
    '    varWHERE    WHERE portion of SQL
    '    varORDERBY  ORDER BY portion of SQL
    '    varGROUPBY  GROUP BY portion of SQL
    '    varHAVING   HAVING portion of SQL
    '
    'Note: While the subroutine will accept the ';' character in varSQL,
    '      there is no ';' character passed back at any time.
    '
    Dim intStartINSERT As Integer
    Dim intStartSELECT As Integer
    Dim intStartWHERE As Integer
    Dim intStartORDERBY As Integer
    Dim intStartGROUPBY As Integer
    Dim intStartHAVING As Integer
    
    Dim intLenINSERT As Integer
    Dim intLenSELECT As Integer
    Dim intLenWHERE As Integer
    Dim intLenORDERBY As Integer
    Dim intLenGROUPBY As Integer
    Dim intLenHAVING As Integer
        
    Dim intLenSQL As Integer
    
    'Remove CR and LF characters from SQL, replace with spaces
    varSQL = Replace(varSQL, vbCr, " ")
    varSQL = Replace(varSQL, vbLf, " ")
    
    intStartINSERT = InStr(varSQL, "INSERT ")
    intStartSELECT = InStr(varSQL, "SELECT ")
    intStartWHERE = InStr(varSQL, "WHERE ")
    intStartORDERBY = InStr(varSQL, "ORDER BY ")
    intStartGROUPBY = InStr(varSQL, "GROUP BY ")
    intStartHAVING = InStr(varSQL, "HAVING ")
    
    'if there's no GROUP BY, there can't be a HAVING
    If intStartGROUPBY = 0 Then
        intStartHAVING = 0
    End If
    
    If InStr(varSQL, ";") Then       'if it exists, trim off the ';'
        varSQL = Left(varSQL, InStr(varSQL, ";") - 1)
    End If

    intLenSQL = Len(varSQL)

    ' find length of Insert portion
    If intStartINSERT > 0 Then
        ' start with longest it could be
        intLenINSERT = intLenSQL - intStartINSERT + 1
        If intStartSELECT > 0 And intStartSELECT > intStartINSERT And intStartSELECT < intStartINSERT + intLenINSERT Then
            'we found a new portion closer to this one
            intLenINSERT = intStartSELECT - intStartINSERT
        End If
        If intStartWHERE > 0 And intStartWHERE > intStartINSERT And intStartWHERE < intStartINSERT + intLenINSERT Then
            'we found a new portion closer to this one
            intLenINSERT = intStartWHERE - intStartINSERT
        End If
        If intStartORDERBY > 0 And intStartORDERBY > intStartINSERT And intStartORDERBY < intStartINSERT + intLenINSERT Then
            'we found a new portion closer to this one
            intLenINSERT = intStartORDERBY - intStartINSERT
        End If
        If intStartGROUPBY > 0 And intStartGROUPBY > intStartINSERT And intStartGROUPBY < intStartINSERT + intLenINSERT Then
            'we found a new portion closer to this one
            intLenINSERT = intStartGROUPBY - intStartINSERT
        End If
        If intStartHAVING > 0 And intStartHAVING > intStartINSERT And intStartHAVING < intStartINSERT + intLenINSERT Then
            'we found a new portion closer to this one
            intLenINSERT = intStartHAVING - intStartINSERT
        End If
    End If

    ' find length of Select portion
    If intStartSELECT > 0 Then
        ' start with longest it could be
        intLenSELECT = intLenSQL - intStartSELECT + 1
        If intStartWHERE > 0 And intStartWHERE > intStartSELECT And intStartWHERE < intStartSELECT + intLenSELECT Then
            'we found a new portion closer to this one
            intLenSELECT = intStartWHERE - intStartSELECT
        End If
        If intStartORDERBY > 0 And intStartORDERBY > intStartSELECT And intStartORDERBY < intStartSELECT + intLenSELECT Then
            'we found a new portion closer to this one
            intLenSELECT = intStartORDERBY - intStartSELECT
        End If
        If intStartGROUPBY > 0 And intStartGROUPBY > intStartSELECT And intStartGROUPBY < intStartSELECT + intLenSELECT Then
            'we found a new portion closer to this one
            intLenSELECT = intStartGROUPBY - intStartSELECT
        End If
        If intStartHAVING > 0 And intStartHAVING > intStartSELECT And intStartHAVING < intStartSELECT + intLenSELECT Then
            'we found a new portion closer to this one
            intLenSELECT = intStartHAVING - intStartSELECT
        End If
    End If

    ' find length of GROUPBY portion
    If intStartGROUPBY > 0 Then
        ' start with longest it could be
        intLenGROUPBY = intLenSQL - intStartGROUPBY + 1
        If intStartWHERE > 0 And intStartWHERE > intStartGROUPBY And intStartWHERE < intStartGROUPBY + intLenGROUPBY Then
            'we found a new portion closer to this one
            intLenGROUPBY = intStartWHERE - intStartGROUPBY
        End If
        If intStartORDERBY > 0 And intStartORDERBY > intStartGROUPBY And intStartORDERBY < intStartGROUPBY + intLenGROUPBY Then
            'we found a new portion closer to this one
            intLenGROUPBY = intStartORDERBY - intStartGROUPBY
        End If
        If intStartHAVING > 0 And intStartHAVING > intStartGROUPBY And intStartHAVING < intStartGROUPBY + intLenGROUPBY Then
            'we found a new portion closer to this one
            intLenGROUPBY = intStartHAVING - intStartGROUPBY
        End If
    End If
    
    ' find length of HAVING portion
    If intStartHAVING > 0 Then
        ' start with longest it could be
        intLenHAVING = intLenSQL - intStartHAVING + 1
        If intStartWHERE > 0 And intStartWHERE > intStartHAVING And intStartWHERE < intStartHAVING + intLenHAVING Then
            'we found a new portion closer to this one
            intLenHAVING = intStartWHERE - intStartHAVING
        End If
        If intStartORDERBY > 0 And intStartORDERBY > intStartHAVING And intStartORDERBY < intStartHAVING + intLenHAVING Then
            'we found a new portion closer to this one
            intLenHAVING = intStartORDERBY - intStartHAVING
        End If
        If intStartGROUPBY > 0 And intStartGROUPBY > intStartHAVING And intStartGROUPBY < intStartHAVING + intLenHAVING Then
            'we found a new portion closer to this one
            intLenHAVING = intStartGROUPBY - intStartHAVING
        End If
    End If
    
    
    ' find length of ORDERBY portion
    If intStartORDERBY > 0 Then
        ' start with longest it could be
        intLenORDERBY = intLenSQL - intStartORDERBY + 1
        If intStartWHERE > 0 And intStartWHERE > intStartORDERBY And intStartWHERE < intStartORDERBY + intLenORDERBY Then
            'we found a new portion closer to this one
            intLenORDERBY = intStartWHERE - intStartORDERBY
        End If
        If intStartGROUPBY > 0 And intStartGROUPBY > intStartORDERBY And intStartGROUPBY < intStartORDERBY + intLenORDERBY Then
            'we found a new portion closer to this one
            intLenORDERBY = intStartGROUPBY - intStartORDERBY
        End If
        If intStartHAVING > 0 And intStartHAVING > intStartORDERBY And intStartHAVING < intStartORDERBY + intLenORDERBY Then
            'we found a new portion closer to this one
            intLenORDERBY = intStartHAVING - intStartORDERBY
        End If
    End If
    
    ' find length of WHERE portion
    If intStartWHERE > 0 Then
        ' start with longest it could be
        intLenWHERE = intLenSQL - intStartWHERE + 1
        If intStartGROUPBY > 0 And intStartGROUPBY > intStartWHERE And intStartGROUPBY < intStartWHERE + intLenWHERE Then
            'we found a new portion closer to this one
            intLenWHERE = intStartGROUPBY - intStartWHERE
        End If
        If intStartORDERBY > 0 And intStartORDERBY > intStartWHERE And intStartORDERBY < intStartWHERE + intLenWHERE Then
            'we found a new portion closer to this one
            intLenWHERE = intStartORDERBY - intStartWHERE
        End If
        If intStartHAVING > 0 And intStartHAVING > intStartWHERE And intStartHAVING < intStartWHERE + intLenWHERE Then
            'we found a new portion closer to this one
            intLenWHERE = intStartHAVING - intStartWHERE
        End If
    End If

    ' set each output portion
    If intStartINSERT > 0 Then
        varInsert = Mid$(varSQL, intStartINSERT, intLenINSERT)
    End If
    If intStartSELECT > 0 Then
        varSELECT = Mid$(varSQL, intStartSELECT, intLenSELECT)
    End If
    If intStartGROUPBY > 0 Then
        varGROUPBY = Mid$(varSQL, intStartGROUPBY, intLenGROUPBY)
    End If
    If intStartHAVING > 0 Then
        varHAVING = Mid$(varSQL, intStartHAVING, intLenHAVING)
    End If
    If intStartORDERBY > 0 Then
        varOrderBy = Mid$(varSQL, intStartORDERBY, intLenORDERBY)
    End If
    If intStartWHERE > 0 Then
        varWhere = Mid$(varSQL, intStartWHERE, intLenWHERE)
    End If

Exit_Procedure:
    Exit Sub
Error_Handler:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & ", " & Err.Description
            Resume Exit_Procedure
            Resume
    End Select
End Sub


Public Function ReplaceWhereClause(varSQL As Variant, varNewWHERE As Variant)
On Error GoTo Error_Handler
'2/15/96  AS    Created
'
'This subroutine accepts a valid SQL string and Where clause, and
'returns the same SQL statement with the original Where clause (if any)
'replaced by the passed in Where clause.
'
'INPUT:
'    varSQL      valid SQL string to change
'OUTPUT:
'    strNewWHERE New WHERE clause to insert into SQL statement
'
'To use it, try the function ReplaceWhereClause.
'You send in a whole SQL statement and the new desired Where clause, and it locates and
'snips out the old one, inserts your new one, and gives you back the
'new statement. If you send in a null or empty Where clause, the
'function just removes any existing one from the statement.

Dim strInsert As String
Dim strSELECT As String
Dim strWhere As String
Dim strOrderBy As String
Dim strGROUPBY As String
Dim strHAVING As String

Call ParseSQL(varSQL, strInsert, strSELECT, strWhere, strOrderBy, strGROUPBY, strHAVING)

ReplaceWhereClause = strInsert & " " & strSELECT & " " & varNewWHERE & " " & strGROUPBY & " " & strHAVING & " " & strOrderBy

Exit_Procedure:
    Exit Function
Error_Handler:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & ", " & Err.Description
            Resume Exit_Procedure
            Resume
    End Select
End Function

Public Function ReplaceHavingClause(varSQL As Variant, varNewHAVING As Variant)
On Error GoTo Error_Handler
'
'This subroutine accepts a valid SQL string and Having clause, and
'returns the same SQL statement with the original Having clause (if any)
'replaced by the passed in Having clause.
'
'INPUT:
'    varSQL      valid SQL string to change
'OUTPUT:
'    strNewHaving New Having clause to insert into SQL statement
'
Dim strInsert As String
Dim strSELECT As String
Dim strWhere As String
Dim strOrderBy As String
Dim strGROUPBY As String
Dim strHAVING As String

Call ParseSQL(varSQL, strInsert, strSELECT, strWhere, strOrderBy, strGROUPBY, strHAVING)

ReplaceHavingClause = strInsert & " " & strSELECT & " " & strWhere & " " & strGROUPBY & " " & varNewHAVING & " " & strOrderBy

Exit_Procedure:
    Exit Function
Error_Handler:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & ", " & Err.Description
            Resume Exit_Procedure
            Resume
    End Select
End Function

Public Function ModifyWhereClause(varSQL As Variant, varNewWHERE As Variant)
On Error GoTo Error_Handler
'5/27/98  MT    Created
'7/9/98   AS&SL Modified to handle empty new where clause
'
'This subroutine accepts a valid SQL string and Where clause, and
'returns the same SQL statement with the passed in Where clause
'appended to the original Where clause (if any) with an AND operator.
'
'INPUT:
'    varSQL      valid SQL string to change
'
'OUTPUT:
'    strNewWHERE New WHERE clause to insert into SQL statement
'
Dim strInsert As String
Dim strSELECT As String
Dim strWhere As String
Dim strOrderBy As String
Dim strGROUPBY As String
Dim strHAVING As String
Dim strModWHERE As String
Dim intStartWHERE As Integer

Call ParseSQL(varSQL, strInsert, strSELECT, strWhere, strOrderBy, strGROUPBY, strHAVING)

If varNewWHERE & "" = "" Then
    ModifyWhereClause = varSQL
Else
    If strWhere = "" Then
        ModifyWhereClause = strInsert & " " & strSELECT & " " & varNewWHERE & " " & strGROUPBY & " " & strHAVING & " " & strOrderBy
    Else
        'strip off WHERE from strNewWHERE
        strModWHERE = Mid$(varNewWHERE, 6)
        ModifyWhereClause = strInsert & " " & strSELECT & " " & strWhere & " AND " & strModWHERE & " " & strGROUPBY & " " & strHAVING & " " & strOrderBy
    End If
End If
      
Exit_Procedure:
    Exit Function
Error_Handler:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & ", " & Err.Description
            Resume Exit_Procedure
            Resume
    End Select
End Function

Public Function ReplaceOrderByClause(varSQL As Variant, varNewOrderBy As Variant)
On Error GoTo Error_Handler
    '
    'This subroutine accepts a valid SQL string and Where clause, and
    'returns the same SQL statement with the original Where clause (if any)
    'replaced by the passed in Where clause.
    '
    'INPUT:
    '    varSQL      valid SQL string to change
    'OUTPUT:
    '    strNewOrderBy New OrderBy clause to insert into SQL statement
    '
    Dim strInsert As String
    Dim strSELECT As String
    Dim strWhere As String
    Dim strOrderBy As String
    Dim strGROUPBY As String
    Dim strHAVING As String
    
    Call ParseSQL(varSQL, strInsert, strSELECT, strWhere, strOrderBy, strGROUPBY, strHAVING)
    
    ReplaceOrderByClause = strInsert & " " & strSELECT & " " & strWhere & " " & strGROUPBY & " " & strHAVING & " " & varNewOrderBy
    
Exit_Procedure:
    Exit Function
Error_Handler:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & ", " & Err.Description
            Resume Exit_Procedure
            Resume
    End Select
End Function


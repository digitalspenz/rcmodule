Attribute VB_Name = "modBuildWHERE"
Option Compare Database
Option Explicit

'*******************************************
'*  Constant Declarations for this module  *
'*******************************************

    Const mstrcTagID As String = "Where="               'Tag Property identifier to build Where Clause
    Const mstrcRange_Begin As String = "_BeginR"        'String to be appended to the base name of a control that represents the Begin part of a range
    Const mstrcRange_End As String = "_EndR"            'String to be appended to the base name of a control that represents the End part of a range
    Const mstrcDateSymbol = "#"                         'Symbol for dates (i.e. Date = #10/24/2004#)
    Const mstrcFldSeparator = ","                       'Separates items with Where= Tag
    Const mstrcStringSymbol = "'"                       'Symbol for strings (i.e. where name = 'dog')
    Const mstrcTagSeparator = ";"                       'Separates Tag items in controls Tag Property
    
'+********************************************************************************************
'*
'$  Function:   BuildWhere
'*
'*  Author:     FancyPrairie
'*
'*  Date:       October, 1998
'*
'*  Purpose:    This routine will build a Where clause based on the items the user selected/entered on a
'*              Report Criteria form (or any type of form that the caller may use to build a Where clause
'*              based on the items selected).
'*
'*              This routine assumes that, in order for the control to be included in the criteria,
'*              it must satisfy the following conditions:
'*
'*                  1. Tag Property must contain:  Where=TableName.FieldName,FieldType[,Operator,Value]
'*                         a. TableName.FieldName is the name of the Table and field to filter on
'*                            (example: tblEmployee.lngEmpID)
'*                         b. FieldType must be one of the following words: String, Date
'*                            (Note: String and Date are the only ones that have special meaning at this time
'*                                   (' or #).  However, if the FieldType is not String or Date then set
'*                                   you can leave the FieldType empty or (anticipating future uses) set the FieldType
'*                                   to Long, Integer, Byte, Single, Double, Boolean.)
'*                         c. Operator (optional) can be one of the following: = <> < > <= >= Like IsNull Between.  Or it could specify a control on your form (i.e. forms!formName!YourControl)
'*                            (Default is =)
'*                         d. Value (optional) to be filtered on (more for option groups and check boxes)
'*                            (Note: This is primarily intended to be used by Option groups and check boxes.
'*                                   For example, suppose you have a check box.  If the box is checked, then
'*                                   the you need to set the Value argument equal to True and the FieldType
'*                                   should be set to Boolean.  Else Value should be set to False.)
'*                  2. Control must be Enabled
'*                  3. Control must be visible
'*                  4. If the control is part of a range, then the name of the control must end with
'*                     either _BeginR (begin range control) or _EndR (end range control)
'*
'*              As of this writing, the following controls are checked:
'*                  1. List Boxes (multiselect and single select)
'*                  2. Ranges (2 controls that are grouped together (i.e. Begin and End Dates)
'*                  3. Text Boxes
'*                  4. Combo Boxes
'*                  5. Option Groups
'*                  6. Check Boxes
'*
'*  Arguments:  frm (form)
'*              ----------
'*              Represents the form (Report Criteria form) that contains the controls
'*              from which the Where clause is to be created.
'*
'*              varCtl (ParamArray)
'*              -------------------
'*              Indicates which controls you want checked.  If this argument is missing then
'*              this routine loops thru all of the controls on the form searching for the ones
'*              whose tag property is set accordingly.  Else it only checks the ones passed by the caller.
'*
'*              NOTE..NOTE..NOTE..NOTE..NOTE..NOTE..NOTE..NOTE..NOTE..NOTE..NOTE..NOTE..NOTE..NOTE..NOTE
'*
'*              This routine assumes that the Where clause will include the AND statement between each set of statements.
'*              For Example, (strLastName = 'Smith') AND (strFirstName = 'Justin')  However, there are situations where
'*              you may want the Where clause to include the OR statement between each set of statements.  For example,
'*              (strLastName = 'Smith') OR (strFirstName = 'Justin')  In this case you need to tell this routine to use
'*              OR rather than AND.  Therefore, the first argument in this ParamArray should be " OR ".
'*
'*  Control Examples:
'*
'*              ListBox
'*              -------
'*                Visible ........ Yes
'*                Enabled ........ Yes
'*                Multi Select ... None, Simple, or Extended
'*                Tag ............ Where=tblEmployee.lngDepartmentID,Number;
'*
'*              TextBox
'*              -------
'*                Visible ........ Yes
'*                Enabled ........ Yes
'*                Tag ............ Where=tblEmployee.dteBirth,Date,>;
'*
'*              Range
'*              -----
'*                Text Box (Begin Date Range: Name must end with _BeginR)
'*                  Name ...... txtHireDate_BeginR
'*                  Visible ... Yes
'*                  Enabled ... Yes
'*                  Tag ....... Where=tblEmployee.dteHire,Date;
'*
'*                Text Box (End Date Range: Name must end with _EndR)
'*                  Name ...... txtHireDate_EndR
'*                  Visible ... Yes
'*                  Enabled ... Yes
'*                  Tag ....... (not used...function relies on tag property of BeginDateRange text box)
'*
'*              Option Group
'*              ------------
'*                  Visible ... Yes
'*                  Enabled ... Yes
'*                  Tag ....... Where=tblEmployee.lngID,Long,=,3;
'*
'*              Check Boxes
'*              -----------
'*                  Visible ... Yes
'*                  Enabled ... Yes
'*                  Tag ....... Where=tblEmployee.ysnActive,Boolean,=,True;
'*
'*  Calling Example:
'*
'*              Docmd.OpenReport "rptYourReport", acViewPreview, ,BuildWhere(Me)
'*
'*        or    strWhere = BuildWhere(Me)
'*
'*        or    Docmd.OpenReport "rptYourReport", acViewPreview, ,BuildWhere(Me, " OR ")
'*
'*        or    strWhere = BuildWhere(Me, " OR ")
'*
'-*********************************************************************************************************'
'
Function BuildWhere(frm As Form, _
         ParamArray varCtl() As Variant)
             
'********************************
'*  Declaration Specifications  *
'********************************

    Dim ctl As Control                  'Control currently being processed
    Dim ctlEndR As Control              '2nd control of Range pair
    
    Dim varItem As Variant              'Items within multiselect list box
    
    Dim strAnd As String                'And or ""
    Dim strAndOr As String              'Either "And" or "Or"
    Dim strCtlType As String            'Control Type of control being processed (also see error handler)
    Dim strFieldName As String          'Table.FieldName value
    Dim strFieldType As String          'FieldType ' or #
    Dim strFieldValue As String         'Value
    Dim strOperator As String           'Operator (= <> > < etc)
    Dim strOpt As String                'Operator between controls (AND or OR)
    Dim strSuffix As String             'Suffix to be appended at end of a string
    Dim strWhere As String              'Where clause to be returned to caller
    
    Dim i As Integer                    'Working Variable
    Dim k As Integer                    'Working Variable
    
    Dim bolAll As Boolean               'True if All controls on form to be processed
    
'****************
'*  Initialize  *
'****************

    On Error GoTo Errhandler

    strWhere = vbNullString
    strAnd = vbNullString
    
'**************************************
'*  Begin loop thru controls on form  *
'**************************************

    
    If (UBound(varCtl) >= 0) Then
        If (varCtl(0) = " OR ") Then strOpt = " OR " Else strOpt = " AND "
        If (UBound(varCtl) = 0) And ((varCtl(0) = " OR ") Or (varCtl(0) = " AND ")) Then bolAll = True Else bolAll = False
    Else
        strOpt = " AND "
        bolAll = True
    End If
    
    If (bolAll) Then
        For Each ctl In frm.Controls
            GoSub CreateWhere
        Next
    Else
        If (varCtl(0) = " OR ") Or (varCtl(0) = " AND ") Then k = 1 Else k = 0
        For i = k To UBound(varCtl)
            Set ctl = varCtl(i)
            GoSub CreateWhere
        Next i
    End If

'***********************
'*  Save Where Clause  *
'***********************

    BuildWhere = strWhere
    
'********************
'*  Exit Procedure  *
'********************
        
ExitProcedure:

    Exit Function

'*********************************************
'*  Create Where Clause for current control  *
'*********************************************

CreateWhere:

    strCtlType = BuildWhere_ControlType(frm, ctl)
    
    Select Case strCtlType
            
        Case "Range":       GoSub GetTag: GoSub BuildRange   'strWhere = strWhere & strAnd & BuildWhere_Range(frm, ctl)
        Case "ListBox":     GoSub GetTag: GoSub BuildListBox   'strWhere = strWhere & strAnd & BuildWhere_ListBox(frm, ctl)
        Case "TextBox":     GoSub GetTag: GoSub BuildTextBox   'strWhere = strWhere & strAnd & BuildWhere_TextBox(frm, ctl)
        Case "OptionGroup": GoSub GetTag: GoSub BuildOptionGroup   'strWhere = strWhere & strAnd & BuildWhere_OptionGroup(frm, ctl)
        Case "CheckBox":    GoSub GetTag: GoSub BuildCheckBox      'strWhere = strWhere & strAnd & BuildWhere_CheckBox(frm, ctl)
        Case "ComboBox":    GoSub GetTag: GoSub BuildComboBox
        
    End Select

    If (Len(strWhere) > 0) Then strAnd = strOpt
    
    Return

'##########################################################################################
'#                                    Retrieve Tag Items                                  #
'##########################################################################################

GetTag:

    strFieldName = BuildWhere_GetTag("FieldName", ctl.Tag)
    If (Len(strFieldName) = 0) Then Err.Raise vbObjectError + 2000, "BuildWhere (GetTag);" & Err.Source, "Invalid Tag Property (" & ctl.Tag & ")" & vbCrLf & vbCrLf & "The Tag Property should look like this:" & vbCrLf & mstrcTagID & "TableName.FieldName" & mstrcFldSeparator & "FieldType[" & mstrcFldSeparator & "Operator]" & vbCrLf & vbCrLf & "(where FieldType is either String, Date, or Number and Operator (optional) is either =, <>, <, >, <=, >=, Like)"
    strFieldType = BuildWhere_GetTag("FieldType", ctl.Tag)
    
    strOperator = BuildWhere_GetTag("Operator", ctl.Tag)
    strFieldValue = BuildWhere_GetTag("Value", ctl.Tag)
    
    Return
    
'##########################################################################################
'#                                         Build Range                                    #
'##########################################################################################

BuildRange:

    '**********************************************************
    '*  Determine control that represents the End date range  *
    '**********************************************************
    
    On Error Resume Next
    
    Set ctlEndR = frm(Left$(ctl.Name, Len(ctl.Name) - Len(mstrcRange_Begin)) & mstrcRange_End)
    
    If (Err.Number = 2465) Then
        On Error GoTo 0
        Err.Raise vbObjectError + 2001, "BuildWhere (" & strCtlType & ");" & Err.Source, "Invalid Range" & vbCrLf & vbCrLf & "You have declared a text box that represents the Begin Range (" & ctl.Name & ") but you did not declare a control that represents the End Range (" & Left$(ctl.Name, Len(ctl.Name) - Len(mstrcRange_Begin)) & mstrcRange_End & ")"
    End If
    
    If (IsNull(ctlEndR)) Then
        On Error GoTo 0
        Err.Raise vbObjectError + 2002, "BuildWhere (" & strCtlType & ");" & Err.Source, "Invalid Range" & vbCrLf & vbCrLf & "Your entered a value for the Begin Range but failed to enter a value for the End Range."
    End If
    
    '*****************
    '*  Build Where  *
    '*****************
    
    On Error GoTo Errhandler
                
    If (Len(strFieldValue) = 0) Then strFieldValue = CStr(ctl.value)
    
    If (strFieldType = "#") Then   'IFT, then Date Field
        strWhere = strWhere & strAnd & " (" & strFieldName & " Between #" & strFieldValue & "# AND #" & ctlEndR.value & " 23:59:59#) "
    Else
        strWhere = strWhere & strAnd & " (" & strFieldName & " Between " & strFieldType & strFieldValue & strFieldType & " AND " & strFieldType & ctlEndR.value & strFieldType & ") "
    End If
    
    strAnd = strOpt
                
    Return

'##########################################################################################
'#                                         List Box                                       #
'##########################################################################################

BuildListBox:

'*******************************************
'*  Determine Operator (=, >, Like, etc.)  *
'*******************************************
    
    strAndOr = vbNullString
    If (Len(strOperator) > 0) Then
        If (strOperator = "<>") Then strAndOr = " AND " Else strAndOr = " OR "
    End If

    If (Len(strOperator) = 0) Or (strOperator = "=") Then
        strWhere = strWhere & strAnd & " (" & strFieldName & " In ("
        strSuffix = ", "
    Else
        strSuffix = ") "
    End If
                
        
    If (ctl.MultiSelect) Then
        For Each varItem In ctl.ItemsSelected
            If (Len(strOperator) = 0) Or (strOperator = "=") Then
                strWhere = strWhere & strFieldType & ctl.Column(ctl.BoundColumn - 1, varItem) & strFieldType & strSuffix
            Else
                strWhere = strWhere & strAnd & "(" & strFieldName & " " & strOperator & " " & strFieldType & ctl.Column(ctl.BoundColumn - 1, varItem) & strFieldType & strSuffix
                strAnd = " AND "
            End If
        Next varItem

        If (Len(strOperator) = 0) Or (strOperator = "=") Then
            strWhere = Mid(strWhere, 1, Len(strWhere) - Len(strSuffix)) & ")) "
        Else
            strWhere = Mid(strWhere, 1, Len(strWhere) - Len(strSuffix)) & ") "
        End If
    
    Else
        If (Len(strFieldValue) = 0) Then strFieldValue = ctl.Column(ctl.BoundColumn - 1)
        If (Len(strOperator) = 0) Or (strOperator = "=") Then
            strWhere = strWhere & strFieldType & strFieldValue & strFieldType & ")) "
        Else
            strWhere = strWhere & strAnd & "(" & strFieldName & " " & strOperator & " " & strFieldType & ctl.Column(ctl.BoundColumn - 1, varItem) & strFieldType & ") "
        End If
    End If
    
    strAnd = strOpt
    
    Return

'##########################################################################################
'#                                         Text Box                                       #
'##########################################################################################

BuildTextBox:

    If (Len(strFieldValue) = 0) Then strFieldValue = ctl.value
    
    strWhere = strWhere & strAnd & " (" & strFieldName & " " & strOperator & " " & strFieldType & strFieldValue & strFieldType & ") "
    
    strAnd = strOpt
                
    Return

'##########################################################################################
'#                                       Option Group                                     #
'##########################################################################################

BuildOptionGroup:
    If (Len(strFieldValue) = 0) Then strFieldValue = ctl.value
    strWhere = strWhere & strAnd & " (" & strFieldName & " " & strOperator & " " & strFieldType & strFieldValue & strFieldType & ") "
    strAnd = strOpt
    Return
'##########################################################################################
'#                                         Check Box                                      #
'##########################################################################################
BuildCheckBox:
    If (Len(strFieldValue) = 0) Then strFieldValue = ctl.value
    strWhere = strWhere & strAnd & " (" & strFieldName & " " & strOperator & " " & strFieldType & strFieldValue & strFieldType & ") "
    strAnd = strOpt
    Return
'##########################################################################################
'#                                         Combo Box                                      #
'##########################################################################################
BuildComboBox:
'*******************************************
'*  Determine Operator (=, >, Like, etc.)  *
'*******************************************
    strAndOr = vbNullString
    If (Len(strOperator) = 0) Or (strOperator = "=") Then
        strWhere = strWhere & strAnd & " (" & strFieldName & " = "
        strSuffix = ") "
    Else
        strWhere = strWhere & strAnd & " (" & strFieldName & " " & strOperator & " "
        strSuffix = ") "
    End If
    If (Len(strFieldValue) = 0) Then strFieldValue = ctl.Column(ctl.BoundColumn - 1)
    strWhere = strWhere & strFieldType & strFieldValue & strFieldType & strSuffix
    strAnd = strOpt
    Return
'****************************
'*  Error Recovery Section  *
'****************************
Errhandler:
    Err.Raise Err.Number, "BuildWhere (" & strCtlType & ");" & Err.Source, Err.Description
End Function

'+********************************************************************************************
'*
'$  Function:   BuildWhere_ControlType
'*
'*  Author:     FancyPrairie
'*
'*  Date:       April, 1998
'*
'*  Purpose:    This routine determines if a control should be included (valid) if it meets the following conditions:
'*
'-********************************************************************************************
'
Function BuildWhere_ControlType(frm As Form, ctl As Control) As String
'********************************
'*  Declaration Specifications  *
'********************************
    Dim strTemp As String       'Working variable
    Dim varValue As Variant     'Working variable
'****************
'*  Initialize  *
'****************
    On Error GoTo Errhandler

    BuildWhere_ControlType = vbNullString       'Assume invalid
    strTemp = Replace(ctl.Tag, " ", vbNullString)   'Strip out all spaces
    If (InStr(strTemp, mstrcTagID) > 0) Then        'If true, Tag Property contains "Where="
        If (ctl.ControlType = acListBox) Then
            If (ctl.MultiSelect) And (ctl.ItemsSelected.Count > 0) And (ctl.Enabled) And (ctl.Visible) Then BuildWhere_ControlType = "ListBox"
            If (Not ctl.MultiSelect) And (Not IsNull(ctl.value)) And (ctl.Enabled) And (ctl.Visible) Then BuildWhere_ControlType = "ListBox"
        ElseIf ((ctl.ControlType = acTextBox) Or (ctl.ControlType = acComboBox) Or (ctl.ControlType = acOptionGroup) Or (ctl.ControlType = acOptionButton) Or (ctl.ControlType = acCheckBox)) And (ctl.Enabled) And (ctl.Visible) Then
            If (ctl.ControlType = acOptionButton) Then
                If (ctl.Parent.ControlType = acOptionGroup) Then
                    If (ctl.OptionValue = ctl.Parent.value) Then varValue = ctl.Parent.value Else varValue = Null
                Else
                    varValue = ctl.value
                End If
            Else
                varValue = ctl.value
            End If
            If (Not IsNull(varValue)) And (Len(varValue) > 0) Then
                If (Right$(ctl.Name, Len(mstrcRange_End)) = mstrcRange_End) Then GoTo ExitProcedure
                If (Right$(ctl.Name, Len(mstrcRange_Begin)) = mstrcRange_Begin) Then
                    On Error GoTo 0
                    If (Not BuildWhere_ValidRange(frm, ctl)) Then Err.Raise vbObjectError + 2002, "BuildWhere_ControlType", "Incomplete Range Specifications." & vbCrLf & vbCrLf & "You failed to enter/select a value for " & ctl.Name
                    On Error GoTo Errhandler
                    BuildWhere_ControlType = "Range"
                ElseIf (ctl.ControlType = acTextBox) Then
                    BuildWhere_ControlType = "TextBox"
                ElseIf (ctl.ControlType = acComboBox) Then
                    BuildWhere_ControlType = "ComboBox"
                ElseIf (ctl.ControlType = acOptionGroup) Or (ctl.ControlType = acOptionButton) Then
                    BuildWhere_ControlType = "OptionGroup"
                ElseIf (ctl.ControlType = acCheckBox) Then
                    BuildWhere_ControlType = "CheckBox"
                End If
            End If
        End If
    End If
'********************
'*  Exit Procedure  *
'********************
ExitProcedure:
    Exit Function
'****************************
'*  Error Recovery Section  *
'****************************
Errhandler:
    Err.Raise Err.Number, "BuildWhere_ControlType;" & Err.Source, Err.Description
End Function


'+********************************************************************************************
'*
'$  Function:   BuildWhere_ValidRange
'*
'*  Author:     FancyPrairie
'*
'*  Date:       April, 1998
'*
'*  Purpose:    Determines if the values of 2 controls, that represent ranges, are valid.  They are not valid
'*              if one or both controls do not contain a '-********************************************************************************************
'
Function BuildWhere_ValidRange(frm As Form, ctl As Control) As Boolean

'********************************
'*  Declaration Specifications  *
'********************************
    Dim ctlEnd As Control
'******************************************************************
'*  Create array with just those text boxes that meet the specs.  *
'******************************************************************
    On Error GoTo Errhandler
    Set ctlEnd = frm(Left(ctl.Name, Len(ctl.Name) - Len(mstrcRange_Begin)) & mstrcRange_End)
    If (IsNull(ctl)) Or (Len(ctl) = 0) Or (IsNull(ctlEnd)) Or (Len(ctlEnd) = 0) Then
        BuildWhere_ValidRange = False
    Else
        BuildWhere_ValidRange = True
    End If
'********************
'*  Exit Procedure  *
'********************
ExitProcedure:
    Exit Function
'****************************
'*  Error Recovery Section  *
'****************************
Errhandler:
    Err.Raise Err.Number, "BuildWhere_ValidRange;" & Err.Source, Err.Description
End Function

'+********************************************************************************************
'*
'$  Function:   BuildWhere_GetTag
'*
'*  Author:     FancyPrairie
'*
'*  Date:       April, 1998
'*
'*  Purpose:    This routine will return the value for a given Tag Item.  The caller specifies
'*              which item they want returned.  Possible items are:
'*                  1. FieldName
'*                  2. FieldType
'*                  3. Operator
'*                  4. Value
'*
'-********************************************************************************************
'
Function BuildWhere_GetTag(strFunction As String, strTag As String) As String
'********************************
'*  Declaration Specifications  *
'********************************
    Dim i As Integer            'Working variable
    Dim j As Integer            'Working variable
    Dim k As Integer            'Working variable
    Dim strTemp As String       'Working variable
    Dim var As Variant          'Working variable
    On Error GoTo Errhandler
'**************************************************************
'*  Loop to find "Where=" Tag within Controls Tag Property    *
'*  When the "Where=" is found, then parse out the item       *
'*  the caller requested (i.e. FieldName,FieldType,Operator)  *
'**************************************************************
    var = Split(strTag, mstrcTagSeparator)
    BuildWhere_GetTag = vbNullString
    For i = 0 To UBound(var)
        strTemp = Replace(CStr(var(i)), " ", vbNullString)
        j = InStr(1, strTemp, mstrcTagID)   'Find "Where="
        If (j = 1) Then                     'If true, found "Where="
            strTemp = CStr(var(i))
            j = InStr(1, strTemp, "=")
            var = Split(Mid(strTemp, j + 1), mstrcFldSeparator)
            '***************
            '*  FieldName  *
            '***************
            If (strFunction = "FieldName") Then
                On Error Resume Next
                Eval (var(0))
                If (Err.Number) Then
                    Err.Clear
                    BuildWhere_GetTag = Trim(CStr(var(0)))
                Else
                    BuildWhere_GetTag = Eval(var(0))
                End If
                On Error GoTo Errhandler
            '***************
            '*  FieldType  *
            '***************
            ElseIf (strFunction = "FieldType") Then
                If (UBound(var) < 1) Then
                    BuildWhere_GetTag = vbNullString
                Else
                    Select Case Trim(CStr(var(1)))
                        Case "String": BuildWhere_GetTag = mstrcStringSymbol
                        Case "Date": BuildWhere_GetTag = mstrcDateSymbol
                        Case Else: BuildWhere_GetTag = vbNullString
                    End Select
                End If
            '**************
            '*  Operator  *
            '**************
            ElseIf (strFunction = "Operator") Then
                If (UBound(var) < 2) Then
                    BuildWhere_GetTag = "="
                Else
                    On Error Resume Next
                    Eval (var(2))
                    If (Err.Number) Then
                        Err.Clear
                        BuildWhere_GetTag = Trim(CStr(var(2)))
                    Else
                        BuildWhere_GetTag = Eval(var(2))
                    End If
                    On Error GoTo Errhandler
                End If
            '***********
            '*  Value  *
            '***********
            ElseIf (strFunction = "Value") Then
                If (UBound(var) < 3) Then
                    BuildWhere_GetTag = vbNullString
                Else
                    On Error Resume Next
                    Eval (var(3))
                    If (Err.Number) Then
                        Err.Clear
                        BuildWhere_GetTag = Trim(CStr(var(3)))
                    Else
                        BuildWhere_GetTag = Eval(var(3))
                    End If
                    On Error GoTo Errhandler
                End If
            End If
            Exit Function
        End If
    Next i
'********************
'*  Exit Procedure  *
'********************
ExitProcedure:
    Exit Function
'********************
'*  Error Recovery  *
'********************
Errhandler:
    Err.Raise Err.Number, "BuildWhere_GetTag;" & Err.Source, Err.Description
End Function


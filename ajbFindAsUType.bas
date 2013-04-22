Attribute VB_Name = "ajbFindAsUType"
'Author:        Allen Browne.   allen@allenbrowne.com
'Date:          August, 2006.
'Limitations:   Combos where the bound column is hidden have these limitations:
'                   -RowSourceType is "Table/Query": filtering supported in Access 2002 or later.
'                   -RowSourceType is "Value List" or a callback function: filtering not supported.
'Instructions:  1. Copy this module into your database.
'               2. Copy the combo and text box onto your form.
'               3. Set the form's On Load property to:  [Event Procedure]
'               4. Click the Build button beside the form's On Load property. Access opens the code window.
'               5. In the code window, between the "Private Sub Form_Load" and "End Sub" lines, enter this:
'                       Call FindAsUTypeLoad(Me)
'Documentation: http://allenbrowne.com/AppFindAsUType.html
Option Compare Database
Option Explicit

'Configuration options
Private Const mbcStartOfField = False   'True to match only the start of the field; False for anywhere in field.
Private Const mstrcWildcardChar = "*"   'Pattern matching wildcard. "*" for Access. "%" for SQL Server.
Private Const mstrcSep = ";"            'Separator between list items. May need changing for some regional settings.

'Columns of cboFindAsUTypeField
Private Const micControlName = 0
Private Const micControlLabel = 1
Private Const micControlType = 2
Private Const micFilterField = 3
Private Const micFieldType = 4

'Constant to indicate a control is sitting on the form (not on the page of a tab control.)
Private Const mlngcOnTheForm = -1&

'Module name (for error handler.)
Private Const conMod = "ajbFindAsUType"

Public Function FindAsUTypeLoad(frm As Form, ParamArray avarExceptionList()) As Boolean
On Error GoTo Err_Handler
    'Purpose:   Initialize the code for Find.
    'Return:    True on success.
    'Arguments: - frm = a reference to the form where you want this filtering.
    '           - Optionally, you can specify controls NOT to offer filtering on, by putting the control names in quotes.
    'Note:      The form must contain the 2 controls, cboFindAsUTypeField and txtFindAsUTypeValue,
    '               with the combo set up correctly.
    'Usage:     Set the Load event procedure of the form to:
    '               Call FindAsUType(Me)
    '           To suppress filtering on controls FirstName and City, use:
    '               Call FindAsUType(Me, "FirstName", "City")
    Dim rs As DAO.Recordset             'Clone set of the form.
    Dim ctl As Control                  'Each control on the form.
    Dim strForm As String               'Name of form (for error handler.)
    Dim strControl As String            'Name of the control.
    Dim strField As String              'Name of the filter to use in the filter string.
    Dim strControlSource As String      'Name of the field the control is bound to.
    Dim strOut As String                'List for the RowSource of cboFindAsUTypeField.
    Dim lngI As Long                    'Loop counter.
    Dim lngJ As Long                    'Page counter loop controller.
    Dim bSkip As Boolean                'Flag to provide no filtering for this control.
    Dim bResult As Boolean              'Return value for this function.
    Dim lngParentNumber As Long         '-1 if the control is directly on the form, else PageIndex of it parent.
    Dim lngMaxParentNumber As Long      'PageIndex of last page of tab control. -1 if no tab control.
    Dim astrControls() As String        'Array to handle the controls on the form.
    Const lngcControl = 0&              'First element of array astrControls is the control name.
    Const lngcField = 1&                'Second element of the array is the field name to filter on.
    
    'The form must have a control source if we are to filter it, and needs our 3 controls.
    strForm = frm.Name
    
    If HasUnboundControls(frm, "cboFindAsUTypeField", "txtFindAsUTypeValue") And (frm.RecordSource <> vbNullString) Then
        'Set the event handers for the 2 contorls
        frm!cboFindAsUTypeField.AfterUpdate = "=FindAsUTypeChange([Form])"
        frm.txtFindAsUTypeValue.OnChange = "=FindAsUTypeChange([Form])"
        'Calculate the number of pages on the tab control if there is one.
        lngMaxParentNumber = MaxParentNumber(frm)
        
        'Declare an array large enough to handle the controls on the form,
        '   for each page of any tab control (since these have their own tab index),
        '   and for storing the control name and the filter field name.
        ReDim astrControls(0& To frm.Controls.Count - 1&, mlngcOnTheForm To lngMaxParentNumber, lngcControl To lngcField) As String
        Set rs = frm.RecordsetClone             'For info about the fields the controls are bound to.
        
        'Loop through the controls on the form.
        For Each ctl In frm.Controls
            'Ignore hidden controls, and limit ourselves to text boxes and combos.
            If ctl.Visible Then
                If (ctl.ControlType = acTextBox) Or (ctl.ControlType = acComboBox) Then
                    bSkip = False
                    strField = vbNullString
                    strControl = ctl.Name
                    'Ignore if the control name is in the exception list.
                    For lngI = LBound(avarExceptionList) To UBound(avarExceptionList)
                        If avarExceptionList(lngI) = strControl Then
                            bSkip = True
                            Exit For
                        End If
                    Next
                    
                    If Not bSkip Then
                        'Ignore if unbound, or bound to an expression.
                        strControlSource = ctl.ControlSource
                        If (strControlSource = vbNullString) Or (strControlSource Like "=*") Then
                            bSkip = True
                        Else
                            'Ignore yes/no fields, binary (JET uses for unknown), and complex data types (> 100.)
                            Select Case rs(strControlSource).Type
                            Case dbBoolean, dbLongBinary, dbBinary, dbGUID, Is > 100
                                bSkip = True
                            End Select
                        End If
                    End If
                    
                    'Ignore if we cannot specify the field to filter on.
                    If Not bSkip Then
                        strField = GetFilterField(ctl)
                        If strField = vbNullString Then
                            bSkip = True
                        End If
                    End If
                    
                    'Add this control name to our array, in the order of the tab index.
                    If Not bSkip Then
                        lngParentNumber = ParentNumber(ctl)
                        astrControls(ctl.TabIndex, lngParentNumber, lngcControl) = strControl
                        astrControls(ctl.TabIndex, lngParentNumber, lngcField) = strField
                    End If
                End If
            End If
        Next
        
        'Loop through the array of controls, to build the string for the RowSource of cboFindAsUTypeField (5 columns.)
        For lngJ = LBound(astrControls, 2) To UBound(astrControls, 2)
            For lngI = LBound(astrControls) To UBound(astrControls)
                If astrControls(lngI, lngJ, lngcControl) <> vbNullString Then
                    Set ctl = frm.Controls(astrControls(lngI, lngJ, lngcControl))
                    strOut = strOut & """" & ctl.Name & """" & mstrcSep & _
                        """" & Caption4Control(frm, ctl) & """" & mstrcSep & _
                        ctl.ControlType & mstrcSep & _
                        """" & astrControls(lngI, lngJ, lngcField) & """" & mstrcSep & _
                        """" & rs(ctl.ControlSource).Type & """" & mstrcSep
                End If
            Next
        Next
        rs.Close
        
        'Remove the trailing separator, and assign to the RowSource of cboFindAsUTypeField.
        lngI = Len(strOut) - Len(mstrcSep)
        If lngI > 0 Then
            With frm.cboFindAsUTypeField
                .RowSource = Left(strOut, lngI)
                .value = .ItemData(0)           'Initialize to the first item in the list.
            End With
            bResult = True            'Return True: the list loaded successfully.
        End If
    End If
    
    'Show the filter controls. (Separate routine, since they could fail if the control does not exist.)
    Call ShowHideControl(frm, "cboFindAsUTypeField", bResult)
    Call ShowHideControl(frm, "txtFindAsUTypeValue", bResult)
    
    'Return value
    FindAsUTypeLoad = bResult

Exit_Handler:
    Set ctl = Nothing
    Set rs = Nothing
    Exit Function

Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".FindAsUTypeLoad", "Form " & strForm)
    Resume Exit_Handler
End Function

Public Function FindAsUTypeChange(frm As Form) As Boolean
On Error GoTo Err_Handler
    'Purpose:   Filter the form, by the control named in cboFindAsUTypeField and the value in txtFindAsUTypeValue.
    'Return:    True unless an error occurred.
    'Usage:     The code assigns this to the Change event of the text box, and the AfterUpdate event of the combo.
    Dim strText As String       'The text of the text box.
    Dim lngSelStart As Long     'Selection Starting point.
    Dim strField As String      'Name of the field to filter on.
    Dim bHasFocus As Boolean    'True if the text box has focus (since it can be called from the combo too.)
    Const strcTextBox = "txtFindAsUTypeValue"

    'If the text box has focus, remember the selection insert point and use its Text. Otherwise use its Value.
    bHasFocus = (frm.ActiveControl.Name = strcTextBox)
    If bHasFocus Then
        strText = frm!txtFindAsUTypeValue.Text
        lngSelStart = frm!txtFindAsUTypeValue.SelStart
    Else
        strText = Nz(frm!txtFindAsUTypeValue.value, vbNullString)
    End If
    
    'Save any uncommitted edits in the form. (This loses the insertion point, and converts Text to Value.)
    If frm.Dirty Then
        frm.Dirty = False
    End If
    
    'Read the filter field name from the combo.
    strField = Nz(frm.cboFindAsUTypeField.Column(micFilterField), vbNullString)
    
    'Unfilter if there is no text to find, or no control to filter. Otherwise, filter.
    If (strText = vbNullString) Or (strField = vbNullString) Then
        frm.FilterOn = False
    Else
        frm.Filter = strField & " Like """ & IIf(mbcStartOfField, vbNullString, mstrcWildcardChar) & _
            strText & mstrcWildcardChar & """"
        frm.FilterOn = True
    End If
    
    'If the control had focus, restore focus if necessary, and set the insertion point.
    If bHasFocus Then
        If frm.ActiveControl.Name <> strcTextBox Then
            frm(strcTextBox).SetFocus
        End If
        If strText <> vbNullString Then
            frm!txtFindAsUTypeValue = strText
            frm!txtFindAsUTypeValue.SelStart = lngSelStart
        End If
    End If
    
    'Return True if the routine completed without error.
    FindAsUTypeChange = True

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
    Case 2474
        Resume Next
    Case 2185   'Text box loses focus when no characters left.
        Resume Exit_Handler
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "txtFindAsUTypeValue_Change"
        Resume Exit_Handler
    End Select
End Function

Private Function Caption4Control(frm As Form, ctl As Control) As String
On Error GoTo Err_Handler
    'Purpose:
    Dim strCaption As String

    '1st choice: Assign the caption of the attached label.
    strCaption = ctl.Controls(0).Caption
    
    '2nd choice: Read the caption from the label over the column in a continuous form.
    If strCaption = vbNullString Then
        strCaption = CaptionFromHeader(frm, ctl)
    End If
    
    'Strip the trailing semicolon.
    If Right$(strCaption, 1&) = ":" Then
        strCaption = Left$(strCaption, Len(strCaption) - 1&)
    End If
    'Strip the ampersand hotkey.
    If InStr(strCaption, "&") > 0& Then
        strCaption = Replace(strCaption, "&&", Chr$(31))
        strCaption = Replace(strCaption, "&", vbNullString)
        strCaption = Replace(strCaption, Chr$(31), "&")
    End If
    
    '3rd choice: Use the control name.
    If strCaption = vbNullString Then
        strCaption = ctl.Name
    End If
    
    Caption4Control = strCaption
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
    Case 2467&
        Resume Next
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Caption4Control()"
        Resume Exit_Handler
    End Select
End Function

Private Function CaptionFromHeader(frm As Form, ctl As Control) As String
On Error GoTo Err_Handler
    'Purpose:   Look for a label in the column header, directly over the control, in continuous form view.
    'Return:    Caption of the label if found.
    Dim ctlHeader As Control    'controls in the header of the form.
    Const icRadius = 120        'one twelveth of an inch, in twips.
    
    'If we are in Form view, and it's a Continuous Form,
    '   and there is a label in the Form Header directly above the column, return its Caption.
    If (frm.CurrentView = 1) And (frm.DefaultView = 1) Then
        For Each ctlHeader In frm.Section(acHeader).Controls
            If ctlHeader.ControlType = acLabel Then
                If (ctlHeader.Left > ctl.Left - icRadius) And (ctlHeader.Left < ctl.Left + icRadius) Then
                    CaptionFromHeader = ctlHeader.Caption
                End If
            End If
        Next
    End If
    
Exit_Handler:
    Set ctlHeader = Nothing
    Exit Function

Err_Handler:
    If Err.Number <> 2462& Then     'No such Section.
        Call LogError(Err.Number, Err.Description, conMod & ".CaptionFromHeader")
    End If
    Resume Exit_Handler
End Function

Private Function HasUnboundControls(frm As Form, ParamArray avarControlNames()) As Boolean
On Error GoTo Err_Handler
    'Purpose:   Return true if all the controls named in the array are present on the form, and are unbound.
    Dim lngI As Long
    Dim bCancel As Boolean
    
    If UBound(avarControlNames) > 0& Then
        'Loop through the named controls on the form.
        For lngI = LBound(avarControlNames) To UBound(avarControlNames)
            If frm.Controls(avarControlNames(lngI)).ControlSource <> vbNullString Then
                bCancel = True
                Exit For
            End If
        Next
        'If we did not drop to the error handler, the form has the named controls.
        HasUnboundControls = Not bCancel
    End If
    
Exit_Handler:
    Exit Function

Err_Handler:
    Resume Exit_Handler
End Function

Private Function MaxParentNumber(frm As Form) As Long
On Error GoTo Err_Handler
    'Purpose:   Return the PageIndex of the tab page that the control is on.
    'Return:    -1 if setting directly on the form, else the PageIndex of the last page of the tab control.
    'Note:      PageIndex is zero based, so subtract 1 from the count of pages.
    Dim ctl As Control          'Each control on the form.
    Dim lngReturn As Long
    
    lngReturn = mlngcOnTheForm      'Initialize to no tab control.
    For Each ctl In frm.Controls
        If ctl.ControlType = acTabCtl Then
            lngReturn = ctl.Pages.Count - 1
            Exit For                    'A form can have only one tab control.
        End If
    Next
    
    MaxParentNumber = lngReturn
    
Exit_Handler:
    Exit Function

Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".MaxParentNumber")
    Resume Exit_Handler
End Function

Private Function ParentNumber(ctl As Control) As Integer
On Error Resume Next
    'Purpose:   Return the PageIndex of the tab page that the control is on.
    'Return:    -1 if setting directly on the form, else the page of the tab control.
    'Note:      This works for text boxes and combos, not for labels or controls in an option group.
    Dim iReturn As Integer
    
    iReturn = ctl.Parent.PageIndex
    If Err.Number <> 0& Then
        iReturn = mlngcOnTheForm
    End If
    ParentNumber = iReturn
End Function

Private Function ShowHideControl(frm As Form, strControlName As String, bShow As Boolean) As Boolean
On Error Resume Next
    'Purpose:   Show or hide a control on the form, without error message.
    'Return:    True if the contorl's Visible property was set successfully.
    'Arguments: frm            = a reference to the form where the control is expected.
    '           strControlName = the name of the control to show or hide.
    '           bShow          = True to make visible; False to make invisible.
    'Note:      This is a separate routine, since hiding a non-existant control will error.
    frm.Controls(strControlName).Visible = bShow
    ShowHideControl = (Err.Number = 0&)
End Function

Private Function GetFilterField(ctl As Control) As String
On Error GoTo Err_Handler
    'Purpose:   Determine the field name to use when filtering on this control.
    'Return:    The field name the control is bound to, except for combos.
    '               In Access 2002 and later, we return the syntax Access uses for filtering these controls.
    'Argument:  The control we are trying to filter.
    'Note:      We don't use the Recordset of the combo, because:
    '               a) it's not supported earlier than Access 2002, and
    '               b) it's often not loaded at this point.
    '               Instead, we OpenRecordset to get the source field name,
    '               which works even if the field is aliased in the RowSource.
    '               Opening for append only is quicker, as it loads no existing records.
    Dim rs As DAO.Recordset     'To get information about the combo's RowSource.
    Dim iColumn As Integer      'The first visible column of the combo (zero-based.)
    Dim strField As String      'Return value: the field name to use for the filter string.
    Dim bCancel As Boolean      'Flag to not filter on this control.
    
    If ctl.ControlType = acComboBox Then
        iColumn = FirstVisibleColumn(ctl)
        If iColumn = ctl.BoundColumn - 1 Then
            'The bound column is the first visible column: filter on the control source field.
            strField = "[" & ctl.ControlSource & "]"
        Else
            'In Access 2002 and later, we can use the lookup syntax Access uses, if the source is a Table/Query.
            If Int(Val(SysCmd(acSysCmdAccessVer))) >= 10 Then
                If ctl.RowSourceType = "Table/Query" Then
                    Set rs = DBEngine(0)(0).OpenRecordset(ctl.RowSource, dbOpenDynaset, dbAppendOnly)
                    With rs.Fields(iColumn)
                        strField = "[Lookup_" & ctl.Name & "].[" & .SourceField & "]"
                    End With
                    rs.Close
                Else
                    bCancel = True  'Hidden bound column not supported if RowSourceType is Value List or call-back function.
                End If
            Else
                bCancel = True      'Hidden bound column not supported for versions earlier than Access 2002.
            End If
        End If
    Else
        'Not a combo: filter on the control source field.
        strField = "[" & ctl.ControlSource & "]"
    End If
    
    If strField <> vbNullString Then
        GetFilterField = strField
    ElseIf Not bCancel Then
        GetFilterField = "[" & ctl.ControlSource & "]"
    End If
    Set rs = Nothing

Exit_Handler:
    Exit Function

Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".GetFilterField")
    Resume Exit_Handler
End Function

Private Function FirstVisibleColumn(cbo As ComboBox) As Integer
On Error GoTo Err_Handler
    'Purpose:   Return the column number of the first visible column in a combo.
    'Return:    Column number. ZERO-BASED!
    'Argument:  The combo to examine.
    'Note:      Also returns zero on error.
    Dim i As Integer            'Loop controller.
    Dim varArray As Variant     'Array of the combo's ColumnWidths values.
    Dim iResult As Integer      'Colum number to return.
    Dim bFound As Boolean       'Flag that we found a value to return.
    
    If cbo.ColumnWidths = vbNullString Then
        'If no column widths are specified, the first column is visible.
        iResult = 0
        bFound = True
    Else
        'Parse the ColumnWidths string into an array, and find the first non-zero value.
        varArray = Split(cbo.ColumnWidths, mstrcSep)
        For i = LBound(varArray) To UBound(varArray)
            If varArray(i) <> 0 Then
                iResult = i
                bFound = True
                Exit For
            End If
        Next
        'If the column widths ran out before all columns were checked, the next column is the first visible one.
        If Not bFound Then
            If i < cbo.ColumnCount Then
                iResult = i
                bFound = True
            End If
        End If
    End If
    
    FirstVisibleColumn = iResult

Exit_Handler:
    Exit Function

Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".FirstVisibleColumn")
    Resume Exit_Handler
End Function

'------------------------------------------------------------------------------------------------
'You may prefer to replace this with a true error logger. See http://allenbrowne.com/ser-23a.html
Private Function LogError(ByVal lngErrNumber As Long, ByVal strErrDescription As String, _
    strCallingProc As String, Optional vParameters, Optional bShowUser As Boolean = True) As Boolean
On Error GoTo Err_LogError
    ' Purpose: Generic error handler.
    ' Arguments: lngErrNumber - value of Err.Number
    ' strErrDescription - value of Err.Description
    ' strCallingProc - name of sub|function that generated the error.
    ' vParameters - optional string: List of parameters to record.
    ' bShowUser - optional boolean: If False, suppresses display.
    ' Author: Allen Browne, allen@allenbrowne.com

    Dim strMSG As String      ' String for display in MsgBox

    Select Case lngErrNumber
    Case 0
        Debug.Print strCallingProc & " called error 0."
    Case 2501                ' Cancelled
        'Do nothing.
    Case 3314, 2101, 2115    ' Can't save.
        If bShowUser Then
            strMSG = "Record cannot be saved at this time." & vbCrLf & _
                "Complete the entry, or press <Esc> to undo."
            MsgBox strMSG, vbExclamation, strCallingProc
        End If
    Case Else
        If bShowUser Then
            strMSG = "Error " & lngErrNumber & ": " & strErrDescription
            MsgBox strMSG, vbExclamation, strCallingProc
        End If
        LogError = True
    End Select

Exit_LogError:
    Exit Function

Err_LogError:
    strMSG = "An unexpected situation arose in your program." & vbCrLf & _
        "Please write down the following details:" & vbCrLf & vbCrLf & _
        "Calling Proc: " & strCallingProc & vbCrLf & _
        "Error Number " & lngErrNumber & vbCrLf & strErrDescription & vbCrLf & vbCrLf & _
        "Unable to record because Error " & Err.Number & vbCrLf & Err.Description
    MsgBox strMSG, vbCritical, "LogError()"
    Resume Exit_LogError
End Function

'------------------------------------------------------------------------------------------------
'Replace() and Split() are needed for Access 97 only. Leave them commented out for later versions.
'------------------------------------------------------------------------------------------------
'Private Function Replace(strExpr As String, strFind As String, strReplace As String, Optional lngStart As Long = 1) As String
'    Dim strOut As String
'    Dim lngLenExpr As Long
'    Dim lngLenFind As Long
'    Dim lng As Long
'
'    lngLenExpr = Len(strExpr)
'    lngLenFind = Len(strFind)
'
'    If (lngLenExpr > 0&) And (lngLenFind > 0&) And (lngLenExpr >= lngStart) Then
'        lng = lngStart
'        If lng > 1 Then
'            strOut = Left$(strExpr, lng - 1&)
'        End If
'        Do While lng <= lngLenExpr
'            If Mid(strExpr, lng, lngLenFind) = strFind Then
'                strOut = strOut & strReplace
'                lng = lng + lngLenFind
'            Else
'                strOut = strOut & Mid(strExpr, lng, 1&)
'                lng = lng + 1
'            End If
'        Loop
'        Replace = strOut
'    End If
'
'End Function
'
'Private Function Split(strIn As String, strDelim As String) As Variant
'    'Purpose:   Return a variant array from the items in the string, delimited the 2nd argument.
'    Dim varArray As Variant     'Variant array for output.
'    Dim lngStart As Long        'Position in string where argument starts.
'    Dim lngEnd As Long          'Position in string where argument ends.
'    Dim lngLenDelim As Long     'Length of the delimiter string.
'    Dim i As Integer            'index to the array.
'
'    lngLenDelim = Len(strDelim)
'    If (lngLenDelim = 0&) Or (Len(strIn) = 0&) Then
'        ReDim varArray(0)       'Initialize a zero-item array.
'    Else
'        ReDim varArray(9)       'Initialize a 10 item array.
'        i = -1                  'First item will be zero when we add 1.
'        lngStart = 1            'Start searching at first character of string.
'
'        'Search for the delimiter in the input string, until not found any more.
'        Do
'            i = i + 1
'            If i > UBound(varArray) Then    'Add more items if necessary
'                ReDim Preserve varArray(UBound(varArray) + 10)
'            End If
'
'            lngEnd = InStr(lngStart, strIn, strDelim)
'            If lngEnd = 0 Then  'This is the last item.
'                varArray(i) = Trim(Mid(strIn, lngStart))
'                Exit Do
'            Else
'                varArray(i) = Trim(Mid(strIn, lngStart, lngEnd - lngStart))
'                lngStart = lngEnd + lngLenDelim
'            End If
'        Loop
'        'Set the upper bound of the array to the correct number of items.
'        ReDim Preserve varArray(i)
'    End If
'
'    Split = varArray
'End Function
'------------------------------------------------------------------------------------------------


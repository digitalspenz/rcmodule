Attribute VB_Name = "modUtilities"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : FileExist
' DateTime  : 2007-Mar-06 13:51
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Test for the existance of a file; Returns True/False
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strFile - name of the file to be tested for including full path
'---------------------------------------------------------------------------------------
Function FileExist(strFile As String) As Boolean
On Error GoTo Err_Handler
 
    FileExist = False
    If Len(Dir(strFile)) > 0 Then
        FileExist = True
    End If
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: FileExist" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo Exit_Err_Handler
End Function

'***************** Code Start *******************
'This code was originally written by Dev Ashish
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish
'
'use to check if file or directory exists.
'
Function fIsFileDIR(stPath As String, _
                    Optional lngType As Long) _
                    As Integer
'Fully qualify stPath
'To check for a file
'   ?fIsFileDIR("c:\winnt\win.ini")
'To check for a Dir
'   ?fIsFileDir("c:\msoffice",vbdirectory)
'
    On Error Resume Next
    fIsFileDIR = Len(Dir(stPath, lngType)) > 0
End Function
'***************** Code End *********************


' ********** Code Start **************
'This code was originally written by Dev Ashish
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish
'
Public Function Round( _
ByVal Number As Variant, NumDigits As Long, _
Optional UseBankersRounding As Boolean = False) As Double
'
' ---------------------------------------------------
' From "Visual Basic Language Developer's Handbook"
' by Ken Getz and Mike Gilbert
' Copyright 2000; Sybex, Inc. All rights reserved.
' ---------------------------------------------------
'
  Dim dblPower As Double
  Dim varTemp As Variant
  Dim intSgn As Integer

  If Not IsNumeric(Number) Then
    ' Raise an error indicating that
    ' you've supplied an invalid parameter.
    Err.Raise 5
  End If
  dblPower = 10 ^ NumDigits
  ' Is this a negative number, or not?
  ' intSgn will contain -1, 0, or 1.
  intSgn = Sgn(Number)
  Number = Abs(Number)

  ' Do the major calculation.
  varTemp = CDec(Number) * dblPower + 0.5
  
  ' Now round to nearest even, if necessary.
  If UseBankersRounding Then
    If Int(varTemp) = varTemp Then
      ' You could also use:
      ' varTemp = varTemp + (varTemp Mod 2 = 1)
      ' instead of the next If ...Then statement,
      ' but I hate counting on TRue == -1 in code.
      If varTemp Mod 2 = 1 Then
        varTemp = varTemp - 1
      End If
    End If
  End If
  ' Finish the calculation.
  Round = intSgn * Int(varTemp) / dblPower
End Function


' ********** Code End **************
'---------------------------------------------------------------------------------------
' Procedure : AdjustPercent
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Automatically adjust whole number to percentage values
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2010-Sep-21                 Initial Release
'---------------------------------------------------------------------------------------
Function AdjustPercent(sValue As Variant) As Double
On Error GoTo Error_Handler
 
    If IsNumeric(sValue) = True Then            'Only treat numeric values
        If Right(sValue, 1) = "%" Then
            sValue = Left(sValue, Len(sValue) - 1)
            AdjustPercent = CDbl(sValue)
        End If
 
        If sValue > 1 Then
            sValue = sValue / 100
            AdjustPercent = sValue
        Else
            AdjustPercent = sValue
        End If
    Else                                        'Data passed is not of numeric type
        AdjustPercent = 0
    End If
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: AdjustPercent" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, _
           "An Error has Occured!"
    Resume Error_Handler_Exit
End Function


Public Function GetSpacesCount(fieldValue As String) As Long
If gbooHandleErrors Then On Error GoTo Errhandler
'This will count the number of spaces for every word
'Use on query to identify count of spaces like this: GetSpacesCount([YourField] & "") as CountSpaces
'then on where clause of the query put: GetSpacesCount([YourField] & "") > 0
    
    Dim i As Long, intLen As Long, intCount As Long
    Dim trimmedVal As String, isEnd As Boolean
    
    trimmedVal = Trim(fieldValue)
    intLen = Len(trimmedVal)
    
    For i = 1 To intLen
        If Mid(trimmedVal, i, 1) = " " Then
            intCount = intCount + 1
        End If
    Next
    
    intLen = Len(fieldValue)
    intCount = IIf(intCount > 1, intCount, 0)
    
    For i = 1 To intLen
        If Mid(fieldValue, i, 1) = " " Then
            intCount = intCount + 1
        Else
            If isEnd = False Then
                isEnd = True
                i = i + Len(trimmedVal) - 1
            End If
        End If
    Next
    GetSpacesCount = intCount


Exit_Now:
Exit Function

Errhandler:
DisplayUnexpectedError Err.Number, Err.Description
Resume Exit_Now
Resume
End Function


Public Function RemoveExtraSpaces(ByVal strText As String) As String
' Remove extra spaces
' http://www.utteraccess.com/wiki/index.php/Remove_extra_spaces
' Code courtesy of UtterAccess Wiki
' Licensed under Creative Commons License
' http://creativecommons.org/licenses/by-sa/3.0/
'
' You are free to use this code in any application,
' provided this notice is left unchanged.
'
' rev  date                          brief descripton
' 1.0  2011-07-26

If gbooHandleErrors Then On Error GoTo Errhandler

   Do Until InStr(1, strText, "  ") = 0
       strText = Replace(strText, "  ", " ", 1)
   Loop
   RemoveExtraSpaces = Trim(strText)
   
Exit_Now:
Exit Function

Errhandler:
DisplayUnexpectedError Err.Number, Err.Description
Resume Exit_Now
Resume
End Function

Public Function dev_TbrCloseFrmsRpts()
'Purpose  : Close all forms and reports opened in design view
'           except the current one.
'DateTime : 11/2/2007 07:46
'Author   : Bill Mosca
'Reference:
    Dim objObject As AccessObject
    Dim intPrompt As Integer
    Dim intUserResp As Integer


    'Set Save flag. Automatically save unless user wants prompt.
    'acSavePrompt=0; acSaveYes=1
    intUserResp = MsgBox("Objects will automatically be saved unless " _
            & "you specify a prompt." _
            & vbNewLine & vbNewLine _
            & "Prompt to save?", vbQuestion + vbYesNo + vbDefaultButton2, "Save?")

    Select Case intUserResp
        Case vbYes
            intPrompt = acSavePrompt
        Case vbNo
            intPrompt = acSaveYes
    End Select
    
    For Each objObject In CurrentProject.AllForms
        If objObject.IsLoaded = True _
                And Application.CurrentObjectName <> objObject.Name Then
            DoCmd.Close acForm, objObject.Name, intPrompt
        End If
    Next

    For Each objObject In CurrentProject.AllReports
        If objObject.IsLoaded = True _
                And Application.CurrentObjectName <> objObject.Name Then
            DoCmd.Close acReport, objObject.Name, intPrompt
        End If
    Next

End Function

'---------------------------------------------------------------------------------------
' Procedure : FrmRequery
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Requery the form to apply the choosen ordering,
'               but ensure we remain on the current record after the requery
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' frm       : The form to requery
' sPkField  : The primary key field of that form
'
' Usage:
' ~~~~~~
' Call FrmRequery(Me, "Id")
' Call FrmRequery(Me, "ContactId")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-Oct-19                 Initial Release
'---------------------------------------------------------------------------------------
Sub FrmRequery(frm As Form, sPkField As String)
    On Error GoTo Error_Handler
    Dim rs              As DAO.Recordset
    Dim pk              As Long
 
    pk = frm(sPkField)
    frm.Requery
    Set rs = frm.RecordsetClone
    rs.FindFirst "[" & sPkField & "]=" & pk
    frm.Bookmark = rs.Bookmark
 
Error_Handler_Exit:
    On Error Resume Next
    Set rs = Nothing
    Exit Sub
 
Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: FrmRequery" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Sub


'===================================================== FORMS
Function CancelMe() As Byte
   'example useage: OnClick event of an Undo command button
   ' = CancelMe()
   'if there is nothing to Undo, this will create an error -- just ignore
  On Error Resume Next
  DoCmd.RunCommand acCmdUndo
End Function

Public Function HideNavPane() As Byte

   On Error Resume Next
   '-Not being used as of the moment
      DoCmd.SelectObject acTable, , True
   '    DoCmd.SelectObject acTable, "MSysObjects", True
    DoCmd.RunCommand acCmdWindowHide
    
End Function

Function ClearList(lst As ListBox) As Boolean
        
    'Purpose:   Unselect all items in the listbox.
    'Return:    True if successful
    'Author:    Allen Browne. http://allenbrowne.com  June, 2006.
    On Error Resume Next
    Dim varItem As Variant
    If lst.MultiSelect = 0 Then
        lst = Null
    Else
        For Each varItem In lst.ItemsSelected
            lst.Selected(varItem) = False
        Next
    End If
    ClearList = True

End Function

Public Function SelectAll(lst As ListBox) As Boolean
    'Purpose:   Select all items in the multi-select list box.
    'Return:    True if successful
    'Author:    Allen Browne. http://allenbrowne.com  June, 2006.
    Dim lngRow As Long

    If lst.MultiSelect Then
        For lngRow = 0 To lst.ListCount - 1
            lst.Selected(lngRow) = True
        Next
        SelectAll = True
    End If
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ RequeryMe
Function RequeryMe(Optional pC As Control) As Byte
   'used to rebuild combo box and listbox lists
   'put on the double-click event of a combobox
   ' =RequeryMe()
   ' =RequeryMe([listRels])
   ' if control is not specief. ActiveControl will be used
   On Error GoTo Proc_Err
   If pC Is Nothing Then
      Screen.ActiveControl.Requery
   Else
      pC.Requery
   End If
   Exit Function
Proc_Err:
   MsgBox Err.Number & " " & Err.Description, , "Cannot Requery control right now"
End Function

Function DisableControls(frm As Form)
On Error GoTo Errhandler

   Dim ctl As Control
   For Each ctl In frm.Controls
      Select Case ctl.ControlType
         Case acTextBox, acComboBox, acCommandButton, acOptionGroup, acSubform
            If ctl.Tag <> "*" Then ctl.Enabled = False
      End Select
   Next

Exit_Now:
Exit Function

Errhandler:
DisplayUnexpectedError Err.Number, Err.Description
Resume Exit_Now
Resume
End Function

Function EnableControls(frm As Form)
On Error GoTo Errhandler
   
   Dim ctl As Control
   For Each ctl In frm.Controls
      Select Case ctl.ControlType
         Case acTextBox, acComboBox, acCommandButton, acOptionGroup, acSubform
            If ctl.Tag <> "*" Then ctl.Enabled = True
      End Select
   Next

Exit_Now:
Exit Function

Errhandler:
DisplayUnexpectedError Err.Number, Err.Description
Resume Exit_Now
Resume
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~IsBlank
Public Function IsBlank(arg As Variant) As Boolean
'-----------------------------------------------------------------------------
' True if the argument is Nothing, Null, Empty, Missing or an empty string .
'-----------------------------------------------------------------------------
' Here assume that CustomerReference is a control on a form.
' By using IsBlank() we avoid having to test both for Null and empty string.

    'If IsBlank(CustomerReference) Then
        'MsgBox "Customer Reference cannot be left blank."
    'End If
    Select Case VarType(arg)
        Case vbEmpty
            IsBlank = True
        Case vbNull
            IsBlank = True
        Case vbString
            IsBlank = (arg = vbNullString)
        Case vbObject
            IsBlank = (arg Is Nothing)
        Case Else
            IsBlank = IsMissing(arg)
    End Select
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ IsLoaded
Function IsLoaded(pFormname As String) As Boolean
   'This function returns  TRUE if the passed form is loaded  FALSE if it is not
   'example useage: call before opening a form
   ' If IsLoaded("Formname") Then DoCmd.SelectObject acForm, "Formname"
   IsLoaded = False
   '  True if the specified form is open not in Design view
   If CurrentProject.AllForms(pFormname).IsLoaded Then
      If Forms(pFormname).CurrentView <> 0 Then IsLoaded = True
   End If
   Exit Function
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ CorrectProper
Function CorrectProper(pString As String) As String
'5-15-07
   'correct:
   'Macx* --> MacX*
   'Vanx* --> VanX*
   'Mcx* --> McX*
   'will only correct if at beginning of passed string, not in the middle
   
   Select Case True
   Case Len(pString) > 4 And InStr(".Mac.Van.", "." & Left(pString, 3) & ".") > 0
      pString = Left(pString, 3) & UCase(Mid(pString, 4, 1)) & Right(pString, Len(pString) - 4)
   Case Left(pString, 2) = "Mc" And Len(pString) > 3
      pString = Left(pString, 2) & UCase(Mid(pString, 3, 1)) & Right(pString, Len(pString) - 3)
      
   End Select
   CorrectProper = pString
             
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ ProperCase
Function ProperCase(Optional pC As Control _
   , Optional pBoo As Boolean = True) As Byte
   
   '6-3-07
   'change active control to Proper case if pBoo = true
   
   'EXAMPLES
   
   ' on AfterUpdate property -->
   ' =ProperCase()
   ' =ProperCase([ControlName])
   ' =ProperCase(ActiveControl, chkConvertCase)
   '   where chkConvertCase is the name of a checkbox control specifying whether or not to change to ProperCase
 
   ' in code -->
   ' ProperCase Me.ActiveControl
   ' ProperCase Me.ActiveControl, me.chkConvertCase
   'OR simply (if you always want to do this regardless) -->
   'ProperCase

   'if there is an error ignore it (and it probably won't happen anyway
   '-- unless, for instance, you try this with a number)
   'since null values are tested <smile>
   
   'for a more in-depth version that corrects for things like MacDonald, O'Hare, etc, look here:
   '   Uppercase converter (incl. Auto-Correct) by Rob richards (r_Cubed)
   '   http://www.utteraccess.com/forums/showflat.php?Cat=&Board=48&Number=619856
   
   On Error Resume Next
   If Not pBoo Then Exit Function
   'if a control reference was not passed, use the active control on the screen
   If pC Is Nothing Then Set pC = Screen.ActiveControl
   'if the control is not filled out, don't do anything
   If IsNull(pC) Then Exit Function
   'convert the contents of the control to ProperCase
   pC = CorrectProper(StrConv(pC, vbProperCase))
        
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ UpperCase
Function UpperCase(Optional pC As Control _
   , Optional pBoo As Boolean = True) As Byte
   
   '6-2-07
   'change active control to Proper case if pBoo = true
   'EXAMPLE
   ' on AfterUpdate property -->
   ' =UpperCase([ControlName])
   '=UpperCase(ActiveControl, chkConvertCase)
   '   where chkConvertCase is the name of a checkbox control specifying whether or not to change to UpperCase
   ' in code --> UpperCase Me.ActiveControl

   'if there is an error ignore it (and it probably won't happen anyway
   '-- unless, for instance, you try this with a number)
   'since null values are tested <smile>
   
   'for a more in-depth version that corrects for things like MacDonald, O'Hare, etc, look here:
   '   Uppercase converter (incl. Auto-Correct) by Rob richards (r_Cubed)
   '   http://www.utteraccess.com/forums/showflat.php?Cat=&Board=48&Number=619856
   
   On Error Resume Next
   'if a control reference was not passed, use the active control on the screen
   If pC Is Nothing Then Set pC = Screen.ActiveControl
   'if the control is not filled out, don't do anything
   If IsNull(pC) Then Exit Function
   'convert the contents of the control to UpperCase
   pC = UCase(pC)
      
End Function

'===================================================== CONTROLS
'~~~~~~~~~~~~~~~~~~~~~~~~~~ DropMe
Function dropMe() As Byte
   On Error GoTo Proc_Err
   'usually used on the MouseUp event of a Combo Box
   'so you can click anywhere and drop the list
   'instead of just on the arrow
   '=DropMe()
   Screen.ActiveControl.Dropdown
   Exit Function
Proc_Err:
   MsgBox Err.Number & " " & Err.Description, , "Cannot drop list right now"
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ DropMeIfNull
Function DropMeIfNull() As Byte
   'usually used on the GotFocus event of a Combo Box
   'so if there is nothing filled out yet, the list will drop
   'Do NOT use on the first control in the tab order
   '=DropMeIfNull()
   On Error GoTo Proc_Err
   If IsNull(Screen.ActiveControl) Then Screen.ActiveControl.Dropdown
   Exit Function
Proc_Err:
   MsgBox Err.Description, , "Cannot drop list right now"
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ ZoomMe
Function ZoomMe() As Byte
   'pop up the ZOOM box for editing
   'used in text boxes where the text
   'may be longer than the display
   'put on the Double-Click event of a textbox
   '=ZoomMe()
   On Error Resume Next
   DoCmd.RunCommand acCmdZoomBox
   'this is the old way
   'SendKeys "+{F2}", True
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ Sort123
Function Sort123( _
   pF As Form _
   , pField1 As String _
   , Optional pField2 = "" _
   , Optional pField3 = "" _
   ) As Byte
'091203

   'written by Crystal
   'strive4peace2009 at yahoo dot com

   'sort form by specified field(s)
   'sending the same sort fields
   'toggles Ascending and Descending order

   ' --------------------------------------------------------
   ' PARAMETERS
   '  pF = form reference
   '       if in code behind a form, this is
   '                   Me
   '  pField1 -- name of field for first sort
   '  pField2 -- optional, name of field for second sort
   '  pField3 -- optional, name of field for third sort
   '
   ' --------------------------------------------------------
   ' NOTES
   '  you must specify FIELD names in the RecordSource
   '  control names do not matter
   ' --------------------------------------------------------
   '
   'USEAGE
   ' commonly called on
   '     CLICK event of column header label control
   '
   ' in code behind form to specify main and secondary sort fields
   '    Sort123 Me, "Fieldname1", "Fieldname2"
   
   'set up Error Handler
   On Error GoTo Proc_Err
   
   'dimension sort string variables
   ' for both ascending and descending cases
   Dim mOrderBy As String _
      , mOrderByZA As String
   
   'initialize the OrderBy string for ascending order
   mOrderBy = ""
   
   '  assign the first field to the OrderBy string
   If Len(Trim(pField1)) > 0 Then
      mOrderBy = pField1
      mOrderByZA = pField1 & " desc"
   
      '  assign the second field to the OrderBy string
      '  if it is specified
      If Len(Trim(pField2)) > 0 Then
         mOrderBy = (mOrderBy + ", ") & pField2
         mOrderByZA = (mOrderByZA + ", ") & pField2
      End If
      
      '  assign the third field to the OrderBy string
      '  if it is specified
      If Len(Trim(pField3)) > 0 Then
         mOrderBy = (mOrderBy + ", ") & pField3
         mOrderByZA = (mOrderByZA + ", ") & pField3
      End If
      
   Else
      ' no sort string specified
      ' remove OrderBy from the form
      pF.OrderByOn = False
      ' exit the procedure
      GoTo Proc_Exit
   End If
   
   ' use WITH to minimize the number of times
   ' this code will access the form
   
   With pF
      ' if the form is already sorted
      ' by the ascending sort string,
      ' then change order to be descending
      If .OrderBy = mOrderBy Then
         .OrderBy = mOrderByZA
      Else
         ' change the sort order to ascending
         ' if form is not sorted this way
         If .OrderBy <> mOrderBy Then
            .OrderBy = mOrderBy
         End If
      End If
      ' make the form use the specified sort order
      .OrderByOn = True
   End With
   
Proc_Exit:
   Exit Function
  
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   Sort123"
 
   Resume Proc_Exit

   'if you want to single-step code to find error, CTRL-Break at MsgBox
   'then set this to be the next statement
   Resume
   
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ BoldMe
Function BoldMe(Optional pF As Form _
   , Optional pControlname As String = "" _
   , Optional pNumOptions As Integer = 0 _
   , Optional pValue As Variant _
   ) As Byte
'9-9-08,12-4-09

   'Crystal
   'strive4peace2009 at yahoo dot com

   'Bold the label is the option is chosen or value is true
   'remove Bold is the value is not true or the option is not chosen
   
   ' --------------------------------------------------------
   'PARAMETERS
   '  pF = form reference
   '       if in code behind a form, this is
   '                   Me
   '
   '  pControlName is name of control to test
   '               if not specified, ActiveControl will be used
   '
   '  pNumOptions is the number of options in the frame (group)
   '              must be specified for option frame
   '
   '  pValue is the comparison value for deciding Bold
   '         if parameter is passed
   '         then the opn frame will not be tested
   '
   '               Labels MUST be named like this:
   '               Label_Controlname
   '
   '               NOTE: the "label" control
   '                         does not have to be a label ControlType
   '                     It can be, for instance, a textbox
   ' --------------------------------------------------------
   ' NOTES
   '
   ' for checkboxes and toggle buttons
   '
   '               if checkbox Name = MyCheckbox
   '                  then label Name = Label_MyCheckbox
   '
   ' for options in a frame
   '
   '    if Frame Name = MyOptionFrame
   '
   '    then Frame Option Buttons are Named:
   '          MyOptionFrame1, MyOptionFrame2, etc
   '
   '    Labels for Frame Option Buttons are Named:
   '          Label_MyOptionFrame1, Label_MyOptionFrame2, etc
   '
   '    Numbers in the name correspond to the Option Value
   '
   '    Option Values can be any number
   '

   ' --------------------------------------------------------
   'USEAGE
   '   BoldMe Me
   '       Bold the label of the
   '         active checkbox or toggle control
   '       if the control value = True
   '
   '   BoldMe Me, "Mycheckbox_controlname"
   '       Bold the label of the
   '         specified checkbox or toggle control
   '       if the control value = True
   '
   '   BoldMe Me, "Mycheckbox_controlname",,True
   '       Bold the label of the
   '         specified checkbox or toggle control
   '
   '   BoldMe Me, "MyFrame_controlname", 4
   '       Bold the label of the option
   '            in the specified frame control
   '            if the Option Value = the Frame Value
   '       where there are 4 options to pick from
   '
   '   BoldMe Me, "MyFrame_controlname", 4, 999
   '       Bold the label of the option
   '            in the specified frame control
   '            if the Option Value = 999
   '       where there are 4 options to pick from
   '

   
   On Error GoTo Proc_Err
   
   If pF Is Nothing Then Set pF = Screen.ActiveForm
   
   Dim mBoo As Boolean _
      , mControlName As String _
      , mControlnameOption As String
   
   If Len(pControlname) > 0 Then
      mControlName = pControlname
   Else
      mControlName = pF.ActiveControl.Name
   End If
   
   If IsMissing(pValue) Then
      pValue = pF(mControlName).value
   End If
   
   ' use WITH to minimize the number of times
   ' this code has to access the object
   
   'checkbox or toggle button
   With pF(mControlName)
   
      Select Case .ControlType
      Case acCheckBox, acToggleButton
         
         If IsMissing(pValue) Then
            mBoo = Nz(.value, False)
         Else
            'note: Null cannot be passed
            mBoo = pValue
         End If
         
         With pF("Label_" & mControlName)
            ' see if Bold is already right
            If .FontBold <> mBoo Then
               ' Bold needs to change
               .FontBold = mBoo
            End If
         End With
         
         GoTo Proc_Exit
   
      'option box - MUST SPECIFY pNumOptions
      Case acOptionGroup
      
         Dim i As Integer
                  
         For i = 1 To pNumOptions
            mControlnameOption = mControlName & Format(i, "0")
            If IsNull(pValue) Then
               ' if the comparison is blank
               ' no option will be bolded
               mBoo = False
            Else
               ' if the option value = the comparison value
               ' then mBoo = TRUE
               mBoo = IIf( _
               pF(mControlnameOption).OptionValue = pValue, True, False)
            End If
            
            With pF("Label_" & mControlnameOption)
               If .FontBold <> mBoo Then
                  .FontBold = mBoo
               End If
            End With
            
         Next i
         
         GoTo Proc_Exit
         
      End Select
      
   End With
   
Proc_Exit:
   On Error Resume Next
   pF.Repaint
   Exit Function
   
Proc_Err:
   MsgBox Err.Description _
      , , "ERROR " & Err.Number & "  BoldMe " & mControlName
   Resume Proc_Exit
   'if you want to single-step code to find error, CTRL-Break at MsgBox
   'then set this to be the next statement
   Resume
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsRptOpen
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Determine whether a report is open or not
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sRptName  : Name of the report to check if it is open or not
'
' Usage Example:
' ~~~~~~~~~~~~~~~~
' IsRptOpen("Report1")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2010-May-26                 Initial Release
'---------------------------------------------------------------------------------------
Function IsRptOpen(sRptName As String) As Boolean
On Error GoTo Error_Handler
 
    If Application.CurrentProject.AllReports(sRptName).IsLoaded = True Then
        IsRptOpen = True
    Else
        IsRptOpen = False
    End If
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.Number & vbCrLf & "Error Source: IsRptOpen" & vbCrLf & "Error Description: " & _
    Err.Description, vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ CloseMe
Function CloseMe(pF As Form _
   , Optional booSave As Boolean = False _
   , Optional pFormOpen As String = "" _
   ) As Byte
' Crystal (strive4peace)
' 8-17-08, 12-1 pOpenForm, 12-19 dirty, 4-9-09

   'close a form
   'example useage: [Event Procedure] of a Close command button
   '   CloseMe
   '   CloseMe Me
   ' close form and save changes
   '   CloseMe Me, true
   ' close form and open/switch to another
   '   CloseMe Me,, "OtherFormname"
   
   On Error GoTo Proc_Err
   
   If Len(pF.RecordSource) > 0 Then If pF.Dirty Then _
      pF.Dirty = False
   
   DoCmd.Close acForm, pF.Name _
      , IIf(booSave, acSaveYes, acSaveNo)
   
   If pFormOpen <> "" Then
      If CurrentProject.AllForms(pFormOpen).IsLoaded Then
        Forms(pFormOpen).Visible = True
        DoCmd.SelectObject acForm, pFormOpen
      Else
         DoCmd.OpenForm pFormOpen
      End If
   End If
   
   Exit Function
   
Proc_Err:
   MsgBox Err.Number & " " & Err.Description _
     , , "Cannot close right "
End Function
'~~~~~~~~~~~~~~~~~~~~~~~


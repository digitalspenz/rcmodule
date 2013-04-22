Attribute VB_Name = "mod_BackupRestore"
Option Compare Database
Option Explicit

 'Module Name: mod_BackupRestore
 'Crystal
 '4-26-07

Public Sub runBackupRestoreQFRMM()
   'click in this code and press F5 to run
   'make sure you backup up the database first!
   If MsgBox("If you have backed up your database and are ready, click Yes" _
      , vbYesNo + vbDefaultButton2 _
      , "Backup and Restore Queries, Forms, Reports, Macros, and Modules?") = vbNo Then Exit Sub
      
   BackupRestoreQFRMM True
      
End Sub

Public Function BackupRestoreQFRMM(Optional pBooRestore As Boolean = False)
 
   'must close all forms/reports before executing this
   'backup to text files: Queries, Forms, Reports, Macros, and Modules
   'optionally, delete these objects from the databse and restore them from text files
   '(mod_BackupRestore is NOT restored as that is where this code is)
   'DELETES temporary queries -- those starting with "~"
   'after running this, compile the databse
   
   'does NOT backup and restore tables
   
   'this module must be called --> mod_BackupRestore or it will fail due to "reserve error"
   
   'NEEDS REFERENCE TO
   'Microsoft DAO Object Library
   
   'written by Crystal strive4peace2007 at yahoo.com
   'based on code by:
   '-- Arvin Meyer   http://www.datastrat.com/Code/DocDatabase.txt
   '-- Brent Spauling (datAdrenaline)
   
   'PARAMETERS
   'pBooRestore - True = restore also, False = backup only
   '(if not specified, will backup only)

   'text files will be saved in a directory below the database called BAK_TEXT
   
    Dim db As DAO.Database _
      , Cnt As DAO.Container _
      , Doc As DAO.Document
      
   Dim mPath As String _
      , mFilename As String _
      , mFilenameCorrected As String _
      , mFileExt As String _
      , i As Integer _
      , mContainerName As String _
      , mObjType As Long _
      , mMsg As String _
      , mDocName As String
   
   mPath = CurrentProject.path & "\BAK_TEXT\"
   If Len(Dir(mPath, vbDirectory)) = 0 Then
      MkDir mPath
      DoEvents
   End If
   
   Set db = CurrentDb
   
   On Error GoTo Proc_Err
      
   With db
       'close all open forms, reports, queries
      For i = Forms.Count To 1 Step -1
         DoCmd.Close acForm, Forms(i)
      Next i
      For i = Reports.Count To 1 Step -1
         DoCmd.Close acReport, Reports(i)
      Next i
      On Error Resume Next
      For i = .QueryDefs.Count - 1 To 0 Step -1
         DoCmd.Close acQuery, .QueryDefs(i - 1)
      Next i
      On Error GoTo Proc_Err
       '-------------------------
      
      For i = 1 To 4
         Select Case i
          '---------- FORMS
         Case 1
            mContainerName = "Forms"
            mObjType = acForm
            mFileExt = ".frm"
            
          '---------- REPORTS
         Case 2
            mContainerName = "Reports"
            mObjType = acReport
            mFileExt = ".rpt"
         
          '---------- MACROS
         Case 3
            mContainerName = "Scripts"
            mObjType = acMacro
            mFileExt = ".mac"
            
          '---------- MODULES
         Case 4
            mContainerName = "Modules"
            mObjType = acModule
            mFileExt = ".mod"
            
         End Select
         
            
         Set Cnt = db.Containers(mContainerName)
   
         Debug.Print "---" & mContainerName & "---"
         For Each Doc In Cnt.Documents
            mDocName = Doc.Name
            
            mFilename = mPath & Doc.Name & mFileExt
            mFilenameCorrected = mPath & "Correct_" & Local_CorrectFilename(Doc.Name) & mFileExt
            
            mMsg = "backed up "
             'delete file if it already exists
            If Len(Dir(mFilenameCorrected)) > 0 Then
               Kill mFilenameCorrected
               DoEvents
            End If
            
            'Backup object
            Application.SaveAsText mObjType, mDocName, mFilenameCorrected
            
             'copy filename to actual name of object if different than corrected name
            If mFilenameCorrected <> mFilename Then
               FileCopy mFilenameCorrected, mFilename
            End If
            
             'delete object and restore from text file
            If Doc.Name <> "mod_BackupRestore" Then
               If pBooRestore Then
                  DoCmd.DeleteObject mObjType, mDocName
                  On Error Resume Next
                  DoEvents
                  On Error GoTo Proc_Err
                  LoadFromText mObjType, mDocName, mFilenameCorrected
                  mMsg = mMsg & "and restored "
               End If
            End If
            
            mMsg = mMsg & "--> " & mDocName
            Debug.Print mMsg
            
         Next Doc

      Next i
    
   End With
   
    '---------- QUERIES
   Debug.Print "--- Queries ---"
   For i = db.QueryDefs.Count - 1 To 0 Step -1
      mMsg = "backed up "
      mDocName = db.QueryDefs(i).Name
      If Left(mDocName, 1) <> "~" Then
         mFilename = mPath & mDocName & ".qry"
         mFilenameCorrected = mPath _
             & "Correct_" & Local_CorrectFilename(mDocName) & ".qry"
         
          'delete file if it already exists
         If Len(Dir(mFilenameCorrected)) > 0 Then
            Kill mFilenameCorrected
            DoEvents
         End If
         
          'Backup object
         If mFilename <> mFilenameCorrected Then
            If Len(Dir(mFilename)) > 0 Then
               Kill mFilename
               DoEvents
            End If
         End If
         Application.SaveAsText acQuery, mDocName, mFilenameCorrected
         If mFilenameCorrected <> mFilename Then
            FileCopy mFilenameCorrected, mFilename
         End If
         
          'restore object
         If pBooRestore Then
            mMsg = mMsg & "and restored "
            DoCmd.DeleteObject acQuery, mDocName
            On Error Resume Next
            DoEvents
            On Error GoTo Proc_Err
            LoadFromText acQuery, mDocName, mFilenameCorrected
         End If
         mMsg = mMsg & "--> " & mDocName
         Debug.Print mMsg
      End If
   Next i
       
   If pBooRestore Then
      mMsg = "backed up and restored objects" _
         & vbCrLf & vbCrLf _
         & "Compile the database then Compact/Repair"
   Else
      mMsg = "backed up objects"
   End If
   
   mMsg = mMsg & vbCrLf & vbCrLf & " deleted " _
      & DeleteTemporaryQueries() & " temporary queries"

   MsgBox mMsg, , "Done"
   
Proc_Exit:
   On Error Resume Next
    'close and release object variables
   Set Doc = Nothing
   Set Cnt = Nothing
   Set db = Nothing
   
   Exit Function
   
  
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   BackupRestoreQFRMM"
 
    'press F8 to step through code and debug
    'remove next line after debugged
   Stop:    Resume
 
   Resume Proc_Exit

End Function

Private Function Local_CorrectFilename(pName) As String
    'this is a copy of a function that is usually in my main general library
   'included here for convenience
   
   'written by Crystal strive4peace2007 at yahoo.com
   
   Dim i As Integer _
      , mChar As String * 1 _
      , mName As String _
      , mLastChar As String * 1 _
      , mNewChar As String
      
   If IsNull(pName) Then Exit Function
   
   pName = LTrim(Trim(pName))
   
   For i = 1 To Len(pName)
      mChar = Mid(pName, i, 1)
 '      use this line if you also want to replaces spaces
      If InStr("`!@#$%^&*()+=|\:;""'<>,.?/ ", mChar) > 0 Then
 '      If InStr("`!@#$%^&*()+=|\:;""'<>,.?/", mChar) > 0 Then
         mNewChar = "_"
      Else
         mNewChar = mChar
      End If
      
      If mLastChar = "_" And mNewChar = "_" Then
      Else
         mName = mName & mNewChar
      End If
      
      mLastChar = mNewChar
   Next i
   Local_CorrectFilename = mName
   
End Function

Function DeleteTemporaryQueries() As Integer
   Dim qdf As DAO.QueryDef _
      , i As Integer
   i = 0
   Debug.Print "--- deleted temporary queries ---"
   For Each qdf In CurrentDb.QueryDefs
      If Left(qdf.Name, 1) = "~" Then
         Debug.Print qdf.Name
         DoCmd.DeleteObject acQuery, qdf.Name
         i = i + 1
      End If
   Next qdf
   CurrentDb.QueryDefs.Refresh
   Debug.Print "--> " & i & " temporary queries deleted"
   DeleteTemporaryQueries = i
   
   Set qdf = Nothing
End Function
  
Sub RestoreOne()
   'change acForm to acReport, acQuery, acMacro, acModule as it applies
   'change Formname to the name of your form or other object to restore
   LoadFromText acForm, "Formname", "C:\path\filename.frm"
   
End Sub



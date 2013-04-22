Attribute VB_Name = "modOpenCloseDb"
Option Compare Database
Option Explicit

Public Sub AccessExample(strPath As Variant)
If gbooHandleErrors Then On Error GoTo Errhandler

    Dim objAccess As Access.Application

    'Create an instance of the Access application object.
    Set objAccess = CreateObject("Access.Application")
    
'    strPATH = "C:\Users\spenz\Desktop\cal2007.accdb"

    'Open the sample Northwind database that ships with
    'Access and display the application window.
   
      
    objAccess.OpenCurrentDatabase strPath
    objAccess.Visible = True
    
    ' Maximize other Access window and quit current application.
    objAccess.DoCmd.RunCommand acCmdAppMaximize
    Application.Quit

            
Exit_Now:
Exit Sub

Errhandler:
DisplayUnexpectedError Err.Number, Err.Description
Resume Exit_Now
Resume
End Sub

Attribute VB_Name = "modGlobals"
Option Compare Database
Option Explicit


Public booRequery As Boolean

Public Const gstrSaveMsg As String = "Gusto mo bang i-save ang pagbabago?" & vbNewLine _
& "Click Yes para mag-save or No para i-kansela ang pagbabago."

' String constant for application name
Public Const gstrTitle As String = "Ryan Christian"

'-Used to turn ON/OFF error handlder in all modules
Public Const gbooHandleErrors As Boolean = False
'True = production
'False = testing/developing

'-Used for resetting frmRecompute to default
Public Const gcdblDefaultPenaltyRate As Double = 0.03
Public gdblPenaltyRate As Double

Public Function GetPenaltyRate(Optional setValue As Double = 0.03) As Double
'--------------------------------------------------------------------
' Name:  TestFunction
' Purpose: Use to manipulate the result of qryRecompute based on changes in penalty rate.
' Inputs: GetPenaltyRate
' Returns: Penalty Rate
' Author: spenz
' Date:  8/17/2012
' Comment:
'--------------------------------------------------------------------

   If gdblPenaltyRate = 0 Then
      GetPenaltyRate = setValue
   Else
      GetPenaltyRate = gdblPenaltyRate
   End If

End Function


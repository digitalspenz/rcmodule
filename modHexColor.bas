Attribute VB_Name = "modHexColor"
Option Compare Database
Option Explicit

Public Function HexColor(strHex As String) As Long

    'converts Hex string to long number, for colors
    'the leading # is optional

    'example usage
    'Me.iSupplier.BackColor = HexColor("FCA951")
    'Me.iSupplier.BackColor = HexColor("#FCA951")

    'the reason for this function is to programmatically use the
    'Hex colors generated by the color picker.
    'The trick is, you need to reverse the first and last hex of the
    'R G B combination and convert to Long
    'so that if the color picker gives you this color #FCA951
    'to set this in code, we need to return CLng(&H51A9FC)
    
    Dim strColor As String
    Dim strR As String
    Dim strG As String
    Dim strB As String
    
    'strip the leading # if it exists
    If Left(strHex, 1) = "#" Then
        strHex = Right(strHex, Len(strHex) - 1)
    End If
    
    'reverse the first two and last two hex numbers of the R G B values
    strR = Left(strHex, 2)
    strG = Mid(strHex, 3, 2)
    strB = Right(strHex, 2)
    strColor = strB & strG & strR
    HexColor = CLng("&H" & strColor)

End Function

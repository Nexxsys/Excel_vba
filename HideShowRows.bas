Attribute VB_Name = "Module1"
Option Explicit

Public AdHoc As Worksheet
Public Plan As Worksheet

Public adhocrowend As Integer
Public adhoccolend As Integer
Public i As Integer


Public comp As String

Sub VarDefine()

Set Plan = Sheet1
Set AdHoc = Sheet2

adhocrowend = AdHoc.Cells(Rows.Count, 1).End(xlUp).Row

adhoccolend = AdHoc.Cells(9, Columns.Count).End(xlToLeft).Column

End Sub

Sub HideRows()

Application.Run "VarDefine"
Application.ScreenUpdating = False

For i = 1 To adhocrowend

    comp = LCase(Cells(i, 2).Value)
    
    If comp = "complete" Then
        Cells(i, 2).EntireRow.Hidden = True
    Else
    End If

Next i

End Sub

Sub ShowRows()

Application.Run "VarDefine"
Application.ScreenUpdating = False

For i = 1 To adhocrowend

    comp = LCase(Cells(i, 2).Value)
    
    If comp = "complete" Then
        Cells(i, 2).EntireRow.Hidden = False
    Else
    End If

Next i


End Sub

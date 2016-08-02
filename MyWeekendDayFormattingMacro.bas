Attribute VB_Name = "MyWeekendDayFormattingMacro"

Public Function CheckWeekend(MyDate As Variant)
    Dim strCell As String
    Dim DayNum As Variant
    Dim IsWeekEnd As Boolean
    DayNum = Application.Weekday(MyDate)
    If Not IsError(DayNum) Then
        Select Case DayNum
        Case 2 To 6 ' Monday thru Friday
            IsWeekEnd = False
        Case Else
            IsWeekEnd = True
        End Select
    Else
        IsWeekEnd = False
    End If
    'MsgBox IsWeekEnd
    CheckWeekend = IsWeekEnd
End Function

Sub FillDate()
    Dim rng As Range, cell As Range
    Set rng = Range("DateRange1") ' Named Range in the Spreadsheet (Row of Cells)
    For Each cell In rng
        If CheckWeekend(cell.Value) Then
            cell.Interior.Color = RGB(128, 128, 128)
        End If
    Next cell
End Sub

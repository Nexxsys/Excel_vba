Sub InsertSQLGeneration2()

'Dim cATC As String
'Dim cATCDec As String
Dim c As Range ' Columns
Dim r As Range ' Rows
Dim selectedRange As Range
Dim DataToOutput As String
Dim dataCapture As String
Dim myRow As Integer
Dim myCol As Integer
Dim MyFile As String
Dim tableName As String
Dim iTotalCellCount As Integer
Dim iCellCounter As Integer

'MsgBox("Select the range you want to output")
tableName = InputBox("Enter Table Name for SQL Inserts")

Set selectedRange = Selection
' ??? How to handloe Numeric vs Text Data ????
' ??? How to pull selection range and determine where you are at????

MyFile = "insertStmt2.txt"

'set and open file for output
fnum = FreeFile()
Open MyFile For Output As fnum

'iTotalCount = selectedRange.Cells.Count ' Count of the Number of Cells in the list
iTotalCellCount = selectedRange.Rows(1).Cells.Count ' Count of the Number of Cells in a row (i.e row 1)

For Each r In selectedRange.Rows
dataCapture = ""
dataCapture = "INSERT INTO " & tableName & "('ATC', 'ATC_DESC') "
dataCapture = dataCapture & "values('"
iCellCounter = 0
For Each c In r.Cells

iCellCounter = iCellCounter + 1

If iCellCounter = 1 Then
dataCapture = dataCapture & c.Value & "', "
Else
If iCellCounter = iTotalCellCount Then
dataCapture = dataCapture & "'" & c.Value
Else
dataCapture = dataCapture & "'" & c.Value & "', "
End If
End If
Next c

dataCapture = dataCapture & "');"
Print #fnum, dataCapture

Next r

Close #fnum
End Sub
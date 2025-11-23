# VBA/D1_*Excel VBA Programming Introduction*

Sub start()

Dim startrow As Integer
Dim lastrow As Long
Dim lastvalue As String

startrow = 2
lastrow = Sheet1.Range("b" & Rows.Count).End(xlUp).Row
lastvalue = Sheet1.Cells(lastrow, 1).Value

If lastvalue = "total" Then

MsgBox "total already calculated"

Else

Range("a1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.Value = "total"

ActiveCell.Offset(0, 1).Select
ActiveCell.Value = WorksheetFunction.Sum(Range("b" & startrow & ":" & "b" & lastrow))

End If

End Sub

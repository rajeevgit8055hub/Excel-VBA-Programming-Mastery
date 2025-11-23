# VBA/D14_*Project On If*

Sub start()

Dim startrow As Integer
<br>
Dim lastrow As Long
<br>
Dim lastvalue As String

startrow = 2
<br>
lastrow = Sheet1.Range("b" & Rows.Count).End(xlUp).Row
<br>
lastvalue = Sheet1.Cells(lastrow, 1).Value

If lastvalue = "total" Then

MsgBox "total already calculated"

Else

Range("a1").Select
<br>
Selection.End(xlDown).Select
<br>
ActiveCell.Offset(1, 0).Select
<br>
Selection.Value = "total"

ActiveCell.Offset(0, 1).Select
<br>
ActiveCell.Value = WorksheetFunction.Sum(Range("b" & startrow & ":" & "b" & lastrow))

End If

End Sub

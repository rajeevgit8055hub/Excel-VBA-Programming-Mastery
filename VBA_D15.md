Sub GitHub()

Dim i As Integer
<br>
For i = 1 To 10 Step 1

Sheet1.Cells(i, 1).Value = "GitHub"
<br>
Next i

End Sub

Sub GitHub_B()

Dim i As Integer
<br>
For i = 1 To 10 Step 2

Sheet1.Cells(i, 1).Value = "GitHub"
<br>
Next i

End Sub

Sub GitHub_No()

Dim i As Integer
Dim total As Integer

For i = 1 To 10

Sheet1.Cells(i, 1).Value = i

total = total + i

Next i

Sheet1.Range("a11").Value = total

End Sub

Sub Clear_data()

Sheet1.Range("a1:a1000").Value = ""

End Sub


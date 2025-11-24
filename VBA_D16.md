# VBA/D16_*For Loop Decrement*

Sub start()

Dim i As Integer

For i = 10 To 1 Step -1
<br>
Sheet1.Range("a" & i).Value = "vba"
<br>
Next i

End Sub

Sub clean_data()

Sheet1.Range("a1:a10").Value = ""

End Sub


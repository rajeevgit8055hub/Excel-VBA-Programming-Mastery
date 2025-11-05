## VBA/D3_*Sheets & Cell Selection*

Sub start()

Sheets("GitHUb").Activate
<br>
Range("B1").Select
<br>
Selection.Value = "GitHub Gist"

Sheets("Sheet1").Activate
<br>
Range("C5").Select
<br>
Selection.Value = "Excel_VBA"

Sheet3.Activate
<br>
Range("D2").Select
<br>
Selection.Value = 100000

Cells(1,5).Select
<br>
Selection.Value = 500000

End Sub


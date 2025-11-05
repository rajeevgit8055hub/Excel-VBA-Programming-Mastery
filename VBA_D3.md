## VBA/D3_Sheets & Cell Selection

Sub start()

Sheets("GitHUb").Activate
Range("B1").Select
Selection.Value = "GitHub Gist"

Sheets("Sheet1").Activate
Range("C5").Select
Selection.Value = "Excel_VBA"

Sheet3.Activate
Range("D2").Select
Selection.Value = 100000

Cells(1,5).Select
Selection.Value = 500000

End Sub


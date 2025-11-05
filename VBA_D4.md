# VBA/D4_*Dynamic Copy Code Of VBA*

Sub start()

Sheet1.Activate
<br>
Range("A1").Select
<br>
Range(Selection, Selection.End(XlToRight)).Select
<br>
Range(Selection, Selection.End(XlDown)).Select

Selection.Copy

Sheet2.Activate
<br>
Range("A1").Select

ActiveSheet.Paste

End Sub

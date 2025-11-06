## VBA/D6_*Project On Copy Paste*

Sub start()

'Activate Sheet
<br>
Sheets("Jan").Activate
<br>
Range("A1").Select
<br>
Range(Selection, Selection.End(XlToRight)).Select
<br>
Range(Selection, Selection.End(XlDown)).Select

'Copy Sheet
<br>
Selection.Copy

'Activate Target Sheet
<br>
Sheets("Final").Activate
<br>
Range("A1").Select

'Paste Sheet
<br>
ActiveSheet.Paste

Sheets("Feb").Activate
<br>
Range("A2").Select
<br>
Range(Selection, Selection.End(XlToRight)).Select
<br>
Range(Selection, Selection.End(XlDown)).Select
<br>
Selection.Copy

Sheets("Final").Activate
<br>
Range("A1").Select
<br>
Selection.End(XlDown).Select
<br>
ActiveCell.Offset(1,0).Select
<br>
ActiveSheet.Paste

Sheets("Mar").Activate
<br>
Range("A2").Select
<br>
Range(Selection, Selection.End(XlToRight)).Select
<br>
Range(Selection, Selection.End(XlDown)).Select

Selection.Copy

Sheets("Total").Activate
<br>
Range("A1").Select
<br>
Selection.End(XlDown).Select
<br>
ActiveCell.Offset(1,0).Select
<br>
ActiveSheet.Paste

Sheets("Apr").Activate
<br>
Range("A2").Select
<br>
Range(Selection, Selection.End(XlToRight)).Select
<br>
Range(Selection, Seletion.End(XlDown)).Select

Selection.Copy

Sheets("Total").Activate
<br>
Range("A1").Select
<br>
Selection.End(XlDown).Select
<br>
ActiveCell.Offset(1,0).Select
<br>
ActiveSheet.Paste

End Sub



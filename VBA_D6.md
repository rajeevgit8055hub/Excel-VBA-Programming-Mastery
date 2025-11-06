## VBA/D6_*Project On Copy Paste*

Sub start()

'Activate Sheet
Sheets("Jan").Activate
Range("A1").Select
Range(Selection, Selection.End(XlToRight)).Select
Range(Selection, Selection.End(XlDown)).Select

'Copy Sheet
Selection.Copy

'Activate Target Sheet
Sheets("Final").Activate
Range("A1").Select

'Paste Sheet
ActiveSheet.Paste

Sheets("Feb").Activate
Range("A2").Select
Range(Selection, Selection.End(XlToRight)).Select
Range(Selection, Selection.End(XlDown)).Select
Selection.Copy

Sheets("Final").Activate
Range("A1").Select
Selection.End(XlDown).Select
ActiveCell.Offset(1,0).Select
ActiveSheet.Paste

Sheets("Mar").Activate
Range("A2").Select
Range(Selection, Selection.End(XlToRight)).Select
Range(Selection, Selection.End(XlDown)).Select

Selection.Copy

Sheets("Total").Activate
Range("A1").Select
Selection.End(XlDown).Select
ActiveCell.Offset(1,0).Select
ActiveSheet.Paste

Sheets("Apr").Activate
Range("A2").Select
Range(Selection, Selection.End(XlToRight)).Select
Range(Selection, Seletion.End(XlDown)).Select

Selection.Copy

Sheets("Total").Activate
Range("A1").Select
Selection.End(XlDown).Select
ActiveCell.Offset(1,0).Select
ActiveSheet.Paste

End Sub



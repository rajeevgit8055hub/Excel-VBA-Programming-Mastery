Sub start()

Dim name as String
<br>
Dim age as Integer

name = InputBox("enter your name")
<br>
age = InputBox("enter your age")

Range("A1").Value = name
<br>
Range("B1").Value = age

MsgBox "data submitted successfully"
<br>
End Sub


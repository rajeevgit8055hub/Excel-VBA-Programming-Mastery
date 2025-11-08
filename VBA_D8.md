## VBA/D8_*Variable Types*

Global total As Integer   ' Public /Global Variable
<br>
'dim total as Integer     - Module Level

Sub marks()


' Local Variable

Dim english As Integer
<br>
Dim hindi As Integer
<br>
Dim math As Integer


english = 50
<br>
hindi = 60
<br>
math = 80

total = english + hindi + math

End Sub

Sub percentage()

Dim per As Single

per = total / 3

End Sub


'Module_2

Sub average()

Dim avg As Single

Call marks

avg = total / 3

End Sub


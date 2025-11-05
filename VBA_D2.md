# VBA/D3_**What Is Variable And Data Type**

Sub Start()

Dim name as String
<br>
Dim age as Integer
<br>
Dim dob as Date
<br>
Dim salary as Long
<br>
Dim bonus as double
<br>
Dim cgpa as Variant
<br>
Dim attendence as Boolean
<br>
Dim address as Variant

name = "Raj"
<br>
age = 25
<br>
dob = #15/05/2000#
<br>
salary = 25000
<br>
bonus = 5000
<br>
cgpg = CDec(2.5)
<br>
attendence = True
<br>
address = "Lucknow"

Range("A1").Value = name
<br>
Range("B1").Value = age
<br>
Range("c1").Value = dob
<br>
Range("d1").Value = salary
<br>
Range("E1").Value = bonus
<br>
Range("F1").Value = cgpa
<br>
Range("G1").Value = attendence
<br>
Range("H1").Value = address

End Sub

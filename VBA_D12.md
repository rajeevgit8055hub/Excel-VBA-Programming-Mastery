## VBA/D12_*If Condition With OR*

Sub start()

'if_or()
<br>
Dim age As Integer
<br>
Dim designation As String
<br>
age = Sheet1.Range("c2").Value
<br>
designation = Sheet1.Range("d2").Value
<br>
If ((age >= 35) Or (designation = "manager")) Then
<br>
Sheet1.Range("e2").Value = "yes"
<br>
Else
<br>
Sheet1.Range("e2").Value = "no"
<br>
End If
<br>
End Sub


# VBA/D5_*Offset Command*

sub start()

Range("A1").Select
<br>
Selection.End(xltoRight).Select
<br>
Activecell.offset(-5,0).Select

End sub

Public Sub Main()

Dim c As Range

 
For Each c In Range("N3:O40,V2:BR2")

If Range("V2:BR2").Text = Range("N3:O40").Text Then

MsgBox "Equal value : " & c.Text

Else

MsgBox "No"


End If

Next c
 
 
 
 
 
 
 
End Sub

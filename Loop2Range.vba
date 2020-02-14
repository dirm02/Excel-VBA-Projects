Public Sub Main()

Dim c As Range

 ' Looping through 2 Range 
For Each c In Range("N3:O40,V2:BR2")
 ' Checking if the 2 range have some Text in common
 ' Range can be compared with .Value property instead of .Text
If Range("V2:BR2").Text = Range("N3:O40").Text Then
  ' Test with a prompt box if there is a match
MsgBox "Equal value : " & c.Text

Else
  'Test has only showned a no match!
MsgBox "No"


End If

Next c
 
 
 
 
 
 
 
End Sub

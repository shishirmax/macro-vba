Attribute VB_Name = "FindAll"
Sub FindAll()
Dim c As Range
Dim findWhat As String
Dim i As Long
Dim lastrowA As Long
Dim lastrowC As Long

i = 2
lastrowA = Sheet1.Cells(Rows.Count, "A").End(xlUp).Row
lastrowC = Sheet1.Cells(Rows.Count, "C").End(xlUp).Row

'MsgBox (lastrowA)
findWhat = InputBox("Enter the name", "Find All")
If findWhat = "" Then
MsgBox "You Entered Nothing!"
Exit Sub
End If

With Worksheets(1).Range("A2:A" & lastrowA)
Range("C2:C" & lastrowC).Clear
Set c = .Find(findWhat, LookIn:=xlValues)
If Not c Is Nothing Then
firstAddress = c.Address

Do
Cells(i, 3) = c.Value
Set c = .FindNext(c)
i = i + 1
Loop While Not c Is Nothing And c.Address <> firstAddress
End If
End With



End Sub

Attribute VB_Name = "Module1"
Sub yearly_change()
Dim i As Integer
For i = 2 To 4999:
Cells(i, 9) = (Cells(i, 6).Value - Cells(i, 3).Value)

Next i
End Sub

Sub yearly_Percentage()
Dim i As Integer
For i = 2 To 4999:
Cells(i, 10) = ((Cells(i, 6).Value / Cells(i, 3).Value) * 100)

Next i
End Sub

Sub Ticker()
Dim i As Integer
For i = 2 To 4999
Cells(i, 8) = Cells(i, 1).Value
Next i
End Sub

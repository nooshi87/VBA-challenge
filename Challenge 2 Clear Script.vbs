Sub Reset()
Dim i As Double
Dim LastRow As Double
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 1 To LastRow
Cells(i, 9).Value = ""
Cells(i, 10).Value = ""
Cells(i, 11).Value = ""
Cells(i, 12).Value = ""
Cells(i, 16).Value = ""
Cells(i, 17).Value = ""
Cells(i, 10).Interior.ColorIndex = 2
Next i

End Sub

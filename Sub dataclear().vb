Sub dataclear()

Sheets("data").Select
Dim i As Integer
For i = 1 To 1000
    If Cells(i + 1, 1) <> "" Then Range(Cells(i + 1, 1), Cells(i, 11)).Select
    Selection.ClearContents
    Next i


End Sub


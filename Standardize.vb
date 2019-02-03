Option Explicit
Private Sub generate_Click()
Dim i As Integer
For i = 2 To 500
Cells(i, 1) = Application.WorksheetFunction.RandBetween(10, 1289)
Next i
End Sub
Private Sub standardize_Click()
Dim i As Double
Dim s As Double
Dim m As Double
s = Application.WorksheetFunction.StDev(Range("a1:a500"))
m = Application.WorksheetFunction.Average(Range("a1:a500"))
For i = 2 To 500
Cells(i, 2) = (Cells(i, 1) - m) / s
If Cells(i, 2) < -3 Or Cells(i, 2) > 3 Then  'Different font for those that are outliers i.e. above or below 3sd
Cells(i, 2).Font.ColorIndex = 4
End If
Next i

Range("a2:b500").Copy   ' Custom sort according to standardized values
Range("a2:b500").PasteSpecial Paste:=xlPasteValues
Range("A2:b500").Sort key1:=Range("b2"), order1:=xlAscending, _
                           Header:=xlNo, Orientation:=xlSortColumns, _
ordercustom:=Index + 1
End Sub
Private Sub clear_Click()
Worksheets("Sheet1").Range("A2:b500").clear
End Sub


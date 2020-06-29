Sub wallstreet()

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


Dim Tickername As String
Dim Yearlychange As Double
Dim Percentchange As Double
Dim Totalvol As Double
Totalvol = 0
Dim Yearclose As Double
Dim Yearopen As Double
Dim Summarytablerow As Integer
Summarytablerow = 2
Dim openprice As Integer
oprice = 2
Dim tincrease As String
Dim tdecrease As String
Dim ttotalvol As String

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

Totalvol = Totalvol + ws.Cells(i, 7).Value

If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then

Tickername = ws.Cells(i, 1).Value
ws.Range("I" & Summarytablerow).Value = Tickername

ws.Range("l" & Summarytablerow).Value = Totalvol
Totalvol = 0

Yearopen = ws.Range("C" & oprice)
Yearclose = ws.Range("F" & i)

Yearlychange = Yearclose - Yearopen
ws.Range("J" & Summarytablerow).Value = Yearlychange

If Yearopen = 0 Then
    Percentchange = 0
Else
Percentchange = Yearclose / Yearopen - 1
ws.Range("K" & Summarytablerow).Value = Percentchange

End If

If ws.Range("J" & Summarytablerow) >= 0 Then

ws.Range("J" & Summarytablerow).Interior.ColorIndex = 4

Else
ws.Range("J" & Summarytablerow).Interior.ColorIndex = 3

End If

Summarytablerow = Summarytablerow + 1
oprice = i + 1

End If

    Next i

ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Columns("K"))
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Columns("K"))
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Columns("L"))


lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
For j = 2 To lastrow2
If ws.Cells(j, 11).Value = ws.Range("Q2").Value Then
tincrease = ws.Cells(j, 9).Value

Exit For
End If
Next j

ws.Range("P2").Value = tincrease

For j = 2 To lastrow2
If ws.Cells(j, 11).Value = ws.Range("Q3").Value Then
tdecrease = ws.Cells(j, 9).Value

Exit For
End If
Next j

ws.Range("P3").Value = tdecrease

For j = 2 To lastrow2
If ws.Cells(j, 12).Value = ws.Range("Q4").Value Then
ttotalvol = ws.Cells(j, 9).Value

Exit For
End If
Next j
ws.Range("P4").Value = ttotalvol

Next ws

End Sub

Sub homework2()

Dim tcolumn As Long
tcolumn = 1

Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim tickernum As Integer
tickernum = 2

Dim currentopen As Double
currentopen = Range("C2").Value

Dim currentclose As Double
Dim change As Double
Dim percent As Double

Dim volume As Double
volume = 0

if range("I1").value = "" Then
range("I1").value = "Ticker"
range("J1").value = "Yearly Change"
range("K1").value = "Percent Change"
range("L1").value = "Total Stock Volume"
end if

For i = 2 To lastrow

volume = Cells(i, 7).Value + volume

If Cells(i, tcolumn).Value <> Cells(i + 1, tcolumn).Value Then

Cells(tickernum, 9).Value = Cells(i, tcolumn).Value

currentclose = Cells(i, 6).Value

percent = ((currentclose - currentopen) / currentopen) * 100
Cells(tickernum, 11).Value = percent

change = currentclose - currentopen
Cells(tickernum, 10).Value = change

If Cells(tickernum, 10).Value > 0 Then
Cells(tickernum, 10).Interior.ColorIndex = 4

ElseIf Cells(tickernum, 10) Then
Cells(tickernum, 10).Interior.ColorIndex = 3

End If

Cells(tickernum, 12).Value = volume
volume = 0

currentopen = Cells(i + 1, 3).Value
tickernum = tickernum + 1
End If

Next i

End Sub

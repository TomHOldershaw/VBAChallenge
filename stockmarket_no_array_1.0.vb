Sub stockmarket()

Dim ws As Worksheet
Dim LastRow As Long
Dim TickerIndex As Integer
Dim TickerName As String
Dim Opening As Double
Dim Closing As Double
Dim Volume As LongLong
Dim YearlyChange As Double
Dim PercentChange As Double
Dim colour As Integer
Dim Percentchangemin As Double
Dim Percentchangemax As Double
Dim volumemax As LongLong
Dim pcmaxtick As String
Dim pcmintick As String
Dim volmaxtick As String

For Each ws In Worksheets

'Setup headers
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly change"
ws.Range("K1") = "Percent change"
ws.Range("L1") = "Total stock volume"
'Setup greatest values table
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Check range
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'TickerIndex is 0
TickerIndex = 0
'Reset max and mins
Percentchangemax = 0
Percentchangemin = 0
volumemax = 0

'Loop through stocks, start at row 2.
For Row = 2 To LastRow
  'Test if new ticker
  'If new ticker, set starting values and calculate values from previous ticker
  If ws.Cells(Row, 1) <> ws.Cells(Row - 1, 1) Then
    'Grab current Index
    PreviousIndex = TickerIndex
    'Advance TickerIndex
    TickerIndex = TickerIndex + 1
    'Set ticker name
    ws.Cells(TickerIndex + 1, 9) = ws.Cells(Row, 1).Value
    'Set opening price
    Opening = ws.Cells(Row, 3).Value
    'Set volume
    Volume = ws.Cells(Row, 7).Value
  'Last row, do closing value stuff
  ElseIf ws.Cells(Row, 1) <> ws.Cells(Row + 1, 1) Then
    'Closing value
    Closing = ws.Cells(Row, 6).Value
    'Calculate changes of previous index
    'Yearly change = closing - opening
    YearlyChange = Closing - Opening
    ws.Cells(TickerIndex + 1, 10) = YearlyChange
    'If change > 0, set colour to green
    If YearlyChange > 0 Then colour = 4
    'If change < 0, set colour to red
    If YearlyChange < 0 Then colour = 3
    'Set colour
    ws.Cells(TickerIndex + 1, 10).Interior.ColorIndex = colour
    'Percent change
    If Opening <> 0 Then
      PercentChange = Round(((YearlyChange / Opening) * 100), 2)
      ws.Cells(TickerIndex + 1, 11) = PercentChange & "%"
    End If
    'Add volume to count
    Volume = Volume + ws.Cells(Row, 7).Value
    ws.Cells(TickerIndex + 1, 12) = Volume
  'Normal row
  Else
    'Add volume to count
    Volume = Volume + ws.Cells(Row, 7).Value
End If

If PercentChange > Percentchangemax Then
   Percentchangemax = PercentChange
   pcmaxtick = ws.Cells(Row, 1).Value
End If
If PercentChange < Percentchangemin Then
   Percentchangemin = PercentChange
   pcmintick = ws.Cells(Row, 1).Value
End If
If Volume > volumemax Then
   volumemax = Volume
   volmaxtick = ws.Cells(Row, 1).Value
End If

Next Row

ws.Range("Q2") = Percentchangemax & "%"
ws.Range("Q3") = Percentchangemin & "%"
ws.Range("Q4") = volumemax
ws.Range("P2") = pcmaxtick
ws.Range("P3") = pcmintick
ws.Range("P4") = volmaxtick


'Autofit column width
ws.Columns("I:Q").EntireColumn.AutoFit

Next ws

End Sub

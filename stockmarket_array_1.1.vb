Sub stockmarket()

Dim ws As Worksheet
Dim LastRow As Long
Dim TickerIndex As Integer
Dim TickerName() As String
Dim Opening() As Double
Dim Closing() As Double
Dim Volume() As LongLong
Dim YearlyChange() As Double
Dim PercentChange() As Double
Dim colour As Integer

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
'Reset arrays
ReDim TickerName(1)
ReDim Opening(1)
ReDim Closing(1)
ReDim Volume(1)
ReDim YearlyChange(1)
ReDim PercentChange(1)

'Loop through stocks, start at row 2.
For Row = 2 To LastRow
  'Test if new ticker
  'If new ticker, set starting values
  If ws.Cells(Row, 1) <> ws.Cells(Row - 1, 1) Then
    'Advance TickerIndex
    TickerIndex = TickerIndex + 1
    'Redimension arrays
    ReDim Preserve TickerName(TickerIndex)
    ReDim Preserve Opening(TickerIndex)
    ReDim Preserve Closing(TickerIndex)
    ReDim Preserve Volume(TickerIndex)
    ReDim Preserve YearlyChange(TickerIndex)
    ReDim Preserve PercentChange(TickerIndex)
    'Set ticker name
    TickerName(TickerIndex) = ws.Cells(Row, 1).Value
    'Set opening price
    Opening(TickerIndex) = ws.Cells(Row, 3).Value
    'Set volume
    Volume(TickerIndex) = ws.Cells(Row, 7).Value
  'Last row, do closing value stuff
  ElseIf ws.Cells(Row, 1) <> ws.Cells(Row + 1, 1) Then
    'Closing value
    Closing(TickerIndex) = ws.Cells(Row, 6).Value
    'Calculate changes of previous index
    'Yearly change = closing - opening
    YearlyChange(TickerIndex) = Closing(TickerIndex) - Opening(TickerIndex)
    'Percent change
    If Opening(TickerIndex) <> 0 Then PercentChange(TickerIndex) = Round(((YearlyChange(TickerIndex) / Opening(TickerIndex)) * 100), 2)
    'Add volume to count
    Volume(TickerIndex) = Volume(TickerIndex) + ws.Cells(Row, 7).Value
  'Normal row
  Else
    'Add volume to count
    Volume(TickerIndex) = Volume(TickerIndex) + ws.Cells(Row, 7).Value
End If

Next Row

'Output values
'End stop is TickerIndex
For x = 1 To TickerIndex
  Row = x + 1
  'Ticker Name
  ws.Cells(Row, 9) = TickerName(x)
  ws.Cells(Row, 10) = YearlyChange(x)
  'If change > 0, set colour to green
  If YearlyChange(x) > 0 Then colour = 4
  'If change < 0, set colour to red
  If YearlyChange(x) < 0 Then colour = 3
  'Set colour
  ws.Cells(Row, 10).Interior.ColorIndex = colour
  'Percent change
  ws.Cells(Row, 11).Value = PercentChange(x) & "%"
  'Total stock volume
  ws.Cells(Row, 12) = Volume(x)

'Greatest values
 If PercentChange(x) = Application.Max(PercentChange) Then
    ws.Range("P2").Value = TickerName(x)
    ws.Range("Q2").Value = PercentChange(x) & "%"
 ElseIf PercentChange(x) = Application.Min(PercentChange) Then
    ws.Range("P3").Value = TickerName(x)
    ws.Range("Q3").Value = PercentChange(x) & "%"
 End If
 If Volume(x) = Application.Max(Volume) Then
    ws.Range("P4").Value = TickerName(x)
    ws.Range("Q4").Value = Volume(x)
 End If

Next x

'Autofit column width
ws.Columns("I:Q").EntireColumn.AutoFit

Next ws

End Sub

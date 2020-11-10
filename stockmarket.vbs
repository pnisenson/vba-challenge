Sub StockMarket():
'Create a For Loop to apply code to all worksheets
For Each ws In Worksheets

'Create headers for solutions
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Create variable to count unique ticker symbols
Dim TickCount As Integer
TickCount = 0
'Create variable to track all rows for a given ticker to count volume
Dim VolCount As Double
VolCount = 0
'Create variable to find the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create a For Loop to search for unique tickers by identifying when a cell is different from the cell above it and return in row I
For a = 2 To LastRow
  'Defines Open and Closing price as variables
  Dim Open1 As Double
  Dim Close1 As Double

If ws.Cells(a, 1).Value <> ws.Cells(a - 1, 1).Value Then
    'Counts a unique ticker when changes from the cell above it
    TickCount = TickCount + 1
    'Defines the opening value of the first row of a new ticker
    Open1 = ws.Cells(a, 3).Value
    'Puts Ticker symbols in column 9
    ws.Cells(1 + TickCount, 9).Value = ws.Cells(a, 1).Value
End If
'Defines closing value of the last row of a ticker
If ws.Cells(a, 1).Value <> ws.Cells(a + 1, 1).Value Then
        Close1 = ws.Cells(a, 6).Value
        'Adds up the volume until the end of the ticker is reached, then resets the counter to zero
        VolCount = VolCount + ws.Cells(a, 7).Value
        ws.Cells(1 + TickCount, 12).Value = VolCount
        VolCount = 0
Else: VolCount = VolCount + ws.Cells(a, 7).Value
End If
   'Define yearly change as variable'
   Dim Change As Double
   Change = Close1 - Open1
   'Run formulas of yearly change and percentage change in columns 10 and 11
   ws.Cells(1 + TickCount, 10) = Change
   'Change % needs to include If formulas in case the denominator is zero
   If Open1 = 0 And Close1 = 0 Then
    ws.Cells(1 + TickCount, 11) = 0
   ElseIf Open1 = 0 And Close1 <> 0 Then
    ws.Cells(1 + TickCount, 11) = 1
    Else: ws.Cells(1 + TickCount, 11) = Change / Open1
    End If
   'Set number formatting for percentage change column
   ws.Cells(1 + TickCount, 11).NumberFormat = "#0.00%"
'Set conditional formatting for percentage change column
If ws.Cells(1 + TickCount, 11).Value > 0 Then
    ws.Cells(1 + TickCount, 11).Interior.ColorIndex = 4
ElseIf ws.Cells(1 + TickCount, 11).Value < 0 Then
    ws.Cells(1 + TickCount, 11).Interior.ColorIndex = 3
    End If
    
Next a
Next ws
Call Bonus
End Sub
'------------------------------------------------------------------
Sub Bonus():

'Create a For Loop to apply code to all worksheets
For Each ws In Worksheets

'Create headers for solution
ws.Cells(1, 14).Value = "Bonus Table"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

'Create Variable to find the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create variables to store Max/Min values and associated ticker symbol
Dim MaxPerc As Double
    MaxPerc = 0
Dim MaxTick As String
Dim MinPerc As Double
    MinPerc = 0
Dim MinTick As String
Dim MaxVol As Double
    MaxVol = 0
Dim MaxVolTick As String

'Create a loop to search in the table summarized by ticker
For r = 2 To LastRow
   'Create conditional to set Max Percentage as it finds the Max in the column
    If ws.Cells(r, 11).Value > MaxPerc Then
        MaxPerc = ws.Cells(r, 11).Value
        MaxTick = ws.Cells(r, 9).Value
        ws.Cells(2, 16).Value = MaxPerc
        ws.Cells(2, 15).Value = MaxTick
    End If
    'Create conditional to set Min Percentage as it finds the Min in the column
    If ws.Cells(r, 11).Value < MinPerc Then
        MinPerc = ws.Cells(r, 11).Value
        MinTick = ws.Cells(r, 9).Value
        ws.Cells(3, 16).Value = MinPerc
        ws.Cells(3, 15).Value = MinTick
    End If
    'Create conditional to set Max Volume as it finds the Max in the column
    If ws.Cells(r, 12).Value > MaxVol Then
        MaxVol = ws.Cells(r, 12).Value
        MaxVolTick = ws.Cells(r, 9).Value
        ws.Cells(4, 16).Value = MaxVol
        ws.Cells(4, 15).Value = MaxVolTick
    End If

Next r

ws.Cells(2, 16).NumberFormat = "#0.00%"
ws.Cells(3, 16).NumberFormat = "#0.00%"
ws.Range("J:P").EntireColumn.AutoFit


Next ws

End Sub

# VBA-challenge

Sub Stock_market()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets


'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

'Create variable to hold stock volume
'Dim stock_volume As Double
'stock_volume = 0

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer


Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0

Dim TickerRow As Long: TickerRow = 1


For i = 2 To Lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
TickerRow = TickerRow + 1
Ticker = ws.Cells(i, 1).Value
ws.Cells(TickerRow, "I").Value = Ticker

'
close_price = ws.Cells(i, 6).Value
price_change_percent = close_price - open_price


ElseIf open_price <> 0 Then
price_change_percent = (price_change_percent / open_price) * 100

End If

Next i

Next ws

End Sub

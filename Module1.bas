Attribute VB_Name = "Module1"
Sub stocks()

'declare variables
Dim r As Double
Dim output_row As Integer
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim stock_volume As Double
Dim yr_change As Double
Dim percent_change As Double
Dim current_stock As String
Dim next_stock As String
Dim total_rows As Double
Dim most_increase As Double
Dim most_decrease As Double
Dim most_volume As Double
Dim most_increase_ticker As String
Dim most_decrease_ticker As String
Dim most_volume_ticker As String

'run script for each worksheet
For Each ws In Worksheets

'write output column headers and autofit
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
Columns("I:L").AutoFit

'find total number of rows
total_rows = ws.Cells(1, 1).End(xlDown).Row

'initialize output row and first stock ticker and open price
output_row = 2
open_price = ws.Cells(2, 3).Value
ticker = ws.Cells(2, 1).Value
most_increase = 0
most_decrease = 0
most_volume = 0

'loop through all rows
For r = 2 To total_rows
    
    'calculate cumulative volume for current stock
    stock_volume = stock_volume + ws.Cells(r, 7).Value

    'set current and next stock tickers
    current_stock = ws.Cells(r, 1).Value
    next_stock = ws.Cells(r + 1, 1).Value

    'compare ticker names for current and next rows to find changes
    If current_stock <> next_stock Then
        
        'calculate Yearly Change by subtracting open price from close
        yr_change = ws.Cells(r, 6).Value - open_price

        'calculate Percent Change only if open price is not 0
        'if open price is 0, set Percent Chnage to 0 to avoid dividing by 0
        If open_price <> 0 Then
            percent_change = yr_change / open_price
            Else: percent_change = 0
        End If
        
        'output Ticker, Yearly Change, Percent Change, and Cumulative Volume
        ws.Cells(output_row, 9).Value = current_stock
        ws.Cells(output_row, 10).Value = yr_change
        ws.Cells(output_row, 11).Value = percent_change
        ws.Cells(output_row, 12).Value = stock_volume
        
        'format cell color
        If yr_change > 0 Then
            ws.Range("J" & output_row).Interior.ColorIndex = 4
        Else: ws.Range("J" & output_row).Interior.ColorIndex = 3
        End If
            
        'format Percent Change cell
        ws.Range("K" & output_row).NumberFormat = "0.00%"
        
        'find greatest volume
        If stock_volume > most_volume Then
            most_volume = stock_volume
            most_volume_ticker = current_stock
        End If
        
        'find greatest increase
        If percent_change > most_increase Then
            most_increase = percent_change
            most_increase_ticker = current_stock
        End If
        
        'find greatest decrease
        If percent_change < most_decrease Then
            most_decrease = percent_change
            most_decrease_ticker = current_stock
        End If
        
        'reset volume to 0 for next stock
        stock_volume = 0

        'set open price for next stock
        open_price = ws.Cells(r + 1, 3).Value

        'increment output row
        output_row = output_row + 1
    End If
Next r

'output greatest values
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Volume"
ws.Range("P2").Value = most_increase_ticker
ws.Range("Q2").Value = most_increase
ws.Range("P3").Value = most_decrease_ticker
ws.Range("Q3").Value = most_decrease
ws.Range("P4").Value = most_volume_ticker
ws.Range("Q4").Value = most_volume

'format cells
ws.Columns("O:Q").AutoFit
ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next ws


End Sub



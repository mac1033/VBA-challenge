# VBA-challenge
## The VBA of Wall Street

VBA scripting was used to analyize generated stock market data. This script loops through all the stocks for one year and outputs the following information:
   
   - Ticker symbol
   
   - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
   
   - The percent change from opening price at the beginning of a given year to the closing price at the end of the year
   
   - The total stock volume, greatest % increase, greatest % decrease, and greatest total volume
   
   
  The VBA code:
  Sub Multiple_year_stock_data()

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
    
    Dim ticker_name As String
    Dim last_row As Long
    Dim total_ticker_volume As Double
    total_ticker_volume = 0
    Dim summary_table_row As Long
    summary_table_row = 2
    Dim yearly_open As Double
    Dim yearly_close As Double
    Dim yearly_change As Double
    Dim previous_amount As Long
    previous_amount = 2
    Dim percent_change As Double
    Dim greatest_increase As Double
    greatest_increase = 0
    Dim greatest_decrease As Double
    greatest_decrease = 0
    Dim last_row_value As Long
    Dim greatest_total_volume As Double
    greatest_total_volume = 0
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        
        total_ticker_volume = total_ticker_volume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker_name = ws.Cells(i, 1).Value
        ws.Range("I" & summary_table_row).Value = ticker_name
        ws.Range("L" & summary_table_row).Value = total_ticker_volume
        total_ticker_volume = 0
        
        yearly_open = ws.Range("C" & previous_amount)
        yearly_close = ws.Range("F" & i)
        yearly_change = yearly_close - yearly_open
        ws.Range("J" & summary_table_row).Value = yearly_change
        
        If yearly_open = 0 Then
        percent_change = 0
        
        Else
            yearly_open = ws.Range("C" & previous_amount)
            percent_change = yearly_change / yearly_open
        End If
        
        ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
        ws.Range("K" & summary_table_row).Value = percent_change
        
        If ws.Range("J" & summary_table_row).Value >= 0 Then
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
        Else
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
        End If
        
        summary_table_row = summary_table_row + 1
        previous_amount = i + 1
        End If
    
    Next i
    
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To LastRow
        If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
            ws.Range("Q2").Value = ws.Range("K" & i).Value
            ws.Range("P2").Value = ws.Range("I" & i).Value
        End If
        
        If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Range("K" & i).Value
            ws.Range("P3").Value = ws.Range("I" & i).Value
        End If
        
        If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
            ws.Range("Q4").Value = ws.Range("L" & i).Value
            ws.Range("P4").Value = ws.Range("I" & i).Value
        End If
        
        Next i
        
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ws.Columns("I:Q").AutoFit
    
    Next ws
    
End Sub

The Stock data for 2018:![4E2EDBB8-79E5-449A-9588-31C98E7630A7](https://user-images.githubusercontent.com/119431420/210885473-dd8ede1a-247e-436f-beb8-c167344d4f82.jpeg)

The Stock data for 2019:![EC3AB2FD-9184-48BC-95C4-D6AFDD2EEB79](https://user-images.githubusercontent.com/119431420/210885602-37807688-12dc-4974-b752-f20e96b3ae09.jpeg)

The Stock data for 2020:![40353802-E907-460D-B33E-1074475D1875](https://user-images.githubusercontent.com/119431420/210885647-8dbea08b-c715-421f-8e17-ae6210814819.jpeg)

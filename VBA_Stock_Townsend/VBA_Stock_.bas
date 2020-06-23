Attribute VB_Name = "Module1"

'Attribute VB_Challenge = "Module1"


Sub Calculate_Stock_Challenge_Townsend():

' variable to keep track of current ticker symbol
 'variable to keep track of number of tickers for each worksheet
 'variable to keep track of the last row in each worksheet.
 ' variable to keep track of opening price for specific year
 ' variable to keep track of closing price for specific year
 ' variable to keep track of yearly change
 ' variable to keep track of percent change
 ' variable to keep track of total stock volume
 ' variable to keep track of greatest percent increase value for specific year.
 ' variable to keep track of the ticker that has the greatest percent increase.
 ' varible to keep track of the greatest percent decrease value for specific year.
 ' variable to keep track of the ticker that has the greatest percent decrease.
 ' variable to keep track of the greatest stock volume value for specific year.
 ' variable to keep track of the ticker that has the greatest stock volume.
 
 
 
Dim ticker As String
Dim number_tickers As Integer
Dim lastRowState As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease As Double
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume As Double
Dim greatest_stock_volume_ticker As String

' loop over each worksheet in the workbook
For Each ws In Worksheets

    'Activate the worksheet.
    ws.Activate

    ' Last row of each worksheet
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add header columns for each worksheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Variables for each worksheet.
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    For i = 2 To lastRowState

        
        ticker = Cells(i, 1).Value
        
        ' Start of the year opening price for the ticker.
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        ' Sum of total stock volume values for a ticker.
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        ' Run for a different ticker in the list.
        If Cells(i + 1, 1).Value <> ticker Then
            ' Number of tickers when a different ticker in the list.
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            ' Year end closing price for ticker
            closing_price = Cells(i, 6)
            
            ' Yearly change value
            yearly_change = closing_price - opening_price
            
            ' Sum of yearly change value to the cell in each worksheet.
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' If yearly value is greater than 0, shade cell green.
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ' If yearly value is less than 0, shade cell red.
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            ' If yearly value is 0, shade cell yellow.
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Percent change of value for each ticker.
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            ' Percent_change value as a percent.
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            
            ' Set opening price back to 0 to get a different ticker in the list.
            opening_price = 0
            
            ' Sum of total stock volume value to the cell in each worksheet.
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' Set total stock volume back to 0 to get a different ticker in the list.
            total_stock_volume = 0
        End If
        
    Next i
    
    ' Greatest percent increase, greatest percent decrease, and greatest total volume of each year.
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Last row
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Variables and values of variables for the first row in the list.
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    ' skipping the head, loop through tickers.
    For i = 2 To lastRowState
    
        ' Ticker with the greatest percent increase.
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        ' Ticker with the greatest percent decrease.
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        ' Ticker with the greatest stock volume.
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Sum of values for greatest percent increase, decrease, and stock volume for each worksheet.
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws


End Sub

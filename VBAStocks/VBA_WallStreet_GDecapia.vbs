Attribute VB_Name = "Module1"
'************************
'The VBA of Wall Street
'************************
'(Assumption: Data is sorted by Ticker, then by date in ascending order)

Sub MultiYearStockData()

    'declare and initialize counters and accumulators
    Dim RowNum As Long
    Dim TickerCounter As Integer
    Dim LastRow As Long
    Dim CloseValue As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double
    
    CloseValue = 0
    YearlyChange = 0
    PercentChange = 0
    
    'loop through all worksheets
    For Each ws In Worksheets
    
        'initialize ticket counter at the start of every sheet
        TickerCounter = 1
        
        'print summary header
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Stock Volume"

        'determine number of rows in each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'LastRow = ws.UsedRange.Rows.Count
        
        'store initial values for each ticker set of each sheet
        TickerValue = ws.Cells(2, 1).Value  'NEW
        OpenValue = ws.Cells(2, 3).Value
        StockVolume = ws.Cells(2, 7).Value
        
        'loop through each row
        For RowNum = 2 To LastRow
            
            If ws.Cells(RowNum, 1).Value <> ws.Cells(RowNum + 1, 1).Value Then
               
               'ticker summary prep
               CloseValue = ws.Cells(RowNum, 6).Value
               YearlyChange = CloseValue - OpenValue
               
               'erroneous data handling at the end of each ticker dataset; this prevents div/0
               If OpenValue = 0 Then
                  YearlyChange = 0
                  PercentChange = 0
               Else
                  PercentChange = (CloseValue - OpenValue) / OpenValue
               End If
               
               'display each ticker summary
               ws.Cells(TickerCounter + 1, 9).Value = TickerValue
               ws.Cells(TickerCounter + 1, 10).Value = YearlyChange
               ws.Cells(TickerCounter + 1, 11).Value = PercentChange
               ws.Cells(TickerCounter + 1, 12).Value = StockVolume
               
               'format numeric cells
               ws.Cells(TickerCounter + 1, 10).NumberFormat = "###,##0.00"
               ws.Cells(TickerCounter + 1, 11).NumberFormat = "###,##0.00%"
               If YearlyChange < 0 Then
                  ws.Cells(TickerCounter + 1, 10).Interior.ColorIndex = 3
               Else
                  ws.Cells(TickerCounter + 1, 10).Interior.ColorIndex = 4
               End If
               
               'display/add new ticker row; reset accumulators
               TickerCounter = TickerCounter + 1
               TickerValue = ws.Cells(RowNum + 1, 1).Value  ' next ticker value
               
               OpenValue = ws.Cells(RowNum + 1, 3).Value
               StockVolume = ws.Cells(RowNum + 1, 7).Value
               
            Else
               'find first non-zero Open Value if first row in a ticker dataset is zero
               If OpenValue = 0 Then
                  OpenValue = ws.Cells(RowNum + 1, 3).Value
               End If
               
               'accumulate Stock Volume value for each ticket set
               StockVolume = StockVolume + ws.Cells(RowNum + 1, 7).Value
               
            End If
        
        Next RowNum
        
        ws.Columns("I:L").AutoFit
    
        '*******************
        'CHALLENGES section*
        '*******************
    
        'initialize greatest percent increase/decrease, total volume
        GreatestPercentIncrease = 0
        GreatestPercentDecrease = 0
        GreatestTotalVolume = 0
    
        'print summary header
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'determine the row count of column K
        'this will used to search for max/min values on the non-empty cells only
        LastPercentRow = Application.WorksheetFunction.Count(ws.Range("K2:K1048576"))
        
        'search for the maximum value on Percent Change column
        GreatestPercentIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastPercentRow))
    
        'search for the minimum value on Percent Change column
        GreatestPercentDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastPercentRow))
    
        'search for the greatest value on Stock Volume column
        GreatestTotalVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastPercentRow))
    
        'read each line on columns I to L to determine the corresponding ticker of the top Percent and Stock Volume values
        For SummaryRowNum = 2 To LastPercentRow
            If ws.Cells(SummaryRowNum, 11) = GreatestPercentIncrease Then
               ws.Cells(2, 16).Value = ws.Cells(SummaryRowNum, 9).Value
               ws.Cells(2, 17).Value = GreatestPercentIncrease
               ws.Cells(2, 17).NumberFormat = "###,##0.00%"
            End If
   
            If ws.Cells(SummaryRowNum, 11) = GreatestPercentDecrease Then
               ws.Cells(3, 16).Value = ws.Cells(SummaryRowNum, 9).Value
               ws.Cells(3, 17).Value = GreatestPercentDecrease
               ws.Cells(3, 17).NumberFormat = "###,##0.00%"
            End If
    
            If ws.Cells(SummaryRowNum, 12) = GreatestTotalVolume Then
               ws.Cells(4, 16).Value = ws.Cells(SummaryRowNum, 9).Value
               ws.Cells(4, 17).Value = GreatestTotalVolume
               ws.Cells(4, 17).NumberFormat = "###,###,##0"
            End If
    
        Next SummaryRowNum
    
        ws.Columns("O:Q").AutoFit
    
    Next ws
        
End Sub


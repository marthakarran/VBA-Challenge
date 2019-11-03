' Homework 02 - VBA
' Create a script that will loop through all the stocks for one year for each run and take the following information.
    ' The ticker symbol.
    ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' The total stock volume of the stock.
' You should also have conditional formatting that will highlight positive change in green and negative change in red.
'CHALLENGES 
    ' Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume". 

sub VBA_test()

Dim WS As Worksheet

    for each WS In ActiveWorkbook.Worksheets
    WS.activate 

        'Locate last row
        LastRow = WS.cells(Rows.Count,1).End(xlup).Row 

        'Create headings for the summary cells
        cells(1,"I").value = "Ticker"
        cells(1,"J").value = "YearlyChange" 
        cells(1,"K").value = "PercentChange"
        cells(1,"L").value = "TotalStockVolume"

        'Variables to hold values
        dim TickerName as string 
        dim OpeningPrice as double
        dim ClosingPrice as double
        dim YearlyChange as double
        dim PercentChange as double
        dim Volume as double 
        dim Row as double
        dim Column as integer 
        Dim i as long 

        'Set "starting points" 
        Volume = 0
        Row = 2
        Column = 1 

        'Determine where to find opening price (i.e., what opening price is)
        OpeningPrice = cells(2,Column+2) 

        for i = 2 to LastRow
            if cells(i + 1,Column).value <> cells(i,Column).value then 
                
                ' Ticker Name 
                TickerName = cells(i,Column).value
                cells(Row,Column + 8).value = TickerName
                
                ' Closing Price 
                ClosingPrice = cells(i,Column + 5).value
                
                ' Yearly Change 
                YearlyChange = ClosingPrice - OpeningPrice
                cells(Row,Column + 9).Value = YearlyChange
                
                ' Percent Change 
                if (OpeningPrice = 0 and ClosingPrice = 0) then 
                    PercentChange = 0 
                elseif (OpeningPrice = 0 and ClosingPrice <> 0) then 
                    PercentChange = 1 
                else 
                    PercentChange = YearlyChange / OpeningPrice
                    cells(Row, Column + 10).value = PercentChange
                    cells(Row, Column + 10).NumberFormat = "0.00%"
                end if 
                
                'Total Volumne 
                Volume = Volume + cells(i,Column + 6).value 
                cells(Row,Column + 11).value = Volume 
                
                ' Move to next row 
                Row = Row + 1 
                
                ' Reset 
                OpeningPrice = cells(i + 1, Column + 2) 
                Volume = 0
            else 
                Volume = Volumne + cells(i, Column + 6).value  
            end if 
        next i   
        
        ' Last row in Yearly Change 
        YearlyChangeLastRow = WS.cells(Rows.count, Column + 8).End(xlup).Row

        ' Change cell colors 
        for j = 2 to YearlyChangeLastRow
            if (cells(j, Column + 9).value > 0 or cells(j, Column + 9).value = 0) then 
                cells(j, Column + 9).Interior.ColorIndex = 4
            elseif cells(j, Column + 9).value < 0 then 
                cells(j, Column + 9).Interior.ColorIndex = 3
            end if 
        next j 

        ' Variables (Greatest Percent Increase, Greatest Percent Decrease, Greatest Total Volume)        
        cells(2, Column + 14).value = "GreatestPercentIncrease"
        cells(3, Column + 14).value = "GreatestPercentDecrease"
        cells(4, Column + 14).value = "GreatestTotalVolume"
        cells(1, Column + 15).value = "Ticker"
        cells(1, Column + 16).value = "Value"

        ' Find greatest value and respective ticker name 
        for k = 2 to YearlyChangeLastRow
            if cells(k, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                cells(2, Column + 15).Value = cells(k, Column + 8).Value
                cells(2, Column + 16).Value = cells(k, Column + 10).Value
                cells(2, Column + 16).NumberFormat = "0.00%"
            elseif cells(k, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                cells(3, Column + 15).Value = cells(k, Column + 8).Value
                cells(3, Column + 16).Value = cells(k, Column + 10).Value
                cells(3, Column + 16).NumberFormat = "0.00%"
            elseif cells(k, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YearlyChangeLastRow)) Then
                cells(4, Column + 15).Value = cells(k, Column + 8).Value
                cells(4, Column + 16).Value = cells(k, Column + 11).Value
            end if
        next k

    next WS 

End sub 


Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information.
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.


Sub StockMarket()

'Define variables

Dim TickerSymbol As String
Dim YearlyChange As Double
Dim PercentageChange As Double
Dim TotalVolume As Double
Dim FirstOpenPrice As Double
Dim LastClosePrice As Double
TotalVolume = 0
FirstOpenPrice = 0

'Keep track of the location of each ticker in a summary table
Dim SummaryTableRow As Integer
SummaryTableRow = 2

'Set a variable to get the last row

last_row = Cells(Rows.Count, 1).End(xlUp).Row

'Headers & Titles

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"


'Loop trough tickers' data

For i = 2 To last_row
    
    If FirstOpenPrice = 0 Then
        FirstOpenPrice = FirstOpenPrice + Cells(i, 3).Value
    End If
    
       
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        TickerSymbol = Cells(i, 1).Value
        LastCloseValue = Cells(i, 6).Value
        TotalVolume = TotalVolume + Cells(i, 7).Value
        YearlyChange = LastCloseValue - FirstOpenPrice
        If FirstOpenPrice <> 0 Then
            PercentageChange = LastCloseValue / FirstOpenPrice - 1
        End If
    
        'Print data in the summary table
        Range("I" & SummaryTableRow).Value = TickerSymbol
        Range("J" & SummaryTableRow).Value = YearlyChange
        Range("K" & SummaryTableRow).Value = Format(PercentageChange, "Percent")
        Range("L" & SummaryTableRow).Value = TotalVolume
                
        If PercentageChange > 0 Then
            Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        Else
            Range("J" & SummaryTableRow).Interior.ColorIndex = 3
        End If
        
        'Add one to summary table row
        SummaryTableRow = SummaryTableRow + 1
      
        'Reset Total Volume
        TotalVolume = 0
        FirstOpenPrice = 0
          
    Else
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
    End If
        
Next i

    'Find Max Value of Percentage Change
    
    LastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    GreatestIncrease = Application.WorksheetFunction.Max(Range("K:K"))
    GreatestDecrease = Application.WorksheetFunction.Min(Range("K:K"))
    GreatestTotalVolume = Application.WorksheetFunction.Max(Range("L:L"))
    
    'Assign Tickers to the values
    
    For i = 2 To LastRow
    
        If Cells(i, 11) = GreatestIncrease Then
            TickerGreatestIncrease = Cells(i, 9).Value
        ElseIf Cells(i, 11) = GreatestDecrease Then
            TickerGreatestDecrease = Cells(i, 9).Value
        ElseIf Cells(i, 12) = GreatestTotalVolume Then
            TickerGreatestTotalVolume = Cells(i, 9).Value
        End If
        
    Next i

    'Print Values
    
    Range("Q2") = GreatestIncrease
    Range("Q3") = GreatestDecrease
    Range("Q4") = GreatestTotalVolume
    Range("P2") = TickerGreatestIncrease
    Range("P3") = TickerGreatestDecrease
    Range("P4") = TickerGreatestTotalVolume
    
End Sub

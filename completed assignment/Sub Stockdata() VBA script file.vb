Sub Stockdata()

    ' Set the initial variables
    Dim Ticker As String
    Dim LastRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryTableRow As Integer
    Dim i As Long
    Dim GreatestIncreaseTicker As String
    Dim GreatestIncrease As Double
    Dim GreatestDecreaseTicker As String
    Dim GreatestDecrease As Double
    
    
    
    
    ' Set the headers for the summary table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Set the initial values for the variables
    YearlyChange = 0
    TotalVolume = 0
    SummaryTableRow = 2
    OpeningPrice = Range("C2").Value
    
    ' Get the last row of the data in column A
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through the data
    For i = 2 To LastRow
        
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Set the closing price and add the total volume to the current stock
            ClosingPrice = Cells(i, 6).Value
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            ' Calculate the yearly change and percent change
            YearlyChange = ClosingPrice - OpeningPrice
            If OpeningPrice <> 0 Then
                PercentChange = YearlyChange / OpeningPrice
            Else
                PercentChange = 0
            End If
            
            ' Add the data to the table
            Range("I" & SummaryTableRow).Value = Cells(i, 1).Value
            Range("J" & SummaryTableRow).Value = YearlyChange
            Range("K" & SummaryTableRow).Value = PercentChange
            Range("L" & SummaryTableRow).Value = TotalVolume
            
              ' Update the greatest increase and greatest decrease variables
            If PercentChange > GreatestIncrease Then
                GreatestIncrease = PercentChange
                GreatestIncreaseTicker = Cells(i, 1).Value
            End If
            If PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                GreatestDecreaseTicker = Cells(i, 1).Value
            End If

            ' Format the yearly change cell with conditional formatting
            If YearlyChange > 0 Then
                Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            ElseIf YearlyChange < 0 Then
                Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            End If
            
            ' Reset the initial values for the variables
            YearlyChange = 0
            TotalVolume = 0
            SummaryTableRow = SummaryTableRow + 1
            OpeningPrice = Cells(i + 1, 3).Value
            
        Else
            
            ' If the ticker symbol is the same, add to the total volume
            TotalVolume = TotalVolume + Cells(i, 7).Value
        End If
        
    
    Next i
    
    Range("N2").Value = "Stock with the greatest % increase: "
    Range("N3").Value = "Stock with the greatest % decrease: "
    Range("O2").Value = GreatestIncreaseTicker
    Range("O3").Value = GreatestDecreaseTicker
    Range("P2").Value = Format(GreatestIncrease, "0.00%")
    Range("P3").Value = Format(GreatestDecrease, "0.00%")
    
   End Sub
   
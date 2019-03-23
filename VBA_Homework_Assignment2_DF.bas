Attribute VB_Name = "Module1"
Sub TotalVolume()

    Dim WorksheetName, Stock, BestStock, WorstStock, BestTotalVolumeStock As String
    Dim StockVolume, StockVolumeTotal, OpenPrice, ClosePrice, YearlyChange, PercentChange, BestPercentChange, WorstPercentChange, BestTotalVolume As Double
    Dim SummaryTableRow, BestWorstTableRow As Integer
  
    StockVolumeTotal = 0
    
    For Each ws In Worksheets
    
        ws.Activate
  
		' find last row of main table using column A
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        SummaryTableRow = 2

        For i = 2 To LastRow
    
            StockVolume = Cells(i, 7).Value
            
            'If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                'OpenPrice = Cells(i, 3).Value
               ' StockVolumeTotal = StockVolumeTotal + StockVolume
           ' ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
		   
		   ' Check to see if last row for Stock Ticker
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Stock = Cells(i, 1).Value
                OpenPrice = Cells(i, 3).Value
                StockVolumeTotal = StockVolumeTotal + StockVolume
                ClosePrice = Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
				' If OpenPrice is 0, force PercentChange to be 0 so we don't get divide by 0 error (Stack Overflow 11)
                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice  ' Complete division if denominator is not 0
                End If
                Range("I" & SummaryTableRow).Value = Stock
                Range("J" & SummaryTableRow).Value = StockVolumeTotal
                Range("K" & SummaryTableRow).Value = YearlyChange
				' Format background color based on value being greater than or = to 0
                If Range("K" & SummaryTableRow).Value > 0 Then
                    Range("K" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    Range("K" & SummaryTableRow).Interior.ColorIndex = 3
                End If
                Range("L" & SummaryTableRow).Value = PercentChange
                
                SummaryTableRow = SummaryTableRow + 1
                StockVolumeTotal = 0
            Else
                StockVolumeTotal = StockVolumeTotal + StockVolume
            End If

        Next i
        
		' Find last row in summary table
        LastRowSummary = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        BestWorstTableRow = 2
        
        For j = 2 To LastRowSummary
        
			' If first row of summary table, set each variable equal to value from first row
            If j = 2 Then
                BestStock = Cells(j, 9).Value
                WorstStock = Cells(j, 9).Value
                BestTotalVolumeStock = Cells(j, 9).Value
                BestPercentChange = Cells(j, 12).Value
                WorstPercentChange = Cells(j, 12).Value
                BestTotalVolume = Cells(j, 10).Value
				
			' Compare current WorstPercentageChange to existing value; update if new row is less than previous value
            ElseIf Cells(j, 12).Value < WorstPercentChange Then
                WorstPercentChange = Cells(j, 12).Value
                WorstStock = Cells(j, 9).Value
				
			' Compare current BestPercentChange with new row value and update variable if new value is greater than previous
            ElseIf Cells(j, 12).Value > BestPercentChange Then
                BestPercentChange = Cells(j, 12).Value
                BestStock = Cells(j, 9).Value
            Else
                End If
				
            ' Compare current value for BestTotalVolume and update as needed
            If Cells(j, 10).Value > BestTotalVolume Then
                BestTotalVolume = Cells(j, 10).Value
                BestTotalVolumeStock = Cells(j, 9).Value
            Else
                End If
            
        Next j
        
        Range("O" & BestWorstTableRow).Value = BestStock
        Range("P" & BestWorstTableRow).Value = BestPercentChange
        
        Range("O" & BestWorstTableRow + 1).Value = WorstStock
        Range("P" & BestWorstTableRow + 1).Value = WorstPercentChange
        
        Range("O" & BestWorstTableRow + 2).Value = BestTotalVolumeStock
        Range("P" & BestWorstTableRow + 2).Value = BestTotalVolume
        
    Next ws
    
    MsgBox ("Macro is done!")
    
    
'ERROR_HANDLER:
    'Select Case Err.Number
        'Case 11 'Division by zero
            'PercentChange = 0
            'Err.Clear
            'Resume Next
    
    'End Select
    
End Sub

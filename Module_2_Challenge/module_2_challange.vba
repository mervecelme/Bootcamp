Sub Calculate()

    Dim ws As Worksheet
    Dim tickerName As String
    
    Dim summaryTableRow As Integer
    
    Dim yearlyChange As Double, percentageChange As Double
    Dim openValue As Double, closeValue As Double
    
    Dim totalStock As Double
    Dim lastRow As Long
    Dim lastRowSummary As Integer
    Dim G_Increase_Name As String
    Dim G_Decrease_Name As String
    Dim G_TotalVolume_Name As String
    Dim G_Increase_Value As Double
    Dim G_Decrease_Value As Double
    Dim G_TotalVolume_Value As Double
    Dim MaxRow As Double
    Dim MinRow As Double
    Dim MaxVolume As Double
    
    For Each ws In ThisWorkbook.Worksheets
        
        summaryTableRow = 2
        
        'Set headers
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percentage Change", "Total Stock Volume")
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1:P1").Value = Array("Ticker", "Value")
        
        'Find last row
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'Find first open value for first ticker
        openValue = ws.Cells(2, 3).Value
        
        'Loop to read all data
        For i = 2 To lastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Find ticker name
                tickerName = ws.Cells(i, 1).Value
                
                'Find close value
                closeValue = ws.Cells(i, 6).Value
                
                'Calculate yearly chage
                yearlyChange = closeValue - openValue
                
                'Find percentage change
                If openValue <> 0 Then
                    percentageChange = 1 - (closeValue / openValue)
                Else
                    percentageChange = 0
                End If
                
                'Get new open value for next ticker
                openValue = ws.Cells(i + 1, 3).Value
                
                'Total stock volume
                totalStock = totalStock + ws.Cells(i, 7).Value
                
                'Printing Summary
                ws.Range("I" & summaryTableRow).Value = tickerName
                ws.Range("J" & summaryTableRow).Value = yearlyChange
                ws.Range("K" & summaryTableRow).Value = percentageChange
                ws.Range("L" & summaryTableRow).Value = totalStock
                
                'Change format to percentage
                ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
                
                summaryTableRow = summaryTableRow + 1
                totalStock = 0
                
            Else
            
                'Calculate total stock volume for the same ticker
                totalStock = totalStock + ws.Cells(i, 7).Value
            
            End If
                               
        Next i
        
        lastRowSummary = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        For i = 2 To lastRowSummary
        
            'Coloring'
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 0
            End If
            
        Next i
        
         'Calculate and set values for the summary section
        G_Increase_Value = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Range("P2").Value = G_Increase_Value
        G_Decrease_Value = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Range("P3").Value = G_Decrease_Value
        G_TotalVolume_Value = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Range("P4").Value = G_TotalVolume_Value
        
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
        
        MaxRow = WorksheetFunction.Match((G_Increase_Value), ws.Range("K:K"), 0)
        ws.Range("O2").Value = ws.Cells(MaxRow, 9).Value
    
        MinRow = WorksheetFunction.Match((G_Decrease_Value), ws.Range("K:K"), 0)
        ws.Range("O3").Value = ws.Cells(MinRow, 9).Value
    
        MaxVolume = WorksheetFunction.Match(G_TotalVolume_Value, ws.Range("L:L"), 0)
        ws.Range("O4").Value = ws.Cells(MaxVolume, 9).Value
             
    Next ws

End Sub



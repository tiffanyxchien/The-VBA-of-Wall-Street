Sub VBAofWallStreet()

    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
    
        Dim Ticker As String
    
        Dim StockVolume As Double
        StockVolume = 0
    
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRow
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                Ticker = ws.Cells(i, 1).Value
            
                StockVolume = StockVolume + ws.Cells(i, 7).Value
            
                ws.Range("I" & SummaryTableRow).Value = Ticker
            
                ws.Range("J" & SummaryTableRow).Value = StockVolume
            
                SummaryTableRow = SummaryTableRow + 1
            
                StockVolume = 0
            
            Else
        
                StockVolume = StockVolume + ws.Cells(i, 7).Value
            
            End If

        Next i
    
    Next ws
    
End Sub

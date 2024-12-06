Sub StockAnalysis()
    
    Dim ws As Worksheet
    
    Dim Ticker As String
    
    Dim OpenPrice As Double
    
    Dim ClosePrice As Double
    
    Dim TotalVolume As Double

    Dim LastRow As Long
    
    Dim i As Long

    For Each ws In ThisWorkbook.Worksheets
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        SummaryRow = 2
        
        TotalVolume = 0
        
        ws.Range("I1").Value = "Ticker"
        
        ws.Range("J1").Value = "Quarterly Change"
        
        ws.Range("K1").Value = "Percent Change"
        
        ws.Range("L1").Value = "Total Stock Volume"
        
        Ticker = ws.Cells(2, 1).Value
        
        OpenPrice = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> Ticker Then
            
                ClosePrice = ws.Cells(i, 6).Value
                
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                ws.Cells(SummaryRow, 9).Value = Ticker
                
                ws.Cells(SummaryRow, 10).Value = ClosePrice - OpenPrice
                
                If OpenPrice <> 0 Then
                
                    ws.Cells(SummaryRow, 11).Value = (ClosePrice - OpenPrice) / OpenPrice
            
                Else
                    ws.Cells(SummaryRow, 11).Value = 0
                                                                                              
                End If
        
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                With ws.Cells(SummaryRow, 10)
                
                    If .Value < 0 Then
                        
                        .Interior.Color = RGB(255, 0, 0)
                    
                    ElseIf .Value > 0 Then
                        
                        .Interior.Color = RGB(0, 255, 0)
                    
                    End If
                
                End With
                
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                
                SummaryRow = SummaryRow + 1
                
                TotalVolume = 0
                
                Ticker = ws.Cells(i + 1, 1).Value
                
                OpenPrice = ws.Cells(i + 1, 3).Value
                
            Else
                
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        CreateGreatestTable ws, SummaryRow - 1
        
    Next ws
                
End Sub

Sub CreateGreatestTable(ws As Worksheet, LastSummaryRow As Long)

    Dim GreatestIncrease As Double
    
    Dim GreatestDecrease As Double
    
    Dim GreatestVolume As Double
    
    Dim i As Long
    
    Dim IncreaseTicker As String
    
    Dim DecreaseTicker As String
    
    Dim VolumeTicker As String
    
    GreatestIncrease = ws.Cells(2, 11).Value
    
    GreatestDecrease = ws.Cells(2, 11).Value
    
    GreatestVolume = ws.Cells(2, 12).Value
    
    For i = 2 To LastSummaryRow
    
        If ws.Cells(i, 11).Value > GreatestIncrease Then
            
                GreatestIncrease = ws.Cells(i, 11).Value
                
                IncreaseTicker = ws.Cells(i, 9).Value
                
        End If
        
        If ws.Cells(i, 11).Value < GreatestDecrease Then
            
                GreatestDecrease = ws.Cells(i, 11).Value
                
                DecreaseTicker = ws.Cells(i, 9).Value
                
        End If
        
        If ws.Cells(i, 12).Value > GreatestVolume Then
        
                GreatestVolume = ws.Cells(i, 12).Value
                
                VolumeTicker = ws.Cells(i, 9).Value
                
        End If
        
    Next i
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
    ws.Cells(1, 16).Value = "Ticker"
    
    ws.Cells(1, 17).Value = "Value"
    
    
    ws.Cells(2, 16).Value = IncreaseTicker
    
    ws.Cells(3, 16).Value = DecreaseTicker
    
    ws.Cells(4, 16).Value = VolumeTicker
    
    
    ws.Cells(2, 17).Value = GreatestIncrease
    
    ws.Cells(3, 17).Value = GreatestDecrease
    
    ws.Cells(4, 17).Value = GreatestVolume
    
    
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    
 End Sub

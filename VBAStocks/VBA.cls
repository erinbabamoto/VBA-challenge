VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks():
    
    'Iterate over each worksheet
    For Each ws In Worksheets
    
        ' Set dimensions
        Dim Ticker  As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        
        ' Summary table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Ticker = ws.Cells(2, 1).Value
        
        TotalStockVolume = 0
        SummaryTableRow = 2
        StartRow = 2
        
        ' Row count
        RowCount = ws.Cells(1, 1).End(xlDown).Row
        
        ' Loop
        For Row = 2 To RowCount
            
            TotalStockVolume = TotalStockVolume + ws.Cells(Row, 7).Value
            
            If ws.Cells(Row, 1) <> ws.Cells(Row + 1, 1) Then
                
                ' TICKER SYMBOL
                ws.Cells(SummaryTableRow, 9).Value = Ticker
                
                ' YEARLY CHANGE
                OpenPrice = ws.Cells(StartRow, 3).Value
                ClosePrice = ws.Cells(Row, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                ws.Cells(SummaryTableRow, 10).Value = YearlyChange
                
                If ws.Cells(SummaryTableRow, 10).Value > 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.Color = RGB(0, 255, 0)
                    
                ElseIf ws.Cells(SummaryTableRow, 10).Value < 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.Color = RGB(255, 0, 0)
                    
                End If
                
                ' PERCENT CHANGE
                If OpenPrice <> 0 Then
                    PercentChange = YearlyChange / OpenPrice
                    
                Else
                    PercentChange = 0
                    
                End If
                
                ws.Cells(SummaryTableRow, 11).Value = PercentChange
                ws.Cells(SummaryTableRow, 11) = Format(PercentChange, "Percent")
                
                ' TOTAL STOCK VOLUME
                ws.Cells(SummaryTableRow, 12).Value = TotalStockVolume
                TotalStockVolume = 0
                
                ' Set up for next ticker calculation
                SummaryTableRow = SummaryTableRow + 1
                StartRow = Row + 1
                
                Ticker = ws.Cells(Row + 1, 1).Value
                
            End If
            
        Next Row
        
    Next ws
    
End Sub


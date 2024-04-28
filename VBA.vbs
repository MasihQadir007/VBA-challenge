Sub StockAnalysis()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize summary table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Initialize variables
        SummaryRow = 2
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Ticker = ws.Cells(2, 1).Value
        OpeningPrice = ws.Cells(2, 3).Value
        TotalVolume = 0
        
        ' Loop through each row in the worksheet
        For i = 2 To LastRow
            ' Check if Ticker symbol has changed
            If ws.Cells(i, 1).Value <> Ticker Then
                ' Calculate Quarterly Change and Percent Change
                ClosingPrice = ws.Cells(i - 1, 6).Value
                QuarterlyChange = ClosingPrice - OpeningPrice
                PercentChange = (ClosingPrice - OpeningPrice) / OpeningPrice
                
                ' Output results to summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = QuarterlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                ' Color code Percent Change
                If PercentChange >= 0 Then
                    ws.Cells(SummaryRow, 11).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(SummaryRow, 11).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                ' Format Quarterly Change
                If QuarterlyChange >= 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                ' Reset variables for next Ticker symbol
                SummaryRow = SummaryRow + 1
                Ticker = ws.Cells(i, 1).Value
                OpeningPrice = ws.Cells(i, 3).Value
                TotalVolume = 0
            End If
            
            ' Accumulate Total Volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        Next i
        
        ' Summarize the last Ticker symbol
        ClosingPrice = ws.Cells(LastRow, 6).Value
        QuarterlyChange = ClosingPrice - OpeningPrice
        PercentChange = (ClosingPrice - OpeningPrice) / OpeningPrice
        ws.Cells(SummaryRow, 9).Value = Ticker
        ws.Cells(SummaryRow, 10).Value = QuarterlyChange
        ws.Cells(SummaryRow, 11).Value = PercentChange
        ws.Cells(SummaryRow, 12).Value = TotalVolume
        
        ' Format Quarterly Change for the last row
        If QuarterlyChange >= 0 Then
            ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
        Else
            ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
        End If
        
        ' Format Percent Change for the last row
        If PercentChange >= 0 Then
            ws.Cells(SummaryRow, 11).Interior.Color = RGB(0, 255, 0) ' Green
        Else
            ws.Cells(SummaryRow, 11).Interior.Color = RGB(255, 0, 0) ' Red
        End If
    Next ws
End Sub

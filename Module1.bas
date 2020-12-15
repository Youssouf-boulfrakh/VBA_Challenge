Attribute VB_Name = "Module1"
Sub Apphabetical_testing()

    Dim worsheetname As String
    Dim LastRow As Long
    Dim SummaryRow As Long
    Dim t_volume As Double
    Dim op_price As Double
    Dim cl_price As Double
    Dim pc_change As Double


    For Each ws In Worksheets
        worksheetname = ws.Name
        SummaryRow = 2
        t_volume = 0
        op_price = ws.Cells(2, 3).Value
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.[I1].Value = "Ticker"
        ws.[J1].Value = "Yearly change"
        ws.[K1].Value = "Yearly percentage"
        ws.[L1].Value = "Total volume"
        ws.[P1].Value = "Ticker"
        ws.[Q1].Value = "Yearly percentage"
        ws.[R1].Value = "Yearly change"
        ws.[O2].Value = "Max Increase"
        ws.[O3].Value = "Min Increase"
        ws.[O4].Value = "Max volume"
        For i = 2 To LastRow
            t_volume = t_volume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(SummaryRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(SummaryRow, 12).Value = t_volume
                yr_change = ws.Cells(i, 6).Value - op_price
                ws.Cells(SummaryRow, 10).Value = yr_change
                
                If yr_change > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                End If
                
                If op_price > 0 Then
                    pc_change = (yr_change / op_price) * 100
                    ws.Cells(SummaryRow, 11).Value = pc_change
                Else
                    pc_change = 0
                    ws.Cells(SummaryRow, 11).Value = pc_change
                End If
                        
                If pc_change > max_inc Then
                    max_inc = pc_change
                    max_ticker = ws.Cells(i, 1).Value
                    ws.Cells(2, 16).Value = max_ticker
                    ws.Cells(2, 17).Value = max_inc
                    ws.Cells(2, 18).Value = yr_change
                End If
    
                If pc_change < min_inc Then
                    min_inc = pc_change
                    min_ticker = ws.Cells(i, 1).Value
                    ws.Cells(3, 16).Value = min_ticker
                    ws.Cells(3, 17).Value = min_inc
                    ws.Cells(3, 18).Value = yr_change
                End If
                
                If t_volume > max2 Then
                    max2 = t_volume
                    v_ticker = ws.Cells(i, 1).Value
                    ws.Cells(4, 16).Value = v_ticker
                    ws.Cells(5, 15).Value = max2
                End If
                SummaryRow = SummaryRow + 1
                op_price = ws.Cells(i + 1, 3).Value
                t_volume = 0
                            
            End If
        Next i
    Next ws
End Sub



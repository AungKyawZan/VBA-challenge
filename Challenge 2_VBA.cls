Sub stockdata()
    Dim ws As Worksheet
    Dim ticker As String
    Dim yearly_change As Double
    Dim volume As Double
    Dim result As Integer
    Dim p As Integer
    Dim lastRow As Long
    
    Dim max_increase As Double, max_decrease As Double, max_volume As Double
    Dim max_inc_ticker As String, max_dec_ticker As String, max_vol_ticker As String

    For Each ws In Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        volume = 0
        result = 2
        p = 2
        
        For o = 2 To lastRow
            If ws.Cells(o + 1, 1).Value <> ws.Cells(o, 1).Value Then
                ticker = ws.Cells(o, 1).Value
                ws.Cells(result, 9).Value = ticker
                
                ws.Cells(result, 10).Value = ws.Cells(o, 6).Value - ws.Cells(p, 3).Value
                
                If ws.Cells(result, 10).Value < 0 Then
                    ws.Cells(result, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(result, 10).Interior.ColorIndex = 4
                End If
                
                If ws.Cells(p, 3).Value <> 0 Then
                    ws.Cells(result, 11).Value = (ws.Cells(o, 6).Value - ws.Cells(p, 3).Value) / ws.Cells(p, 3).Value
                Else
                    ws.Cells(result, 11).Value = 0
                End If
                
                ws.Cells(result, 12).Value = volume + ws.Cells(o, 7).Value
                
                p = o + 1
                volume = 0
                result = result + 1
            Else
                volume = volume + ws.Cells(o, 7).Value
            End If
        Next o

        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
        max_increase = -1
        max_decrease = 1
        max_volume = 0
        
        For q = 2 To result - 1
            If ws.Cells(q, 11).Value > max_increase Then
                max_increase = ws.Cells(q, 11).Value
                max_inc_ticker = ws.Cells(q, 9).Value
            End If
            
            If ws.Cells(q, 11).Value < max_decrease Then
                max_decrease = ws.Cells(q, 11).Value
                max_dec_ticker = ws.Cells(q, 9).Value
            End If
            
            If ws.Cells(q, 12).Value > max_volume Then
                max_volume = ws.Cells(q, 12).Value
                max_vol_ticker = ws.Cells(q, 9).Value
            End If
        Next q

        ws.Cells(2, 16).Value = max_inc_ticker
        ws.Cells(2, 17).Value = max_increase
        ws.Cells(3, 16).Value = max_dec_ticker
        ws.Cells(3, 17).Value = max_decrease
        ws.Cells(4, 16).Value = max_vol_ticker
        ws.Cells(4, 17).Value = max_volume
    Next ws
End Sub

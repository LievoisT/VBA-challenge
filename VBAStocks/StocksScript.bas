Attribute VB_Name = "Module1"


Sub vbastocks():
    
    
    Dim ws As Worksheet
    Dim tick As String
    Dim open_price As Double
    Dim close_price As Double
    Dim percent_change As Double
    Dim stock_vol As LongLong
    Dim great_inc As Double
    Dim great_dec As Double
    Dim great_vol As LongLong
    Dim i As Long
    Dim n As Integer
    
    
    For Each ws In ThisWorkbook.Worksheets:

        tick = ws.Cells(2, 1).Value
        open_price = ws.Cells(2, 3).Value
        stock_vol = 0
        n = 2
        great_inc = 0
        great_dec = 0
        great_vol = 0
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
        
        
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row:
        
            stock_vol = stock_vol + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> tick Then
            
                close_price = ws.Cells(i, 6).Value
                ws.Cells(n, 9).Value = tick
                ws.Cells(n, 10).Value = close_price - open_price
                ws.Cells(n, 12).Value = stock_vol
                
                If open_price <> 0 Then
                
                    percent_change = (close_price - open_price) / open_price
                    ws.Cells(n, 11).Value = percent_change
                    
                    If percent_change > great_inc Then
                    
                        great_inc = percent_change
                        ws.Cells(2, 16).Value = tick
                        ws.Cells(2, 17).Value = percent_change
                        
                    ElseIf percent_change < great_dec Then
                    
                        great_dec = percent_change
                        ws.Cells(3, 16).Value = tick
                        ws.Cells(3, 17).Value = percent_change
                        
                    End If
                    
                End If

                If stock_vol > great_vol Then
                    ws.Cells(4, 16).Value = tick
                    ws.Cells(4, 17).Value = stock_vol
                    great_vol = stock_vol
                End If
                
                If ws.Cells(n, 10).Value > 0 Then
                    ws.Cells(n, 10).Interior.Color = vbGreen
                ElseIf ws.Cells(n, 10).Value <= 0 Then
                    ws.Cells(n, 10).Interior.Color = vbRed
                
                End If
                
                ws.Cells(n, 11).NumberFormat = "0.00%"
                
                tick = ws.Cells(i + 1, 1).Value
                open_price = ws.Cells(i + 1, 3).Value
                stock_vol = 0
                n = n + 1
                
            End If
                
        Next i
          
    Next ws
     
End Sub


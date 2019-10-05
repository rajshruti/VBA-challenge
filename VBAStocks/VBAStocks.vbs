Sub VBAStocks()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        'Variables declaration
        
        Dim ticker As String
        Dim total_stockvol, greatest_totalvol, i, j, open_index, LastRow, LastColumn As Integer
        Dim open_price, close_price, yearly_change, percent_change, greatest_increase, greatest_decrease As Double
        
        ' Initialize the variables
        
        greatest_totalvol = 0
        greatest_increase = 0
        greatest_decrease = 0
        total_stockvol = 0
        open_index = 2
        j = 2
           
        ' Determine the Last Row
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
        ' Print the labels
 
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Value"
        ws.Range("O1").Value = "Ticker"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        For i = 2 To LastRow
                    
            ' Compare current ticker value with the next one
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                open_price = ws.Cells(open_index, 3).Value
                close_price = ws.Cells(i, 6)
                yearly_change = close_price - open_price
                
                ' Error handling: Prints percent_change value as 0 if the open_price=0
                
                If open_price = 0 Then
                
                    percent_change = 0
                    
                Else
                
                    percent_change = Round((yearly_change / open_price) * 100, 2)
                    
                End If
               
                ticker = ws.Cells(i, 1).Value
                total_stockvol = total_stockvol + ws.Cells(i, 7).Value
                
                ' Calculate greatest total stock volume
                    
                If total_stockvol > greatest_totalvol Then
                        
                    greatest_totalvol = total_stockvol
                    ws.Range("P4").Value = greatest_totalvol
                    ws.Range("O4").Value = ticker
                    
                End If
                                        
                ' Calculate greatest %Increase
                                        
                If percent_change > greatest_increase Then
                        
                    greatest_increase = percent_change
                    ws.Range("P2").Value = "%" & greatest_increase
                    ws.Range("O2").Value = ticker
                    
                End If
                        
                ' Calculate greatest %Decrease
                        
                If percent_change < greatest_decrease Then
                    
                    greatest_decrease = percent_change
                    ws.Range("P3").Value = "%" & greatest_decrease
                    ws.Range("O3").Value = ticker
                    
                End If
                    
                ' print the results
                
                ws.Range("I" & j).Value = ticker
                ws.Range("J" & j).Value = Round(yearly_change, 2)
                
                ' Conditional formatting to highlight positive change in green and negative change in red
                
                If yearly_change > 0 Then
                
                    ws.Range("J" & j).Interior.ColorIndex = 4
                    
                ElseIf yearly_change < 0 Then
                
                    ws.Range("J" & j).Interior.ColorIndex = 3
                    
                End If
                
                ws.Range("K" & j).Value = "%" & percent_change
                ws.Range("L" & j).Value = total_stockvol
                
                ' Reset the stock_vol to zero for next ticker iteration
                
                total_stockvol = 0
                
                ' Increment the value of j and open_index value by 1
                
                open_index = i + 1
                
                j = j + 1
                
            Else
            
                total_stockvol = total_stockvol + ws.Cells(i, 7).Value
                
            End If
    
        Next i
        
    Next ws
        
End Sub
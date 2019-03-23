Sub stockExchange()
Dim Annual_Difference As Double
Dim tickerSymbol As String
Dim resultRow As Integer
Dim percentChange As Double

resultRow = 2
Annual_Difference = 0
totalVolume = 0
'percentChange = 0

Dim j As Long
    j = 2


lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        'Write conditional for differnce in ticker symbol
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'define and print ticker symbol in column 9
            tickerSymbol = Cells(i, 1).Value
            
            
            'sum new ticker totals and print it in column 12
            totalVolume = totalVolume + Cells(i, 7).Value
            
            
            Dim open_value As Double
            open_value = Cells(j, 3).Value
            Dim closing_value As Double
            closing_value = Cells(i, 6).Value
            
            Annual_Difference = closing_value - open_value
            
            If open_value <> 0 Then
                percentChange = Annual_Difference / open_value
            Else
                percentChange = 0
           End If
            
            'Round percent change
            'percentChange = Round(percentChange, 4)
            'print values
            Cells(resultRow, 10).Value = Annual_Difference
            Cells(resultRow, 11).Value = percentChange
            Cells(resultRow, 9).Value = (Cells(i, 1).Value)
            Cells(resultRow, 12).Value = totalVolume
            
           
           
            'indicate print position for new values as we move through the loop
            resultRow = resultRow + 1
            j = i + 1
            
            'reset values
            totalVolume = 0
            Annual_Difference = 0
            
            
            
            
                   
            
    
        Else
            'sum individual totals and print it in column 12 for each identical ticker symbol
                totalVolume = totalVolume + (Cells(i, 7).Value)

        End If
        'conditional for changing cell color based on positive or negative value
               If Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.Color = RGB(200, 0, 0)
               End If
               If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.Color = RGB(0, 128, 0)
               Else
               End If
    
    Next i
    
    

End Sub
Attribute VB_Name = "Module1"
    Sub Run()
          
        'Set Variables
        For Each ws In Worksheets
        Dim trow As Long
        Dim lastrow As Long
        
        Dim startprice As Double
        Dim endprice As Double
        Dim yearchange As Double
        Dim pctchange As Double
        Dim count As Integer
        Dim Volume As Double
        
        Dim count2 As Double
        Dim max As Double
        Dim maxTicker As String
        Dim min As Double
        Dim minTicker As String
        Dim maxvolume As Double
        Dim MVTicker As String
        Dim lastrow2 As Long
        
        lastrow = Cells(Rows.count, 1).End(xlUp).Row
        count = 2
        count2 = 2
        
        'Create Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
      
        
        'Extract Data & Create Table
        For trow = 2 To lastrow
        
            If ws.Cells(trow, 1).Value <> ws.Cells(trow - 1, 1).Value Then
                ws.Cells(count, 9).Value = ws.Cells(trow, 1).Value
                startprice = ws.Cells(trow, 3).Value
                
            ElseIf ws.Cells(trow, 1).Value <> ws.Cells(trow + 1, 1).Value Then
                endprice = ws.Cells(trow, 6).Value
                yearchange = (endprice - startprice)
                ws.Cells(count, 10).Value = yearchange
                ws.Cells(count, 11).Value = yearchange / startprice
                count = count + 1
            End If
            
            If ws.Cells(trow, 1).Value = ws.Cells(trow + 1, 1).Value Then
                Volume = Volume + ws.Cells(trow, 7).Value
   
            Else
                ws.Cells(count2, 12).Value = Volume
                Volume = 0
                count2 = count2 + 1
            End If
            
        Next trow
         
    max = 0
    min = 0
    lastrow2 = Cells(Rows.count, 9).End(xlUp).Row
 
 
    For i = 2 To lastrow2

        If ws.Cells(i, 11).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            ws.Cells(i, 11).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 11).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            ws.Cells(i, 11).Interior.ColorIndex = 3
        End If
    
        If ws.Cells(i, 11).Value > max Then
            max = ws.Cells(i, 11).Value
            maxTicker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 11).Value < min Then
            min = ws.Cells(i, 11).Value
            minTicker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 12).Value > maxvolume Then
            maxvolume = ws.Cells(i, 12).Value
            MVTicker = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    ws.Cells(2, 16).Value = maxTicker
    ws.Cells(2, 17).Value = max
    ws.Cells(3, 16).Value = minTicker
    ws.Cells(3, 17).Value = min
    ws.Cells(4, 16).Value = MVTicker
    ws.Cells(4, 17).Value = maxvolume
        
    Next ws
    
End Sub
    

    


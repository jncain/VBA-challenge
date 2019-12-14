Attribute VB_Name = "Module1"
Sub stocksummary():
    'loop through all sheets
    For Each ws In Worksheets
    
        'determining the last row in the worksheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'creating the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        'setting variables to hold values
        Dim tickername As String
        Dim begprice As Double
        Dim endprice As Double
        Dim yearchange As Double
        Dim vol As Double
        
        'setting volume sum to 0
        vol = 0
        
        'keep track of stock through the year
        Dim stockcount As Integer
        stockcount = 0
    
        'keep track of summary table rows
        Dim table_row As Integer
        table_row = 2
    
        'create summary table
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                tickername = ws.Cells(i, 1).Value
                endprice = ws.Cells(i, 6).Value
                vol = vol + ws.Cells(i, 7).Value

                'find the beginning year price
                begprice = ws.Cells(i - stockcount, 3).Value
                
                'caluclate the yearly change amount
                yearchange = endprice - begprice
                
                'calcuate the percent yearly change
                'if statement to fix issue with division by 0
                If begprice = 0 Then
                    perchange = 0
                Else: perchange = yearchange / begprice
                End If
            
                'print summary table
                ws.Range("I" & table_row).Value = tickername
                ws.Range("J" & table_row).Value = yearchange
                ws.Range("K" & table_row).Value = perchange
                ws.Range("L" & table_row).Value = vol
                
                'format percent change to a percent
                ws.Range("K" & table_row).NumberFormat = "0.00%"
            
                'conditional formating for percent column
                If perchange < 0 Then
                    ws.Range("K" & table_row).Interior.ColorIndex = 3
                Else: ws.Range("K" & table_row).Interior.ColorIndex = 4
                End If
                                                
                'loop maintenance
                table_row = table_row + 1
                vol = 0
                stockcount = 0
            
            Else:
                vol = vol + ws.Cells(i, 7).Value
                stockcount = stockcount + 1
            
            End If
        Next i
        
        'create second summary table - greatest increase, decrease, volume
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        'determing last row of the fist summary table
        summary_lastrow = table_row - 1
        
        'Greatest increase and maximum value (set max to 0)
        Dim increase_ticker As String
        Dim max As Double
        max = 0
        
        'Greatest decrease and miniumum value (set min to 0)
        Dim decrease_ticker As String
        Dim min As Double
        min = 0
        
        'Greatest total volume and volume total (set volume to 0)
        Dim volume_ticker As String
        Dim totalvolume As Double
        totalvolume = 0
        
        'loop through summary table to find stocks with the values
        For i = 2 To summary_lastrow
            'find stock with greatest percent increase
            If ws.Cells(i, 11).Value > max Then
                max = ws.Cells(i, 11).Value
                increase_ticker = ws.Cells(i, 9).Value
            End If
            'find stock with greatest percent decrease
            If ws.Cells(i, 11).Value < min Then
                min = ws.Cells(i, 11).Value
                decrease_ticker = ws.Cells(i, 9).Value
            End If
            'find stock with greatest total volume
            If ws.Cells(i, 12).Value > totalvolume Then
                totalvolume = ws.Cells(i, 12).Value
                volume_ticker = ws.Cells(i, 9).Value
            End If
            
        Next i
        
        'print values to second summary table
        ws.Range("O2").Value = increase_ticker
        ws.Range("P2").Value = max
        
        ws.Range("O3").Value = decrease_ticker
        ws.Range("P3").Value = min
        'format the greatest increase and decrease to percents
        ws.Range("P2, P3").NumberFormat = "0.00%"
        
        ws.Range("O4").Value = volume_ticker
        ws.Range("P4").Value = totalvolume
        
    Next ws
    
    MsgBox ("Summarization of Yearly Stocks Complete")
    
End Sub

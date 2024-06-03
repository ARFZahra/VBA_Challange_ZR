Sub challenge()

Dim ws As Worksheet
Dim lastRow As Long

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets

    ' Set an initial variable for ...
    Dim ticker As String
    Dim Total_Ticker_Volume As Double
    Dim Percentage_change As Double
    Dim startprice As Double, endprice As Double
 
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
    ' Insert headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly_change"
    ws.Cells(1, 11).Value = "Percentage_change"
    ws.Cells(1, 12).Value = "Total_Ticker_Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Volume"
     
    
    ' Set an initial variable for total volume of ticker
    Dim Total_Ticker As Double
    Total_Ticker = 0
    
    ' Keep track of the location for each the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
      
    ' Loop through all tickers
    For i = 2 To lastRow
    
        ' Add to the ticker Total
        Total_Ticker = Total_Ticker + ws.Cells(i, 7).Value
        
        ' Check if we are still within the same ticker, if we are not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        
            ' Set the ticker name
            ticker = ws.Cells(i, 1).Value
            
            ' Determine start price
            If startprice = 0 Then
                startprice = ws.Cells(i, 3).Value
            End If
            
            ' Determine end price
            endprice = ws.Cells(i, 6).Value
            
            ' Calculate the price change
            ws.Cells(Summary_Table_Row, 10).Value = endprice - startprice
            
            ' Calculate the percentage price change
            If startprice <> 0 Then
                Percentage_change = ((endprice - startprice) / startprice)
            Else
                Percentage_change = 0
            End If

            
            ' Print the ticker name in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = ticker
            
            ' Print the total ticker to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Total_Ticker
            
            ' Print the price change to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = endprice - startprice
            
            ' Print the percentage price to the Summary Table
            ws.Range("K" & Summary_Table_Row).Value = Percentage_change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
             ' Check for greatest increase, decrease, and volume
            If Percentage_change > maxIncrease Then
                maxIncrease = Percentage_change
                maxIncreaseTicker = ticker
                maxIncreaseVolume = Percentage_change
            End If
            
            If Percentage_change < maxDecrease Then
                maxDecrease = Percentage_change
                maxDecreaseTicker = ticker
                maxDecreaseVolume = Percentage_change
            End If
            
            If Total_Ticker > maxVolume Then
                maxVolume = Total_Ticker
                maxVolumeTicker = ticker
                maxVolumeTicker = ticker
            End If
            
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the ticker Total
            Total_Ticker = 0
            startprice = 0
            endprice = 0
        
        Else
            ' Set the starting price for the next ticker
            If startprice = 0 Then
                startprice = ws.Cells(i, 3).Value
            End If
        End If
        
    Next i
    
     ' Print the stocks with greatest increase, decrease, and volume
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(2, 16).Value = maxIncreaseTicker
    ws.Cells(3, 16).Value = maxDecreaseTicker
    ws.Cells(4, 16).Value = maxVolumeTicker
    ws.Cells(2, 17).Value = maxIncreaseVolume
    ws.Cells(3, 17).Value = maxDecreaseVolume
    ws.Cells(4, 17).Value = maxVolume
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    
     ' Apply conditional formatting to highlight positive in green and negative in red
    Dim formatRange As Range
    Set formatRange = ws.Range("J2:J" & Summary_Table_Row - 1)

' Clear existing conditional formatting
    formatRange.FormatConditions.Delete

' Add new conditional formatting for positive values
    Dim cond As FormatCondition
    Set cond = formatRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
cond.Interior.ColorIndex = 4 ' Light green for positive values

' Add new conditional formatting for negative values
    Set cond = formatRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    cond.Interior.ColorIndex = 3 ' Light red for negative values
    
'Alignment to display data
    ws.Columns("A:Q").HorizontalAlignment = xlRight
    ws.Columns("A:Q").AutoFit

    Next ws
    
End Sub


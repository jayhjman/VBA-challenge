Attribute VB_Name = "Module1"
Sub processStocks():


    Dim ws As Worksheet
    
    Dim tickerCol As Long: tickerCol = 1
    
    Dim firstRow As Long: firstRow = 2
    Dim lastRow As Long
    
    Dim currentSymbol As String
    Dim nextSymbol As String
    
    Dim openCol As Long: openCol = 3
    Dim firstOpen As Boolean: firstOpen = True
    Dim firstOpenPrice As Double
    
    Dim closeCol As Long: closeCol = 6
    Dim lastClosePrice As Double
    
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    Dim stockVolumeCol As Long: stockVolumeCol = 7
    Dim totalStockVolume As Double
    
    Dim summaryStartRow As Long: summaryStartRow = firstRow
    Dim symbolCol As Long: symbolCol = 9
    Dim yearlyChangeCol As Long: yearlyChangeCol = 10
    Dim percentChangeCol As Long: percentChangeCol = 11
    Dim totalStockVolumeCol As Long: totalStockVolumeCol = 12
    
    
    ' Loop through each of the worksheets
    For Each ws In Sheets
          
        ' Get the last row in the worksheet
        lastRow = ws.Cells(Rows.Count, tickerCol).End(xlUp).Row
                
        ' Initialize variables
        currentSymbol = ""
        nextSymbol = ""
        totalStockVolume = 0
        
        ' Summary Table Column headers
        summaryStartRow = firstRow
        ws.Cells(summaryStartRow - 1, symbolCol).Value = "Ticker"
        ws.Cells(summaryStartRow - 1, yearlyChangeCol).Value = "Yearly Change"
        ws.Cells(summaryStartRow - 1, percentChangeCol).Value = "Percent Change"
        ws.Cells(summaryStartRow - 1, totalStockVolumeCol).Value = "Total Stock Volume"
        
        
        ' Loop through each of the rows processing them
        For i = firstRow To lastRow
        
            ' Grab the opening price on first day
            If firstOpen Then
                firstOpenPrice = ws.Cells(i, openCol).Value
                firstOpen = False
            End If
        
            ' Grab current and next stock symbol from sheet
            currentSymbol = ws.Cells(i, tickerCol).Value
            nextSymbol = ws.Cells(i + 1, tickerCol).Value
            
            ' Keep a rolling total of current stock sybmols volume
            totalStockVolume = totalStockVolume + ws.Cells(i, stockVolumeCol).Value
            
            ' Detect the stock symbol change
            If (currentSymbol <> nextSymbol) Then
            
                ' Grab the closing price on last day
                lastClosePrice = ws.Cells(i, closeCol).Value
                
                ' Calculate the change
                yearlyChange = lastClosePrice - firstOpenPrice
                percentChange = (lastClosePrice / firstOpenPrice) - 1#
                
                ' Write summary cells
                ws.Cells(summaryStartRow, symbolCol).Value = currentSymbol
                ws.Cells(summaryStartRow, yearlyChangeCol).Value = yearlyChange
                ws.Cells(summaryStartRow, yearlyChangeCol).Interior.Color = vbGreen
                If yearlyChange < 0 Then
                    ws.Cells(summaryStartRow, yearlyChangeCol).Interior.Color = vbRed
                End If
                ws.Cells(summaryStartRow, percentChangeCol).Value = percentChange
                ws.Cells(summaryStartRow, percentChangeCol).NumberFormat = "0.00%"
                ws.Cells(summaryStartRow, totalStockVolumeCol).Value = totalStockVolume
                
                ' Change to the next row for the summary table
                summaryStartRow = summaryStartRow + 1
                
                ' Processing new symbol reset firstOpen flag
                firstOpen = True
                
                ' Reset total stock volume for next stock symbol
                totalStockVolume = 0
                
            End If
            
        Next i
        
        Call processGreatestSummary(ws, firstRow, percentChangeCol, totalStockVolumeCol)
        
    Next ws
    
End Sub

Private Sub processGreatestSummary(ws As Worksheet, startRow As Long, percentCol As Long, _
    volumeCol As Long):
    
    Dim lastRow As Long
    
    Dim labelCol As Long: labelCol = 15
    Dim tickerCol As Long: tickerCol = 16
    Dim valueCol As Long: valueCol = 17
    
    Dim greatestIncreaseRow As Long
    Dim greatestIncrease As Double
    
    Dim greatestDecreaseRow As Long
    Dim greatestDecrease As Double
    
    Dim greatestVolumeRow As Long
    Dim greatestVolume As Double
    
    ' Setup the headers for the greatest summary table
    ws.Cells(startRow - 1, tickerCol).Value = "Ticker"
    ws.Cells(startRow - 1, valueCol).Value = "Value"
    
    ws.Cells(startRow, labelCol).Value = "Greatest % Increase"
    ws.Cells(startRow + 1, labelCol).Value = "Greatest % Decrease"
    ws.Cells(startRow + 2, labelCol).Value = "Greatest Total Volume"
    
    ' Percent columns row count exact same as volume so we can levarage this for both
    lastRow = ws.Cells(Rows.Count, percentCol).End(xlUp).Row
    
    ' Initialize the variables
    greatestIncrease = ws.Cells(startRow, percentCol).Value
    greatestIncreaseRow = startRow
    
    greatestDecrease = ws.Cells(startRow, percentCol).Value
    greatestDecreaseRow = startRow
    
    greatestVolume = ws.Cells(startRow, volumeCol).Value
    greatestVolumeRow = startRow
    
    ' Loop finding greatest and least values
    For i = startRow To lastRow
        ' Greatest percent
        percentVal = ws.Cells(i, percentCol).Value
        If (ws.Cells(i, percentCol).Value > greatestIncrease) Then
            greatestIncreaseRow = i
            greatestIncrease = ws.Cells(i, percentCol).Value
        End If
        ' Least percent
        If (ws.Cells(i, percentCol).Value < greatestDecrease) Then
            greatestDecreaseRow = i
            greatestDecrease = ws.Cells(i, percentCol).Value
        End If
        ' Greatest volume
        If (ws.Cells(i, volumeCol).Value > greatestVolume) Then
            greatestVolumeRow = i
            greatestVolume = ws.Cells(i, volumeCol).Value
        End If
    Next i
    
    ' Print results
    MsgBox ("p > row " & greatestIncreaseRow)
    MsgBox ("p < row " & greatestDecreaseRow)
    MsgBox ("v > row " & greatestVolumeRow)

    
End Sub


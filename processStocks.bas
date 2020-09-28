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
    
    
    ' Loop through each of the worksheets
    For Each ws In Sheets
          
        ' Get the last row in the worksheet
        lastRow = ws.Cells(Rows.Count, tickerCol).End(xlUp).Row
                
        ' Initialize symbols
        currentSymbol = ""
        nextSymbol = ""
        
        
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
            
            ' Detect the stock symbol change
            If (currentSymbol <> nextSymbol) Then
            
                ' Grab the closing price on last day
                lastClosePrice = ws.Cells(i, closeCol).Value
                
                
                yearlyChange = lastClosePrice - firstOpenPrice
                
                percentChange = (lastClosePrice / firstOpenPrice) - 1#
                
                
                ' Print Results
                MsgBox ("Current Symbol: " & currentSymbol & _
                    ", Next Symbol: " & nextSymbol)
                MsgBox ("First Open Price: " & firstOpenPrice & _
                    ", lastClosePrice: " & lastClosePrice)
                MsgBox ("yearlyChange : " & yearlyChange & _
                    ", percentChange: " & percentChange)
                ' Processing new symbol reset firstOpen flag
                firstOpen = True
                
            End If
            
        Next i
        
    Next ws
    
End Sub

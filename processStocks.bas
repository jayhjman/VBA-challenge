Attribute VB_Name = "Module1"
Sub processStocks():


    Dim ws As Worksheet
    
    Dim tickerCol As Long: tickerCol = 1
    
    Dim firstRow As Long: firstRow = 2
    Dim lastRow As Long
    
    Dim currentSymbol As String
    Dim nextSymbol As String
    
    
    'loop through each of the worksheets
    For Each ws In Sheets
          
        'Get the last row in the worksheet
        lastRow = ws.Cells(Rows.Count, tickerCol).End(xlUp).Row
        
        'loop through each of the rows processing them
        For i = firstRow To lastRow
        
            'Grab current and next stock symbol from sheet
            currentSymbol = ws.Cells(i, tickerCol).Value
            nextSymbol = ws.Cells(i + 1, tickerCol).Value
            
            'detect the stock symbol change
            If (currentSymbol <> nextSymbol) Then
                MsgBox ("Current Symbol: " & currentSymbol & _
                    ", Next Symbol: " & nextSymbol)
            End If
            
        Next i
        
        currentSymbol = ""
        nextSymbol = ""
        
    Next ws
    
End Sub

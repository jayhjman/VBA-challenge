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
    
    ' Loop through each of the worksheets
    For Each ws In Sheets
          
        ' Get the last row in the worksheet
        lastRow = ws.Cells(Rows.Count, tickerCol).End(xlUp).Row
                
        ' Initialize symbols
        currentSymbol = ""
        nextSymbol = ""
        
        
        ' Loop through each of the rows processing them
        For i = firstRow To lastRow
        
            If firstOpen Then
                firstOpenPrice = ws.Cells(i, openCol).Value
                firstOpen = False
            End If
        
            ' Grab current and next stock symbol from sheet
            currentSymbol = ws.Cells(i, tickerCol).Value
            nextSymbol = ws.Cells(i + 1, tickerCol).Value
            
            ' Detect the stock symbol change
            If (currentSymbol <> nextSymbol) Then
                MsgBox ("Current Symbol: " & currentSymbol & _
                    ", Next Symbol: " & nextSymbol)
                MsgBox ("First Open Price: " & firstOpenPrice)
                firstOpen = True
            End If
            
        Next i
        
    Next ws
    
End Sub

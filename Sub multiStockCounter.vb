Sub stockCounter

Dim lastRow as Long
Dim stockNumber as Integer
Dim yearOpen as Double
Dim yearClose as Double
Dim stockVolumeCounter as Double
Dim percentChange as Double
Dim yearChange as Double
Dim greatestIncrease as Double
Dim greatestStock as String
Dim least as Double
Dim leastStock as String
Dim greatestTotal as Double
Dim totalStock as String
Dim stockName as String
For each ws in worksheets

ws.activate

Cells(1,9)="Ticker"
Cells(1,10)="Yearly Change"
Cells(1,11)="Percent Change"
Cells(1,12)="Total Stock Volume"

stockNumber = 1
greatestIncrease = 0
least = 0
greatestTotal = 0

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i =  2 to lastrow
    stockName=Cells(i,1).Value

    If stockName<>Cells(i-1,1).Value Then
        stockNumber=stockNumber+1
        Cells(StockNumber,9)=stockName
        yearOpen=cells(i,3).Value

        stockVolumeCounter=Cells(i,7)
    
    ElseIf stockName<>Cells(i+1,1).Value Then
        yearClose=Cells(i,6).Value
        yearChange=yearClose-yearOpen
        Cells(StockNumber,10)=yearChange
        if yearChange<0 Then
            Cells(StockNumber,10).Interior.ColorIndex = 3
        Else
            Cells(StockNumber,10).Interior.ColorIndex = 4
        End if

        if yearOpen<>0 Then
            percentChange=(yearClose-yearOpen)/yearOpen
        Else
            percentChange=(yearClose)
        End If
        Cells(StockNumber,11)=percentChange
        Cells(StockNumber,11).NumberFormat="0.00%"
        if percentChange<0 Then
            Cells(StockNumber,11).Interior.ColorIndex = 3
        Else
            Cells(StockNumber,11).Interior.ColorIndex = 4
        End if

        if percentChange>greatestIncrease Then
            greatestIncrease=percentChange
            greatestStock=stockName
        End If

        if percentChange<least Then
            least=percentChange
            leastStock=stockName
        End if

        stockVolumeCounter=stockVolumeCounter+Cells(i,7)
        Cells(StockNumber,12)=stockVolumeCounter
        Cells(StockNumber,12).NumberFormat = "@"

        if stockVolumeCounter>greatestTotal Then
            greatestTotal=stockVolumeCounter
            totalStock=stockName
        End if

    Else
        stockVolumeCounter=stockVolumeCounter+Cells(i,7)

    End If

Next i

Cells(1,15)="Ticker"
Cells(1,16)="Value"
Cells(2,14)="Greatest % Increase"
Cells(3,14)="Greatest % Decrease"
Cells(4,14)="Greatest Total Volume"
cells(2,16)=greatestIncrease
Cells(2,16).NumberFormat="0.00%"
cells(2,15)=greatestStock
Cells(3,16)=least
Cells(3,16).NumberFormat="0.00%"
Cells(3,15)=leastStock
Cells(4,16)=greatestTotal
Cells(4,16).NumberFormat="@"
Cells(4,15)=totalStock

Next ws

End Sub
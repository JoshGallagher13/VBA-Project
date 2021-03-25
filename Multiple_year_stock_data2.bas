Attribute VB_Name = "Module1"
Sub ticker()

For Each ws In Worksheets

'setting headers
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Price Change"
ws.Range("K1") = "% Change"
ws.Range("L1") = "Total Volume"

Dim lastRow As Long
Dim startPrice As Double
Dim endPrice As Double
Dim curRow As Long
Dim tot_vol As Double
'row where unique ticker table will be placed
curRow = 2
'returns last row of initial data
lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
'start price of first ticker
startPrice = ws.Range("C2")
'volume counter
tot_vol = 0
    
    'loops through initial data
    For r = 2 To lastRow
        'skips tickers with $0 prices
        If ws.Range("C" & r) <> 0 Then
            'if ticker is a unique value, place into new table
            If ws.Range("A" & r) <> ws.Range("A" & r + 1) Then
                'ticker
                ws.Range("I" & curRow) = ws.Range("A" & r)
                endPrice = ws.Range("F" & r)
                'change in price
                ws.Range("J" & curRow) = endPrice - startPrice
                'percent change in price
                ws.Range("K" & curRow) = (endPrice - startPrice) / startPrice
                tot_vol = tot_vol + ws.Range("G" & r)
                ws.Range("L" & curRow) = tot_vol
                'set starting price of next ticker
                startPrice = ws.Range("C" & r + 1)
                'move to next row on table
                curRow = curRow + 1
                'reset volume counter
                tot_vol = 0
            Else
                'increase volume counter
                tot_vol = tot_vol + ws.Range("G" & r)
            End If
        Else
            'in the event of $0 price, setting value
            startPrice = ws.Range("C" & r + 1)
        End If
        
    Next r
'last row of unique ticker table
lastRow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row
curRow2 = 2
'setting variables for greatest % increase, % decrease, and volume total
Dim G_inc As Double
Dim G_dec As Double
Dim G_tot As Double
G_inc = 0
G_dec = 0
G_tot = 0

'setting headers of bonus table
ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Total Volume"
ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"
    
    'formatting colors of positives and negatives
    For r = 2 To lastRow2
        If ws.Range("J" & r) > 0 Then
            ws.Range("J" & r).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & r) < 0 Then
            ws.Range("J" & r).Interior.ColorIndex = 3
        End If
    Next r
    
    'finding greatest values
    For r = 2 To lastRow2
        If ws.Range("K" & r) > G_inc Then
            G_inc = ws.Range("K" & r)
            ws.Range("O2") = ws.Range("I" & r)
            ws.Range("P2") = G_inc
        End If
    Next r
    
    For r = 2 To lastRow2
        If ws.Range("K" & r) < G_dec Then
            G_dec = ws.Range("K" & r)
            ws.Range("O3") = ws.Range("I" & r)
            ws.Range("P3") = G_dec
        End If
    Next r
    
    For r = 2 To lastRow2
        If ws.Range("L" & r) > G_tot Then
            G_tot = ws.Range("L" & r)
            ws.Range("O4") = ws.Range("I" & r)
            ws.Range("P4") = G_tot
        End If
    Next r

'formatting bonus table
ws.Rows(1).Font.Bold = True
ws.Columns("K").NumberFormat = "0.0%"
ws.Columns("P").NumberFormat = "0.0%"
ws.Columns("J").NumberFormat = "0.00"
ws.Range("P4").NumberFormat = "000"
ws.Columns("L").ColumnWidth = 14
ws.Columns("N").ColumnWidth = 19
            
    
  
 Next ws
        
End Sub


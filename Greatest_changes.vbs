'''' This sub calculates the greatest changes within the prices

Sub Greatest_Changes(ByRef ws As Worksheet):

Dim lRow_unique As Variant 'last row of data form the list of unique tickers
Dim max_pct_increase As Variant 'maximum percet of increase
Dim max_pct_decrease As Variant 'maximum percet of decrease
Dim max_volume As Variant 'maximum volume
Dim max_ticker_increase As Variant 'tikcer name for the maximum increase
Dim max_ticker_decrease As Variant 'ticker name for the maximum decrease
Dim max_ticker_volume As Variant 'ticker name for the maximum volume


Dim i As Double 'control variable
Dim j As Double 'control variable


ws.Activate 'activate the worksheet


lRow_unique = ws.Cells(Rows.Count, 11).End(xlUp).Row 'last row from the unique tickers list


'initialize the variables

max_pct_increase = 0
max_pct_decrease = 0
max_volume = 0

max_ticker_increase = 0
max_ticker_decrease = 0
max_ticker_volume = 0



For i = 2 To lRow_unique
    
    If ws.Cells(i, 14).Value > max_pct_increase Then

        max_pct_increase = ws.Cells(i, 14).Value 'change the value of max_pct_increase if finds a greater value
        max_ticker_increase = ws.Cells(i, 11).Value 'gets the name of the ticker of the max pct increase

    End If

    If ws.Cells(i, 14).Value < max_pct_decrease Then

        max_pct_decrease = ws.Cells(i, 14).Value 'change the value of max_pct_decrease if finds a lower value
        max_ticker_decrease = ws.Cells(i, 11).Value 'gets the name of the ticker of the max pct decrease

    End If

    If ws.Cells(i, 12).Value > max_volume Then

        max_volume = ws.Cells(i, 12).Value 'change the value of max_volume if finds a greater value

        max_ticker_volume = ws.Cells(i, 11).Value 'gets the name of the ticker of the max volume

    End If

Next i

'column headers

ws.Cells(1, 17).Value = "Concept"
ws.Cells(1, 18).Value = "Ticker"
ws.Cells(1, 19).Value = "Value"

'row concepts

ws.Cells(2, 17).Value = "Greatest % increase"
ws.Cells(3, 17).Value = "Greatest % decrease"
ws.Cells(4, 17).Value = "Greatest Total Volume"

'set names of the tickers in the correspondig cells

ws.Cells(2, 18).Value = max_ticker_increase
ws.Cells(3, 18).Value = max_ticker_decrease
ws.Cells(4, 18).Value = max_ticker_volume

'set the values in each corresponding cell

ws.Cells(2, 19).Value = max_pct_increase
ws.Cells(3, 19).Value = max_pct_decrease
ws.Cells(4, 19).Value = max_volume

'format the cells

ws.Cells(2, 19).NumberFormat = "0.00%"
ws.Cells(3, 19).NumberFormat = "0.00%"
ws.Cells(3, 19).Font.ColorIndex = 3

End Sub



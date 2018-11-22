'''' This Sub calculates the yearly change of the price
'''The main idea is to find the first row and last row when the ticker appears to get the first open price and last close price of the year

Sub Year_change(ByRef ws As Worksheet):

Dim lRow_unique As Double 'last row with data from the unique tickers
Dim lRow_tickers As Double 'las row with data form the tickers column
Dim unique As Variant 'store the name of the unique tickers
Dim firstRow As Double 'first row from the tickers column in which the ticker was found
Dim lastRow As Double 'last row from the tickers column in which the ficker was found
Dim open_prc As Double 'open price from the first day of data
Dim close_prc As Double 'close price from the last day of data
Dim yrl_change As Double 'price change
Dim pct_change As Double 'percent of price change

ws.Activate 'activate worksheet


ws.Cells(1, 13) = "Yr_change" 'column header
ws.Cells(1, 14) = "Pct_change" 'column header

lRow_unique = ws.Cells(Rows.Count, 11).End(xlUp).Row 'last row with data from unique tickers list
lRow_tickers = ws.Cells(Rows.Count, 1).End(xlUp).Row 'last row with data from tickers list

For i = 2 To lRow_unique ' loops the unique tickers list

    unique = ws.Cells(i, 11).Value ' store the name of the ticker from the unique tickers list

    firstRow = ws.Range(ws.Cells(1, 1), ws.Cells(lRow_tickers, 1)).Find(what:=unique, LookAt:=xlWhole).Row 'Get the first row where the ticker appears
    lastRow = ws.Range(ws.Cells(1, 1), ws.Cells(lRow_tickers, 1)).Find(what:=unique, LookAt:=xlWhole, searchdirection:=xlPrevious).Row 'Get the last row where the ticker appears

    open_prc = ws.Cells(firstRow, 3).Value 'store the open price from the first day of data of the ticker
    close_prc = ws.Cells(lastRow, 6).Value 'store the close price from the last day of data of the ticker
    
    If open_prc <= 0 Then 'prevents to divide by cero or to have negative prices
    
        yrl_change = close_prc - open_prc
        
        pct_change = 0 ' if the open price is negative or cero then set pct_change to cero
    
        ws.Cells(i, 13).Value = yrl_change
        
        ws.Cells(i, 14).Value = pct_change
    
    Else
    
        yrl_change = close_prc - open_prc 'calculate yrl change
        
        pct_change = yrl_change / open_prc 'calculate % change
    
        ws.Cells(i, 13).Value = yrl_change
        
        ws.Cells(i, 14).Value = pct_change

        ws.Cells(i, 14).NumberFormat = "0.00%"
    
            If yrl_change < 0 Then
    
    
                ws.Cells(i, 13).Interior.ColorIndex = 3  'set red color if the change is negative
    
            Else
            
                ws.Cells(i, 13).Interior.ColorIndex = 4 'set green color if the change is positive
    
            End If
    
    End If

Next i


End Sub



''''' This sub uses a faster way to make a unique values list an gather the traded volume for each ticket

Public Sub List_and_volume(ByRef ws As Worksheet):

Dim lRow_tickers As Double 'last row for the tickers column
Dim lRow_unique As Double 'last row for the unique tickers list
Dim unique As Variant 'ticker name from the unique tickers list
Dim firstrow_ticker As Double 'first row when the ticker name is found
Dim lastrow_ticker As Double 'last row when the ticker name is found
Dim i As Double 'control variable
Dim j As Double 'control variable
Dim unique_row As Double 'keeps the count in which unique ticket list row is the loop to paste the total volume in that cell

    ws.Activate 'activate the worksheet
    
    lRow_tickers = ws.Cells(Rows.Count, 1).End(xlUp).Row 'last row of data for the tickers column
    
    ws.Range("A1:G" & lRow_tickers).Sort Key1:=ws.Range("A1"), Key2:=ws.Range("B1"), Order1:=xlAscending, Header:=xlYes 'sort the range by ticker name and date
     
    ws.Range("A1:A" & lRow_tickers).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("K1"), unique:=True 'makes a list of unique tickers and copy it to a new range
    
    lRow_unique = ws.Cells(Rows.Count, 11).End(xlUp).Row 'gets the last row for the unique tickers list
          
    unique_row = 2 'The first unique ticket list name is in the second row
        
        For Each unique In Range(Cells(2, 11), Cells(lRow_unique, 11)) 'loops the unique tickers list
                                        
            firstrow_ticker = ws.Range(ws.Cells(1, 1), ws.Cells(lRow_tickers, 1)).Find(what:=unique, LookAt:=xlWhole).Row 'gets the first row from the tickers list when the ticker name appears
            lastrow_ticker = ws.Range(ws.Cells(1, 1), ws.Cells(lRow_tickers, 1)).Find(what:=unique, LookAt:=xlWhole, searchdirection:=xlPrevious).Row ' gets the last row from the tickers list when the ticker name appears

            Tot_volume = 0 'reset the total volume variable each time another ticket is looped
            
                For j = firstrow_ticker To lastrow_ticker 'from the first time each ticker name appears to the last time this loops adds the total volume
                
                    Tot_volume = Tot_volume + ws.Cells(j, 7).Value
                    
                 Next j
                 
            ws.Cells(unique_row, 12) = Tot_volume 'print the total volume for the unique ticker list
            
            unique_row = unique_row + 1 ' moves to the next name in the unique ticker list
              
         Next
         
'columns headers

ws.Cells(1, 11).Value = "Ticker"
ws.Cells(1, 12).Value = "Volume"

End Sub


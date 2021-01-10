Attribute VB_Name = "Module1"
Sub stockdata()

'declare variables
Dim ticker As String
Dim beg_price As Double
Dim end_price As Double
Dim price_change As Double
Dim price_percent_change As String
Dim volume As Double

'loop through each worksheet
For Each ws In Worksheets
    
    'make sure it goes through each sheet
    ws.Activate
    
        'create headers for the summary table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
        
        ' Set an initial variable for holding the volume per ticker and the first row of a ticker
        volume = 0
        j = 2
                
        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'count the number of rows in the sheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all stock tickers
        For i = 2 To lastrow
        
          ' Check if we are still within the same ticker, if it is not...
          If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ' Set the ticker name
            ticker = Cells(i, 1).Value
        
            ' Add to the volume
            volume = volume + Cells(i, 7).Value
            
            ' Add the beginning price
            beg_price = Cells(j, 6).Value
            
            'Add the end price
            end_price = Cells(i, 6).Value
            
            'calculate the price change
            price_change = end_price - beg_price
            
            'remove divide by zero errors
            If beg_price = 0 Then
            
                'fill cell if denominator is zero
                price_percent_change = "N/A"
                
                Else
                        
                'calculate the percentage change
                price_percent_change = Format((end_price / beg_price) - 1, "Percent")
            
            End If
            
            ' Print the ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = ticker
            
            ' Print the maxprice in the Summary Table
            Range("J" & Summary_Table_Row).Value = price_change
            
            'format cells as green/red when positive/negative
            If price_change > 0 Then
                
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            Else
            
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
                            
            ' Print the maxprice in the Summary Table
            Range("K" & Summary_Table_Row).Value = price_percent_change
            
            ' Print the volume to the Summary Table
            Range("L" & Summary_Table_Row).Value = volume
        
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the volume
            volume = 0
            'reset j to the row of the last ticker change
            j = i
        
          ' If the cell immediately following a row is the same ticker...
          Else
        
            ' Add to the volume
            volume = volume + Cells(i, 7).Value
                                                  
          End If
        
        Next i
  
    Next ws

End Sub


Sub sorting()

Dim ticker As String
Dim stock_date As Date
Dim stock_vol As Double
Dim stock_yearly_change As Double
Dim stock_percent_change As Double
Dim OpenValue As Double
Dim CloseValue As Double

Dim ws As Worksheet


Dim summary_table_row As Long
summary_table_row = 2

y = Cells(Rows.Count, 1).End(xlUp).Row

' for worksheet1
 For Each ws In ThisWorkbook.Sheets
ticker = " "
summary_table_row = 2

For i = 2 To y


      ' sets up ticker value
    If ws.Cells(i, 1).Value <> ticker Then
      ' set open value
      OpenValue = ws.Cells(i, 3).Value
      ' set ticker
      ticker = ws.Cells(i, 1).Value
    
      ' check if we are still within the same ticker, if it is not...
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the ticker name
          ticker = ws.Cells(i, 1).Value

             ' Add to the Ticker Volume
          stock_vol = stock_vol + ws.Cells(i, 7).Value
 
             ' Print the Ticker label in the Summary Table
          ws.Range("I" & summary_table_row).Value = ticker

             ' Print the total stock volume to the Summary Table
          ws.Range("L" & summary_table_row).Value = stock_vol

             ' assign the CloseValue
          CloseValue = ws.Cells(i, 6).Value
     
             ' print this in a different if statement
          stock_yearly_change = (CloseValue) - (OpenValue)
      
             ' Print the yearly change to the Summmary Table
          ws.Range("J" & summary_table_row).Value = stock_yearly_change
    
    
             ' Add one to the summary table row
          summary_table_row = summary_table_row + 1
          
          stock_vol = 0
        Else
        
        
' If the cell immediately following a row is the same ticker...
        
      ' Add to the Stock Volume
      stock_vol = stock_vol + ws.Cells(i, 7).Value

      End If
      
            ' hard code percent change as 0
    If OpenValue = 0 Or CloseValue = 0 Then
                stock_percent_change = 0
                ws.Range("K" & summary_table_row).Value = stock_percent_change
                
            ' proceed with normal percent change calculation
        ElseIf OpenValue <> 0 Or CloseValue <> 0 Then
                stock_percent_change = ((stock_yearly_change) / OpenValue) * 100
                ws.Range("K" & summary_table_row).Value = stock_percent_change
                 

    End If

' If ws.Cells(y, 11).Value > 0 Then
' ws.Cells(y, 11).Value.Interior.ColorIndex = 4
' Else
' ws.Cells(y, 11).Value.Interior.ColorIndex = 3
' End If
' was not able to get this to work. :(

    Next i
    
    
     Next ws


End Sub



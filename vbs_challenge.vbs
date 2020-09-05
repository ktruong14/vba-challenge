Sub Stock()

Dim ticker_name As String
Dim yearly_change As Double
Dim percent_change As Double
Dim last_row As Long
Dim table_row As Long
Dim total_vol As Double
Dim Ws As Integer
Dim counter As Long

' Looping each worksheet
For Ws = 1 To Sheets.Count
    Sheets(Ws).Activate
    
    ' Initializing counters and variables
    table_row = 2
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    counter = 0
    total_vol = 0
   
 ' Adding headers to Columns I, J, K and L respectively
 Cells(1, 9).Value = "Ticker"
 Cells(1, 10).Value = "Yearly Change"
 Cells(1, 11).Value = "Percent Change"
 Cells(1, 12).Value = "Total Stock Volume"
   
   ' Looping each individual sheet letter
   For i = 2 To last_row
   
     

     ' Checks if ticker name is not the same to the next subsequent cell
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Reading information from each column: ticker name, opening price, closing price, yearly change, percent change and total volume
        ' Yearly change Percent change, Total Volume calculations
        ticker_name = Cells(i, 1).Value
        open_price = Cells(i - counter, 3).Value
        close_price = Cells(i, 6).Value
        yearly_change = close_price - open_price
        total_vol = total_vol + Cells(i, 7).Value
        
        If open_price > 0 Then
        
           percent_change = (close_price - open_price) / open_price
        
        Else
        
           percent_change = 0
        
        End If

        ' Prints info to columns I, J, L and K respectively
        Range("I" & table_row).Value = ticker_name
        Range("J" & table_row).Value = yearly_change
        Range("L" & table_row).Value = total_vol
        Range("K" & table_row).Value = percent_change
        
        ' Add one to table row
        table_row = table_row + 1
        
        ' Resetting the total volume
        total_vol = 0
        
        ' Resetting the counter
        counter = 0
        
        Else
        
        ' Adding total volume if ticker names are the same
        total_vol = total_vol + Cells(i, 7).Value
        counter = counter + 1
        
        End If
            
   Next i

 For j = 2 To last_row
 
   Cells(j, 11).NumberFormat = "0.00%"

   If Cells(j, 10).Value <= 0 Then

      Cells(j, 10).Interior.ColorIndex = 3

   Else
        
      Cells(j, 10).Interior.ColorIndex = 4
        
   End If
   
 Next j
   
Next Ws



End Sub

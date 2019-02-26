Sub Stock_Volume()

Dim i As Long
Dim j As Long
Dim lrow As Long

Dim row_counter As Long
Dim total_volume As Double

Dim start_price As Double
Dim close_price As Double
Dim change_price As Double
Dim percent_change As Double





'Identify start and end dates

start_price = Cells(2, 3).Value



'use ticker_counter to help set cells for ticker names and create total volume variable
row_counter = 2
total_volume = 0

'label our summary table headers

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"

'identify what the last row of this column is

lrow = Cells(Rows.Count, 1).End(xlUp).Row

' Verify last row number with MsgBox (lrow)

For i = 2 To lrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ' Insert the ticker names into our summary table
        Cells(row_counter, 9).Value = Cells(i, 1).Value
        
        
        'We need to make sure to add final volume for each ticker
        total_volume = total_volume + Cells(i, 7).Value
        
        'Insert the total volume into summary table
        Cells(row_counter, 10).Value = total_volume
        
        row_counter = row_counter + 1
        total_volume = 0
        
    Else
    
        'add our total volumes to our counter
        
        total_volume = total_volume + Cells(i, 7).Value
        Cells(row_counter, 10).Value = total_volume
        
     End If
     
    
    Next i
    
    



End Sub

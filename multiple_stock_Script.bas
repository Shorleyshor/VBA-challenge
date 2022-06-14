Attribute VB_Name = "Module1"
Sub worksheet()
     Dim xSh As worksheet
     Application.ScreenUpdating = False
     For Each xSh In Worksheets
        xSh.Select
        Call stock_data
    Next
    Application.ScreenUpdating = True

End Sub

Sub stock_data()
 
 'keep track of locating each ticker in the table
 Dim Table_row As Integer
 Dim starting_price As Double
 Dim closing_price As Double
 Dim yearly_change As Double
 Dim percentage_change As Double
 Dim mycell As Range
 Dim myrange As Range
 Dim myrange2 As Range
 
 
 Total_volume = 0
 Table_row = 2
 
 'print the headings
 Range("I1").Value = "Tcker"
 Range("J1").Value = "Volume"
 Range("M1").Value = "yearly"
 Range("N1").Value = "percentage"
 Range("R1").Value = "Value"
 Range("Q1").Value = "Ticker"
 Range("P2").Value = "Greatest % Increase"
 Range("P3").Value = "Greatest % Decrease"
 Range("P4").Value = "Greatest Total Volume"

'set the last row function
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'loop through all ticker

For i = 2 To lastrow

    'check if we are still within the same ticker. if its not..
    'if the next cell is not equals to the current cell
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'set ticker name
        Ticker_Name = Cells(i, 1).Value
        
        'closing price is the last price when the current cell is not equal to the next cell
        closing_price = Cells(i, 6).Value
        
        
        
        'add to the ticker volume
        Total_volume = Total_volume + Cells(i, 7).Value
        
        'print the ticker name in the Table
        Range("I" & Table_row).Value = Ticker_Name
        
        'print Ticker Volume to the Table
        Range("J" & Table_row).Value = Total_volume
        
        'calculate the yearly change in price
        yearly_change = closing_price - opening_price
        
       'print the yearly change in price in the table
        Range("M" & Table_row).Value = closing_price - opening_price
        
        'set how you want the number to appear in the table
        Range("M" & Table_row).NumberFormat = "0.00"
        
        'calculate for the percentage change in price
        percentage_change = (closing_price - opening_price) / opening_price
        
        'print the percentage change in price on the table
        Range("N" & Table_row).Value = (closing_price - opening_price) / opening_price
        
        'set how you want the number to appear in the table
         Range("N" & Table_row).NumberFormat = "0.00%"
        
        'add one to the table row
        Table_row = Table_row + 1
        
        'reset the ticker volume
        Total_volume = 0
        
        'if the current cell is not equals to the previous cells
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
        'opening price is...
        opening_price = Cells(i, 3).Value
        
        
        'if the cell immediately following a row is the same ticker...
        Else
        
        'add to the total volume
        Total_volume = Total_volume + Cells(i, 7).Value
        
    End If
    
Next i

'conditional formating of negative and positive cells in yearly change column
For i = 2 To lastrow

        If Cells(i, 13).Value < 0 Then
        Cells(i, 13).Interior.ColorIndex = 3
        
        Else
        Cells(i, 13).Interior.ColorIndex = 4
    
        End If
    Next i
    
    
'Bonus question.
Set myrange = Worksheets("2018").Range("N2:N" & lastrow)
Set myrange2 = Worksheets("2018").Range("J2:J" & lastrow)

'Greatest % increase
Range("R2") = Application.WorksheetFunction.Max(myrange)
Range("R2").NumberFormat = "0.00%"

'loop to retreive ticker name
For i = 2 To lastrow
    If Range("R2") = Cells(i, 14).Value Then
    Range("Q2") = Cells(i, 9)
    End If
Next i

'Greatest % decrease
Range("R3") = Application.WorksheetFunction.Min(myrange)
Range("R3").NumberFormat = "0.00%"

'loop to retreive ticker name
For i = 2 To lastrow
    If Range("R3") = Cells(i, 14).Value Then
    Range("Q3") = Cells(i, 9)
    End If
Next i

'Greatest Total volume
Range("R4") = Application.WorksheetFunction.Max(myrange2)
Range("R4").NumberFormat = "0,0"

'loop to retreive ticker name
For i = 2 To lastrow
    If Range("R4") = Cells(i, 10).Value Then
    Range("Q4") = Cells(i, 9)
    End If
Next i

    
    
End Sub




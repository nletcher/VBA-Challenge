Sub Program()
'Declare variables
Dim Ticker As String
Dim Opening_price As Double
Dim Closing_price As Long
Dim Percentage_change As Double
Dim Yearly_change As Double

Dim Total As LongLong

Dim i As Long
Dim Lastrow As Long
Dim Summary_table_row As Integer
Dim Max As Long
Dim Min As Long
Dim Max_Volume As LongLong
Dim No_tickers As Long
Dim Max_ticker As Long
Dim Min_ticker As Long
Dim Max_Volume_ticker As Long

'Initializing variables

Total = 0
Summary_table_row = 2
Lastrow = Cells(Rows.Count, "A").End(xlUp).Row

'Headers for summary table

Range("L1").Value = "Ticker"
Range("M1").Value = "Yearly Change"
Range("N1").Value = "Percent Change"
Range("O1").Value = "Total Stock Volume"

'Headers for #challenge table
Range("Q3").Value = "Greatest % Increase"
Range("Q4").Value = "Greatest % Decrease"
Range("Q5").Value = "Greatest Total Volume"

'For Loop for all rows
For i = 2 To Lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Total = Total + Cells(i, 7).Value
        Opening_price = Cells(i, 3).Value
        
        Closing_price = Cells(i, 6).Value
        
        Yearly_change = (Closing_price - Opening_price)
        
        If Opening_price = 0 Then
            Percentage_change = 0
        Else
        Percentage_change = (Yearly_change / Opening_price) * 100
        End If
        If Yearly_change >= 0 Then
            Range("M" & Summary_table_row).Interior.ColorIndex = 4
            
        Else
            Range("M" & Summary_table_row).Interior.ColorIndex = 3
        
        End If
        

        Range("L" & Summary_table_row).Value = Ticker
        Range("M" & Summary_table_row).Value = Yearly_change
        Range("N" & Summary_table_row).Value = Percentage_change
        Range("O" & Summary_table_row).Value = Total
        
        Summary_table_row = Summary_table_row + 1
        Total = 0
    Else
    
        Total = Total + Cells(i, 7).Value
    
    
    End If
    
Next i


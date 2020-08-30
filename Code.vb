Sub RunSheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Program
    Next
    Application.ScreenUpdating = True
     
End Sub

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

'Formating Percentage Change Column to %
Columns(14).NumberFormat = "0.00%"

'Finding Max and Min Values in summary table
No_tickers = Cells(Rows.Count, "N").End(xlUp).Row

Max = Application.WorksheetFunction.Max(Range("N2:N" & No_tickers))
Min = Application.WorksheetFunction.Min(Range("N2:N" & No_tickers))
Max_Volume = Application.WorksheetFunction.Max(Range("O2:O" & No_tickers))


Cells(3, 19).Value = Max
'Formating to Percentage
Cells(3, 19).NumberFormat = "0.00%"


Cells(4, 19).Value = Min
'Formating to Percentage
Cells(4, 19).NumberFormat = "0.00%"


Cells(5, 19).Value = Max_Volume

'Find the corresponding ticker
Max_ticker = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("N2:N" & No_tickers)), Range("N2:N" & No_tickers), 0)
Min_ticker = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Range("N2:N" & No_tickers)), Range("N2:N" & No_tickers), 0)

Max_Volume_ticker = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("O2:O" & No_tickers)), Range("O2:O" & No_tickers), 0)

Cells(3, 18).Value = Cells(Max_ticker + 1, 12)
Cells(4, 18).Value = Cells(Min_ticker + 1, 12)
Cells(5, 18).Value = Cells(Max_Volume_ticker + 1, 12)

End Sub

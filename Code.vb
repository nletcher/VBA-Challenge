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



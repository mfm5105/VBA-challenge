Attribute VB_Name = "Module1"
Sub stock_data()
'Set worksheet object variable to loops through all of the spreadsheets
Dim ws As Worksheet
'-------------------------------------------
'LOOP THROUGH ALL OF THE WORKSHEETS
'-------------------------------------------
For Each ws In Worksheets
'-------------------------------------------
'SETTING HEADERS
'-------------------------------------------
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
'-------------------------------------------
'SETTING VARIABLES
'-------------------------------------------
'Set ticker name
Dim ticker_name As String
ticker_name = " "
'Set ticker volume total
Dim ticker_volume_total As Double
ticker_volume_total = 0
'Set open price
Dim open_price As Double
open_price = 0
'Set close price
Dim close_price As Double
close_price = 0
'Set price_difference
Dim price_difference As Double
price_difference = 0
'Set percent change
Dim percent_change As Double
percent_change = 0
'summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
'start of stock
Dim start_stock As Long
start_stock = 2
'max volume ticker name
Dim max_volume_ticker As String
max_volume_ticker = " "
'max volume
Dim max_volume As Double
max_volume = 0
'greatest percent increase
greatest_percent_increase = 0

'greatest percent decrease
greatest_percent_decrease = 0

'Set worksheet name variable
Dim WorksheetName As String
Dim LastRow As Long
'Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'-------------------------------------------
'SETTING INITIAL VALUE FOR OPEN PRICE
'----------------------------------------
open_price = ws.Cells(2, 3).Value
'------------------------------------------
'LOOP FROM INITIAL OF CURRENT WORKSHEET UNTIL ITS LAST ROW
'------------------------------------------
For i = 2 To LastRow
'Add to ticker total
ticker_volume_total = ticker_volume_total + ws.Cells(i, 7).Value
'Checking if we are still within the same ticker
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'Set ticker name
    ticker_name = ws.Cells(i, 1).Value
    'Printing ticker name
    ws.Range("I" & Summary_Table_Row).Value = ticker_name
    'Start of stock
    start_stock = i + 1
    'Close pricing
    close_price = ws.Cells(i, 6).Value
    price_difference = close_price - open_price
    ws.Range("J" & Summary_Table_Row).Value = price_difference
    'Divisible by zero resolution
    If open_price > 0 Then
    price_difference = close_price - open_price
    ws.Range("J" & Summary_Table_Row).Value = price_difference
    Else
    ws.Range("J" & Summary_Table_Row).Value = 0
    End If
    ws.Range("L" & Summary_Table_Row).Value = ticker_volume_total
    'Conditional formatting for price change
            If price_difference > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf price_difference <= 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
    'Calculate and Print in Summary Table
            If open_price > 0 Then
                percent_change = ((close_price - open_price) / open_price) * 100
                ws.Range("K" & Summary_Table_Row).Value = (CStr(percent_change) & "%")
            Else
            ws.Range("K" & Summary_Table_Row).Value = 0
            End If
            'Calculate max increase to find the ticker name and value
            If percent_change > greatest_percent_increase Then
            greatest_percent_increase = percent_change
            ws.Range("Q2").Value = percent_change
            ws.Range("P2").Value = ticker_name
            End If
            'Calculate min decrease to find the ticker name and value
            If percent_change < greatest_percent_decrease Then
            greatest_percent_decrease = percent_change
            ws.Range("Q3").Value = percent_change
            ws.Range("P3").Value = ticker_name
            End If
            'Calculate max total value to find the ticker name and value
            If ticker_volume_total > max_volume Then
            max_volume = ticker_volume_total
            max_volume_ticker = ticker_name
            ws.Range("Q4").Value = ticker_volume_total
            ws.Range("P4").Value = ticker_name
            End If
    'Adding another row to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
'Capturing the next ticker's open price
  ElseIf i = start_stock Then
        open_price = ws.Cells(i, 3).Value
        ticker_volume_total = 0
'Add to ticker total
    ticker_volume_total = ticker_volume_total + ws.Cells(i, 7).Value
        End If
   Next i
Next ws
End Sub

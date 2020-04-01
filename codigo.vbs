Option Explicit

Sub stockmarket()

'Declare variables
Dim ticker As String
Dim t, stocksheets_number As Integer
Dim row_count, daterange_count, first_open, last_close, o As Long
Dim total_stock, year_change, percent_change, min_percent, max_percent, max_stock As Double
Dim stocksheet As Worksheet

Application.ScreenUpdating = True

'Loop through workbook sheets
Set stocksheet = ActiveSheet
    stocksheets_number = ThisWorkbook.Worksheets.Count
For t = 1 To stocksheets_number
    ThisWorkbook.Worksheets(t).Activate

'Set variable values
ticker = 2
total_stock = 0
first_open = 0
last_close = 0
year_change = 0
percent_change = 0
daterange_count = 2
min_percent = 0
max_percent = 0
max_stock = 0

'Loop through rows to count rows with data
row_count = Cells(Rows.Count, 1).End(xlUp).Row

'Headers of range of columns to fill with data and titles for table range
Range("I1").Value = "Ticker variables"
Range("J1").Value = "Year change value"
Range("K1").Value = "Year change percentage value"
Range("L1").Value = "Year stock volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest percentage increase"
Range("N3").Value = "Greatest percentage decrease"
Range("N4").Value = "Greatest stock"
Range("J:J").NumberFormat = "0.00"
Range("K:K").NumberFormat = "0.00%"
Range("P2").NumberFormat = "0.00%"
Range("P3").NumberFormat = "0.00%"

'Loop through data with conditionals
For o = 2 To row_count
    total_stock = total_stock + Cells(o, 7).Value
    If Cells(o, 1) <> Cells(o - 1, 1) Then
        first_open = Cells(o, 3).Value
    End If
    'Calculate the year change by subtracting the last date value of the close section per ticker and the first date value of the open section per ticker
    'Calculate the year percentage change by dividing the last date value of the close section per ticker and the first date value of the open section per ticker
    If Cells(o, 1) <> Cells(o + 1, 1) Then
        last_close = Cells(o, 6).Value
        Cells(daterange_count, 9).Value = Cells(o, 1).Value
        Cells(daterange_count, 10).Value = last_close - first_open
        If first_open <> 0 Then
            Cells(daterange_count, 11).Value = (last_close - first_open) / first_open
        Else
            Cells(daterange_count, 11).Value = 0
        End If
            Cells(daterange_count, 12).Value = total_stock
            daterange_count = daterange_count + 1
            'Restart values for new cycle
            first_open = 0
            last_close = 0
            total_stock = 0
    End If
Next o
     
For o = 2 To row_count
    If Cells(o, 11).Value < min_percent Then
        min_percent = Cells(o, 11).Value
        Cells(3, 15).Value = Cells(o, 9).Value
        Cells(3, 16).Value = min_percent
    End If
    If Cells(o, 11).Value > max_percent Then
        max_percent = Cells(o, 11).Value
        Cells(2, 15).Value = Cells(o, 9).Value
        Cells(2, 16).Value = max_percent
    End If
     If Cells(o, 12).Value > max_stock Then
        max_stock = Cells(o, 12).Value
        Cells(4, 15).Value = Cells(o, 9).Value
        Cells(4, 16).Value = max_stock
    End If

    'Color formatting for year_change values positive (green) and negative (red)
    If Cells(o, 10).Value >= 0 Then
        Cells(o, 10).Interior.ColorIndex = 4
    ElseIf Cells(o, 10).Value < 0 Then
        Cells(o, 10).Interior.ColorIndex = 3
    End If
Next o

Next t
stocksheet.Activate

End Sub


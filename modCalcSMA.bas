Attribute VB_Name = "Module1"
Option Explicit

'Calculates Simple Moving Average for a predefined period
'By George Slinger, github/gslinger91, MSc International Banking Student

'Enter closing date data in column A with header

Const PERIOD As Integer = 5

Sub calcSMA()

Dim wksSMA As Worksheet
Set wksSMA = ActiveWorkbook.Sheets("SMA")
 
Dim lastrow As Integer
lastrow = wksSMA.UsedRange.Rows.Count 'finds last row

Dim firstrow As Integer
firstrow = 2 + (PERIOD - 1) 'first row depends on PERIOD

Dim i As Integer
For i = firstrow To lastrow
    wksSMA.Cells(i, 2).Value = WorksheetFunction.Average(Range(Cells(i, 1), _
                                Cells(i - (PERIOD - 1), 1)))  'uses excels average function to find SMA for each period
Next i

Call formatSMA(wksSMA, lastrow)

End Sub

Sub formatSMA(wks As Worksheet, lrow As Integer)

'Optional formatting

With wks

    .Cells(1, 1).Value = "Close Price"
    .Cells(1, 2).Value = "SMA-" & PERIOD

    With .Range("A1:B1")
        .Font.Bold = True
        .ColumnWidth = 10.29
        .RowHeight = 28.5
        .WrapText = True
        .EntireColumn.HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous 'creates single bottom border for cell
    End With

    .Range("B2:B" & PERIOD).Interior.ColorIndex = 16 'fills unused cells (between row 1 and period)
    
    .Range(Cells(2, 1), Cells(lrow, 2)).NumberFormat = "0.0000" ' 3 decimal places

End With

End Sub

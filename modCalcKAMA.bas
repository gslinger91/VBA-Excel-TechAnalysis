Attribute VB_Name = "modKAMA"
'Calculating KAMA (Kaufman Adaptive Moving Average)
'Made by George Slinger github/gslinger91 - International Banking MSc Student
'Insert closing data prices into column A with a Header
' run with calcKAMA
' before: https://i.gyazo.com/8e652e302b381db39b08031987e5867f.png
' after: https://i.gyazo.com/4c6ac4cfba983f1b4bfbb70af0de8994.png

Option Explicit

'Values by Perry Kaufman
Const kPERIOD As Integer = 10 'Period used for moving average
Const kFASTEST As Integer = 2 'Number of periods for the fastest EMA constant
Const kSLOWEST As Integer = 30 'Number of periods for the slowest EMA constant

Public firstrow As Integer, lastrow As Integer

Public wksKAMA As Worksheet

Sub calcKAMA() 'Calculates and fills excel sheet with KAMA computations

Set wksKAMA = ActiveWorkbook.Sheets("KAMA") ' Worksheet is named "KAMA" Adjust accordingly

Dim i As Integer, j As Integer, k As Integer, l As Integer, _
    o As Integer, p As Integer, k2 As Integer  'Declare loop variables
    
With wksKAMA
    firstrow = kPERIOD + 2 'Computations start at end of predefined period (kPERIOD)
    lastrow = .UsedRange.Rows.Count
    
    Dim changeOfDirection As Double 'Calculates the absolute value of change in direction
    
    '= Abs(Closing Price_n - Closing Price_n-10)
    For i = firstrow To lastrow
        changeOfDirection = .Cells(i, 1).Value - .Cells(i - kPERIOD, 1).Value
        .Cells(i, 2).Value = Abs(changeOfDirection)
    Next i
    
    Dim volatility As Double 'Calculates absolute daily difference
    
    .Cells(2, 3).Value = 0 'Sets first variable to 0 as no comparison
    
    'Abs(Close_n - Close_n-1)
    
    For j = firstrow To lastrow + (kPERIOD - 1)
        volatility = .Cells(j - (kPERIOD - 1), 1).Value - .Cells(j - kPERIOD, 1).Value
        .Cells(j - (kPERIOD - 1), 3).Value = Abs(volatility)
    Next j
    
    Dim sumVolatility As Double 'Calculates the sum of n volatilities where: n=kPERIOD
    
    For k = firstrow To lastrow
        sumVolatility = .Cells(k, 3).Value
        For k2 = 1 To kPERIOD - 1
            sumVolatility = sumVolatility + .Cells(k - k2, 3).Value
        
        Next k2
        .Cells(k, 4).Value = sumVolatility
    Next k
    
    Dim efficiencyRatio As Double 'Price change adjusted for daily volatility
    
    'Change of Direction / Volatility
    
    For l = firstrow To lastrow
        efficiencyRatio = .Cells(l, 2).Value / .Cells(l, 4).Value
        .Cells(l, 5).Value = efficiencyRatio
    Next l
    
    
    Dim smoothingConstant As Double ' Smoothing constant uses 2 smoothing constants and ER based on EMA
    
    '(SC=ER*(fastestSC-slowestSC)+slowestSC)^2  where: slowestSC = 2/(kSLOWEST+1)
    
    For o = firstrow To lastrow
        smoothingConstant = (.Cells(o, 5).Value * (2 / (kFASTEST + 1) - 2 / (kSLOWEST + 1)) + 2 / (kSLOWEST + 1)) ^ 2
        .Cells(o, 6).Value = smoothingConstant
    Next o
    
    Dim KAMA As Double  'Calculates the KAMA
    
   .Cells(firstrow, 7).Value = WorksheetFunction.Average(.Range(Cells(firstrow, 1), _
                        Cells(firstrow - (kPERIOD - 1), 1))) 'Sets first KAMA to the SMA as no prior value
    
    'Current KAMA = Prior KAMA + SC * (Closing Price - Prior KAMA)
    
    For p = firstrow + 1 To lastrow
        KAMA = .Cells(p - 1, 7).Value + .Cells(p, 6).Value * (.Cells(p, 1).Value - .Cells(p - 1, 7))
        .Cells(p, 7).Value = KAMA
    Next p

End With

Call formatKAMA

End Sub

Sub formatKAMA()

'Optional formatting

With wksKAMA
    
    'Create headers
    .Cells(1, 1).Value = "Close Price"
    .Cells(1, 2).Value = "Abs Change"
    .Cells(1, 3).Value = "Abs Volatility"
    .Cells(1, 4).Value = "Sum last N Volatility"
    .Cells(1, 5).Value = "Efficiency Ratio"
    .Cells(1, 6).Value = "Smoothing Constant"
    .Cells(1, 7).Value = "KAMA"
    
    With .Range("A1:G1")
        .Font.Bold = True
        .ColumnWidth = 10.29
        .RowHeight = 28.5
        .WrapText = True
        .EntireColumn.HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous 'creates single bottom border for cell
    End With
        
    .Range("B2:B" & kPERIOD + 1).Interior.ColorIndex = 16  'fills the unused cells; 16=grey
    .Range("D2:G" & kPERIOD + 1).Interior.ColorIndex = 16
    
    .Range(Cells(2, 1), Cells(lastrow, 7)).NumberFormat = "0.0000"  '3 decimal places
    

End With

End Sub

Sub drawLineChart()

'to do, need OHLC data

End Sub

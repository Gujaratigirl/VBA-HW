VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockMarketAnalysis():
'Set up variables
Dim Ticker As String
'Dim YearChange As long
Dim PercentChange As Double
Dim TotalStockVolume As Double

'Dim StartYear As Date
'Dim EndYear As Date
Dim LastRow As Variant
Dim i As Variant

Dim MaxCompiled As Double
Dim MinCompiled As Double
Dim LastRowCompiled As Variant

Dim RowLocationResult As Integer

Dim WSCount As Integer
Dim j As Integer

'Set up start positions
RowLocationResult = 2
yearstart = 2

'Set up headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Columns("P").ColumnWidth = 15
'Change alignment to right
Cells(2, 14).HorizontalAlignment = xlRight
Cells(3, 14).HorizontalAlignment = xlRight
Cells(4, 14).HorizontalAlignment = xlRight


'Test Sum function
'ActiveCell.Value = Application.Sum(Range(Cells(1, 7), Cells(4, 7)))

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'Test
'MsgBox LastRow

'To compile Ticker, Change amount, % change,total stock volume
For i = 2 To LastRow
    If Cells(i, 1).Value <> Cells(i + 1, 1) Then
    'Test
    'MsgBox Cells(i, 1).Value
    Cells(RowLocationResult, 9).Value = Cells(i, 1).Value
    'Year end-year start for change
    Cells(RowLocationResult, 10).Value = Cells(i, 6).Value - Cells(yearstart, 3).Value
        If Cells(RowLocationResult, 10) > 0 Then
            Cells(RowLocationResult, 10).Interior.ColorIndex = 4
            Else
            Cells(RowLocationResult, 10).Interior.ColorIndex = 3
        End If
    'year end-year start/year start =Percent change
        If Cells(yearstart, 3).Value <> 0 Then
            Cells(RowLocationResult, 11).Value = (Cells(i, 6).Value - Cells(yearstart, 3).Value) / Cells(yearstart, 3).Value
            Cells(RowLocationResult, 11).NumberFormat = "0.00%"
            'sum startyear -Rowlocationresult =Total stock volume
            Cells(RowLocationResult, 12).Value = Application.Sum(Range(Cells(yearstart, 7), Cells(i, 7)))
            Else
            Cells(RowLocationResult, 11).Value = 0
            Cells(RowLocationResult, 11).NumberFormat = "0.00%"
            'sum startyear -Rowlocationresult =Total stock volume
            Cells(RowLocationResult, 12).Value = Application.Sum(Range(Cells(yearstart, 7), Cells(i, 7)))
    
    End If
    
    
    'Set New postions after change
    RowLocationResult = RowLocationResult + 1
    yearstart = i + 1
    
    End If
Next i
'MsgBox yearstart

'find last row of compiled data
LastRowCompiled = Cells(Rows.Count, 9).End(xlUp).Row
    'MsgBox Lastrowcomplied
'Find max,min, totalstockvolume of complied data
MaxCompiled = Application.WorksheetFunction.Max(Columns("k"))
MinCompiled = Application.WorksheetFunction.Min(Columns("k"))
TotalStockVolume = Application.WorksheetFunction.Max(Columns("L"))

'insert into cell and format to % 2dp
Cells(2, 16).Value = MaxCompiled
Cells(2, 16).NumberFormat = "0.00%"

Cells(3, 16).Value = MinCompiled
Cells(3, 16).NumberFormat = "0.00%"

Cells(4, 16).Value = TotalStockVolume

'find and insert Ticker for MAX
For i = 2 To LastRowCompiled

    If Cells(i, 11) = MaxCompiled Then
    Ticker = Cells(i, 9).Value
    Cells(2, 15).Value = Ticker
    
    End If
Next i
'Find and insert ticker for MIN
'MsgBox Ticker

For i = 2 To LastRowCompiled

If Cells(i, 11) = MinCompiled Then
    Ticker = Cells(i, 9).Value
    Cells(3, 15).Value = Ticker
    
End If
Next i
'
'MsgBox Ticker
'Find and insert ticker for TotalStockVolume
For i = 2 To LastRowCompiled

If Cells(i, 12) = TotalStockVolume Then
    Ticker = Cells(i, 9).Value
    Cells(4, 15).Value = Ticker
    
End If
Next i
'MsgBox Ticker



End Sub
Sub RunInAllWorksheets()
'found on the internet - using to run through all of the 3 tabs
    Dim nWorksheet As Worksheet
    Application.ScreenUpdating = False
    For Each nWorksheet In Worksheets
        nWorksheet.Select
        Call StockMarketAnalysis
    Next
    Application.ScreenUpdating = True
End Sub


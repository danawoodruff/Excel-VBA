Attribute VB_Name = "Module1"

Public Sub UniqueTickerList()

' Add headers for a new table and cull unique ticker symbols from Column A
'
' Add column headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("N2").Value = "Greatest % increase:"
    Range("N3").Value = "Greatest % decrease:"
    Range("N4").Value = "Greatest total volume:"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
'Adjust column widths
    columns("I:J").ColumnWidth = 12.4
    columns("K:K").ColumnWidth = 13
    columns("L:L").ColumnWidth = 16.2
    columns("N:N").ColumnWidth = 18

'Create a list of unique ticker symbols in Column I from the values in Column A
'
Dim lastCell As String

    Range("A1").Select
    Selection.End(xlDown).Select
    lastCell = ActiveCell.Address
    
'Create Column I with unique values
    Range(("A2"), lastCell).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I2"), Unique:=True
    
'Remove duplicate at column top
    Range("I1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Range("$I$1:$L$3170").RemoveDuplicates columns:=1, Header:= _
        xlYes
                         
'
End Sub

Public Sub LoopWSs()

'Adjusts Excel settings for faster VBA processing. Author:https://www.reddit.com/user/ViperSRT3g/
Application.ScreenUpdating = Not Toggle
    Application.EnableEvents = Not Toggle
    Application.DisplayAlerts = Not Toggle
    Application.EnableAnimations = Not Toggle
    Application.DisplayStatusBar = Not Toggle
    Application.PrintCommunication = Not Toggle
    Application.Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
    
'Loop through each worksheet and call each procedure
Dim i As Integer

'Loop through each worksheet
For i = 1 To Worksheets.Count
    Worksheets(i).Select
    
'Call each calculating procedure
    UniqueTickerList
    yearlyChange
    totalStockVolume
    SummaryTable
'Call formatting procedure
    FormatColumns
    Range("M1").Select
Next i

cmdAnalyzer.Hide

End Sub

Public Sub yearlyChange()

'Yearly and percentage changes from opening price to the closing price.

Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim ticker As String
Dim numberTickers As Integer
Dim lastCell As Double

openingPrice = 0
yearlyChange = 0
percentChange = 0
ticker = ""
numberTickers = 0

' Last row of each worksheet
    lastCell = Cells(Rows.Count, "A").End(xlUp).Row
 
' Loop through the list of tickers.
    For i = 2 To lastCell

' Unique ticker symbol being calculating for.
    ticker = Cells(i, 1).Value
        
' Opening price for the ticker.
    If openingPrice = 0 Then
        openingPrice = Cells(i, 3).Value
    End If
                
' Run for the next ticker symbol.
    If Cells(i + 1, 1).Value <> ticker Then
        ' Change ticker in the list.
        numberTickers = numberTickers + 1
    Cells(numberTickers + 1, 9) = ticker
            
' Closing price for ticker
    closingPrice = Cells(i, 6)
            
' Calculate yearly change value
    yearlyChange = closingPrice - openingPrice
            
' Print yearly change value to Column J
    Cells(numberTickers + 1, 10).Value = yearlyChange
            
' Calculate percentChange
    If yearlyChange = 0 Or openingPrice = 0 Then
        percentChange = 0
        Else
            percentChange = Round(yearlyChange / openingPrice, 4)
    End If
'
' Print percentage change value to Column K
    Cells(numberTickers + 1, 11).Value = percentChange
'
' Color format yearlyChange values
    If yearlyChange > 0 Then
        Cells(numberTickers + 1, 10).Interior.ColorIndex = 4
        ElseIf yearlyChange < 0 Then
            Cells(numberTickers + 1, 10).Interior.ColorIndex = 3
            Else
                Cells(numberTickers + 1, 10).Interior.ColorIndex = 6
    End If
'
' Set opening price back to 0 for next ticker.
    openingPrice = 0
' Set yearly change back to 0 for next ticker.
    yearlyChange = 0
' Set percentChange back to 0 for next ticker.
    percentChange = 0
'
    
End If
'
    Next i
'
'Remove duplicate at column end
    Range("I1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Range("$I$1:$L$3170").RemoveDuplicates columns:=1, Header:= _
        xlYes
'
End Sub

Public Sub totalStockVolume()

Dim ticker As String
Dim numberTickers As Integer
Dim lastCell As Double
Dim totalStockVolume As Double

ticker = ""
numberTickers = 0
totalStockVolume = 0

' Last row of each worksheet
    lastCell = Cells(Rows.Count, "A").End(xlUp).Row
 
' Loop through the list of tickers.
    For i = 2 To lastCell

' Unique ticker symbol being calculating for.
    ticker = Cells(i, 1).Value
                
' Sum stock volume values unique ticker.
    totalStockVolume = totalStockVolume + Cells(i, 7).Value
        
' Run for the next ticker symbol.
    If Cells(i + 1, 1).Value <> ticker Then
    ' Change ticker in the list.
    numberTickers = numberTickers + 1
    Cells(numberTickers + 1, 9) = ticker

' Display total stock volume value to Column L
    Cells(numberTickers + 1, 12).Value = totalStockVolume
            
' Set total stock volume back to 0 for next ticker.
    totalStockVolume = 0

End If

    Next i
    
End Sub

Public Sub FormatColumns()
'
'Center data in columns
    columns("I:P").HorizontalAlignment = xlCenter
    columns("I:P").VerticalAlignment = xlBottom
'
' Format column J as Currency
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Currency"
'
' Format column K as percentage with two decimal places.
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"

' Add commas for readability in column L
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

End Sub

Public Sub SummaryTable()

'Bonus section calculations: return greatest increase and decrease in percentages and the greatest total stock volume.

'Filter Columns K and L for largest Max and Min values
        
    MaxValue = Application.WorksheetFunction.Max(Range("K:K"))
    Range("P2").Value = MaxValue

    MinValue = Application.WorksheetFunction.Min(Range("K:K"))
    Range("P3").Value = MinValue
       
    MaxValue2 = Application.WorksheetFunction.Max(Range("L:L"))
    Range("P4").Value = MaxValue2
    
'Retrieve tickers for Max and Min values
    
    Range("O2") = WorksheetFunction.xlookup(MaxValue, [K:K], [I:I], "Not found")
    Range("O3") = WorksheetFunction.xlookup(MinValue, [K:K], [I:I], "Not found")
    Range("O4") = WorksheetFunction.xlookup(MaxValue2, [L:L], [I:I], "Not found")
    
'Format Cells
    Range("P2:P3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    Range("P4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);"

End Sub

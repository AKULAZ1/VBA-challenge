Attribute VB_Name = "Module1"
Sub StockInfo()

Dim WS As Worksheet

' This "Workbook.Worksheets" was sourced from the following link- https://powerspreadsheets.com/excel-vba-sheets-worksheets/

For Each WS In ThisWorkbook.Worksheets

' To set up Summary Table

WS.Range("I1").Value = "Ticker"
WS.Range("J1").Value = "Yearly Change"
WS.Range("K1").Value = "Percent Change"
WS.Range("L1").Value = "Total Stock Volume"
WS.Columns("I:L").AutoFit

' To populate the Summary Table

lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim Ticker As String
Dim SummaryRow As Integer
Dim AnChange As Double
Dim OpenRow As Long
Dim AnChangeP As Double
Dim TotVol As LongLong

SummaryRow = 2
OpenRow = 2

For Row = 2 To lastRow

    If WS.Cells(Row + 1, 1).Value <> WS.Cells(Row, 1).Value Then
        Ticker = WS.Cells(Row, 1).Value
        WS.Cells(SummaryRow, 9).Value = Ticker
        
        AnChange = WS.Cells(Row, 6).Value - WS.Cells(OpenRow, 3).Value
        WS.Cells(SummaryRow, 10).Value = AnChange
        
        'Conditional Formatting
        If WS.Cells(SummaryRow, 10).Value > 0 Then
            WS.Cells(SummaryRow, 10).Interior.ColorIndex = 4
        Else
            WS.Cells(SummaryRow, 10).Interior.ColorIndex = 3
        End If
        
        AnChangeP = AnChange / WS.Cells(OpenRow, 3).Value
        WS.Cells(SummaryRow, 11).Value = AnChangeP
        'The piece of code to format the Percent Change was sourced from a StackOverflow Question Board
        WS.Cells(SummaryRow, 11).NumberFormat = "0.00%"
        
        'Conditional Formatting
         If WS.Cells(SummaryRow, 11).Value > 0 Then
            WS.Cells(SummaryRow, 11).Interior.ColorIndex = 4
        Else
            WS.Cells(SummaryRow, 11).Interior.ColorIndex = 3
        End If
            
        TotVol = TotVol + WS.Cells(Row, 7).Value
        WS.Cells(SummaryRow, 12).Value = TotVol
        
        SummaryRow = SummaryRow + 1
        OpenRow = Row + 1
        TotVol = 0
    Else
        TotVol = TotVol + WS.Cells(Row, 7).Value
    End If
    
Next Row
    
' To set up the Statistics Summary Table

WS.Range("P1").Value = "Ticker"
WS.Range("Q1").Value = "Value"
WS.Range("O2").Value = "Greatest % Increase"
WS.Range("O3").Value = "Greatest % Decrease"
WS.Range("O4").Value = "Greatest Total Volume"
WS.Columns("O:R").AutoFit

' To populate the Statistics Summary Table

'To find the Greatest % Increase
Dim MaxIncrease As Double
Dim TickerIndex As Integer

MaxIncrease = WorksheetFunction.Max(WS.Range("K:K"))
WS.Range("Q2").Value = MaxIncrease
TickerIndex = WorksheetFunction.Match(MaxIncrease, WS.Range("K:K"), 0)
WS.Range("P2").Value = WS.Range("I" & TickerIndex).Value
'The piece of code to format the Percent Change was sourced from a StackOverflow Question Board
WS.Range("Q2").NumberFormat = "0.00%"
        

'To find the Greatest % Decrease
Dim MaxDecrease As Double
Dim TickerIndex2 As Integer

MaxDecrease = WorksheetFunction.Min(WS.Range("K:K"))
WS.Range("Q3").Value = MaxDecrease
TickerIndex2 = WorksheetFunction.Match(MaxDecrease, WS.Range("K:K"), 0)
WS.Range("P3").Value = WS.Range("I" & TickerIndex2).Value
'The piece of code to format the Percent Change was sourced from a StackOverflow Question Board
WS.Range("Q3").NumberFormat = "0.00%"
        

'To find the Greatest Total Volume
Dim MaxVolume As LongLong
Dim TickerIndex3 As Integer

MaxVolume = WorksheetFunction.Max(WS.Range("L:L"))
WS.Range("Q4").Value = MaxVolume
TickerIndex3 = WorksheetFunction.Match(MaxVolume, WS.Range("L:L"), 0)
WS.Range("P4").Value = WS.Range("I" & TickerIndex3).Value

Next WS

End Sub

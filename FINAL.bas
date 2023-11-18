Attribute VB_Name = "Module2"

' Function to check if an item exists in a collection
Function Contains(col As Collection, key As Variant) As Boolean
    On Error Resume Next
    Contains = Not (col(key) Is Nothing)
    On Error GoTo 0
End Function



Sub CalculateStockDataWithFormatting()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Counter As Long
    Dim TickerSymbol As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalVolume As Double
    Dim YearStartDate As Date
    Dim YearEndDate As Date
    Dim TickerDates As New Collection
    Dim TickerDate As Variant
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    ' Initialize variables for greatest values
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row in column A
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables
        Counter = 2
        TotalVolume = 0
        Set TickerDates = New Collection
        
        ' Loop through rows to calculate stock data
        i = 2
        Do While i <= LastRow
            If ws.Cells(i, 1).Value <> "" Then
                ' Get the Ticker Symbol
                TickerSymbol = ws.Cells(i, 1).Value
                
                ' Get the Date
                CurrentDate = ws.Cells(i, 2).Value
                
                ' Check if it's a new year
                If Not Contains(TickerDates, TickerSymbol) Then
                    TickerDates.Add TickerSymbol, CStr(CurrentDate)
                    OpeningPrice = ws.Cells(i, 3).Value  ' Initial opening price for the year
                End If
                
                ' Get the Closing Price (assuming it's in column F)
                ClosingPrice = ws.Cells(i, 6).Value
                
                ' Check if it's the last row for the current ticker or the last row in the sheet
                If i = LastRow Or (ws.Cells(i + 1, 1).Value <> TickerSymbol And ws.Cells(i, 1).Value = TickerSymbol) Then
                    ' Calculate Yearly Change
                    YearlyChange = ClosingPrice - OpeningPrice
                    
                    ' Calculate Percentage Change
                    If OpeningPrice <> 0 Then
                        PercentageChange = (YearlyChange / OpeningPrice) * 100
                    Else
                        PercentageChange = 0
                    End If
                    
                    ' Accumulate Total Volume
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value  ' Assuming Volume is in column G
                    
                    ' Output results to columns I, J, K, L, and M
                    ws.Cells(Counter, 9).Value = TickerSymbol
                    ws.Cells(Counter, 10).Value = YearlyChange
                    ws.Cells(Counter, 11).Value = PercentageChange
                    ws.Cells(Counter, 12).Value = TotalVolume
                    
                    ' Check for greatest values
                    If PercentageChange > GreatestIncrease Then
                        GreatestIncrease = PercentageChange
                        GreatestIncreaseTicker = TickerSymbol
                    End If
                    
                    If PercentageChange < GreatestDecrease Then
                        GreatestDecrease = PercentageChange
                        GreatestDecreaseTicker = TickerSymbol
                    End If
                    
                    If TotalVolume > GreatestVolume Then
                        GreatestVolume = TotalVolume
                        GreatestVolumeTicker = TickerSymbol
                    End If
                    
                    ' Reset variables for the next ticker
                    Counter = Counter + 1
                    OpeningPrice = 0
                    TotalVolume = 0
                    Set TickerDates = New Collection
                End If
            End If
            i = i + 1
        Loop
        
        ' Apply conditional formatting to "Yearly Change" column (Column J)
        ws.Range(ws.Cells(2, 10), ws.Cells(Counter - 1, 10)).FormatConditions.Delete
        ws.Range(ws.Cells(2, 10), ws.Cells(Counter - 1, 10)).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        ws.Range(ws.Cells(2, 10), ws.Cells(Counter - 1, 10)).FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green
        
        ws.Range(ws.Cells(2, 10), ws.Cells(Counter - 1, 10)).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        ws.Range(ws.Cells(2, 10), ws.Cells(Counter - 1, 10)).FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red
    
            ' Output greatest values to separate columns
        ws.Cells(1, 14).Value = "Greatest % Increase"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(2, 14).Value = GreatestIncrease
        ws.Cells(2, 15).Value = GreatestIncreaseTicker
        
        ws.Cells(4, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Ticker"
        ws.Cells(5, 14).Value = GreatestDecrease
        ws.Cells(5, 15).Value = GreatestDecreaseTicker
        
        ws.Cells(7, 14).Value = "Greatest Total Volume"
        ws.Cells(7, 15).Value = "Ticker"
        ws.Cells(8, 14).Value = GreatestVolume
        ws.Cells(8, 15).Value = GreatestVolumeTicker
    
    Next ws
    

End Sub

Attribute VB_Name = "Module1"
Sub CalculateStockData()
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
                    
                    ' Reset variables for the next ticker
                    Counter = Counter + 1
                    OpeningPrice = 0
                    TotalVolume = 0
                    Set TickerDates = New Collection
                End If
            End If
            i = i + 1
        Loop
    Next ws
End Sub

Function Contains(coll As Collection, key As Variant) As Boolean
    On Error Resume Next
    Contains = Not IsEmpty(coll(key))
    On Error GoTo 0
End Function

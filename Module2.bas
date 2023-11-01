Attribute VB_Name = "Module2"
Sub ResetWorksheetsButton()

    ' Iterate over all worksheets
    For Each WS In ThisWorkbook.Worksheets
        ' Clear the contents of the range I:P
        WS.Range("I:P").ClearContents

        ' Set the interior color of the range I:P to none
        WS.Range("I:P").Interior.ColorIndex = xlNone
    Next WS

    ' Activate the worksheet "2018"
    ThisWorkbook.Sheets("A").Activate

End Sub

'Button to complete the entire workbook
Sub AllStockWSButton()

    ' Enable screen updating
    Application.ScreenUpdating = True

    ' Iterate over all worksheets
    For Each WS In Worksheets
        ' Activate the worksheet
        WS.Activate

        ' Autofit all columns in the range A:Q
        WS.Range("A:Q").EntireColumn.AutoFit

        ' Set the number format of the range P4, G:G, L:L to #,##0
        WS.Range("P4, G:G, L:L").NumberFormat = "#,##0"

        ' Call the StockWS procedure
        Call StockWS
    Next WS

    ' Activate the worksheet "2018"
    ThisWorkbook.Sheets("A").Activate

End Sub

' Process variables between sheets
Sub StockWS()

    ' Set the range I1:L1 to the array ["Ticker", "Yearly Change", "Percent Change", "Total Stock Volume"]
    Range("I1:L1") = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

    ' Set the range O1:P1 to the array ["Ticker", "Value"]
    Range("O1:P1") = Array("Ticker", "Value")

    ' Set the range N2:N4 to the transposed array ["Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"]
    Range("N2:N4") = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))

    ' Set the number format of the range J:J to 0.00
    Range("J:J").NumberFormat = "0.00"

    ' Set the number format of the range K:K, P2:P3 to 0.00%
    Range("K:K, P2:P3").NumberFormat = "0.00%"

    ' Declare variables
    Dim ticker As String
    Dim openPrice, yearChanging, percentUpdate, volumeTotal As Double
    Dim increasedTicker, decreasedTicker, volTicker As String
    Dim bestIncrease, bestDecrease, bestVol As Double
    Dim inputRow As Long
    Dim outputRow As Integer

    ' Initialize the input row
    inputRow = 3

    ' Initialize the output row
    outputRow = 2

    ' Loop while the ticker is not empty
    Do While (ticker <> "")

        ' Get the ticker from the current input row
        ticker = Range(inputRow, 1)

        ' Initialize the open price, year changing, percent update, and volume total
        openPrice = Range(inputRow, 3)
        yearChanging = 0
        percentUpdate = 0
        volumeTotal = 0

        ' Loop while the ticker is the same on the current input row
        Do While (ticker = Range(inputRow, 1))

            ' Add the volume to the volume total
            volumeTotal = volumeTotal + Range(inputRow, 7)

            ' Increment the input row
            inputRow = inputRow + 1

        Loop

        ' Calculate the year changing and percent update
        yearChanging = Range(inputRow - 1, 6) - openPrice
        percentUpdate = yearChanging / openPrice

        ' Set the ticker, year changing, percent update, and volume total for the current output row
        Range(outputRow, 9) = ticker
        Range(outputRow, 10) = yearChanging

        ' Set the interior color of the range for the year changing to red if it is negative, or to green if it is positive
        If yearChanging < 0 Then
            Range(outputRow, 10).Interior.Color = vbRed
        ElseIf yearChanging > 0 Then
            Range(outputRow, 10).Interior.Color = vbGreen
        End If

        Range(outputRow, 11) = percentUpdate
    Loop
        ' Update the best increase, best decrease, and best volume
End Sub


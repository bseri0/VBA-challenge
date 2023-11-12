Attribute VB_Name = "Module3"
Sub AnalyzeTicker()

    ' Declare variables
    Dim ticker As String
    Dim earliestOpenPrice, latestClosePrice As Double
    Dim returnn As Double
    ' Create a dictionary to store the earliest open price and latest close price for each ticker
    Dim tickerData As New Dictionary
    
    Set tickerData = New Dictionary

    ' Iterate over all rows in the worksheet
    For Each Row In ActiveSheet.Rows

        ' Get the ticker and date from the current row
        ticker = Row.Cells(1).Value
        Datee = Row.Cells(2).Value

        ' If the ticker is not in the dictionary, add it with an empty earliest open price and latest close price
        If Not tickerData.Exists(ticker) Then
            tickerData.Add ticker, Array(0, 0)
        End If

        ' Get the earliest open price and latest close price for the current ticker from the dictionary
        earliestOpenPrice = tickerData(ticker)(0)
        latestClosePrice = tickerData(ticker)(1)

        ' If the current date is earlier than the earliest open price date, update the earliest open price
        If Datee < earliestOpenPrice Then
            tickerData(ticker)(0) = Date
        End If

        ' If the current date is later than the latest close price date, update the latest close price
        If Date > latestClosePrice Then
            tickerData(ticker)(1) = Date
        End If

    Next Row

    ' Iterate over the ticker data dictionary
    For Each tickerDataEntry In tickerData

        ' Get the ticker, earliest open price, and latest close price
        ticker = tickerDataEntry.Key
        earliestOpenPrice = tickerDataEntry.Value(0)
        latestClosePrice = tickerDataEntry.Value(1)

        ' Calculate the return for the ticker
        
        returnn = (latestClosePrice - earliestOpenPrice) / earliestOpenPrice * 100

        ' Print the ticker, earliest open price, latest close price, and return to the console
        Debug.Print ticker, earliestOpenPrice, latestClosePrice, returnn

    Next tickerDataEntry

End Sub

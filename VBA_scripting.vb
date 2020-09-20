Sub VBAmacro():

'Declarations-work on uniform naming convention for future-not fond of these
Dim ticker As String
Dim tickerCount As Integer
Dim lastRow As Long
Dim openPrice As Double
Dim closePrice As Double
Dim annualChange As Double
Dim percentChange As Double
Dim totalVolume As Double


'Loop through worksheets--reminder:make sure you unfilter the spreadsheet :)
For Each ws In Worksheets

    ws.Activate

    'Identify last row
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'Display header columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Set variables
    tickerCount = 0
    ticker = ""
    annualChange = 0
    openPrice = 0
    percentChange = 0
    totalVolume = 0

    For i = 2 To lastRow

        ticker = Cells(i, 1).Value

        If openPrice = 0 Then
            openPrice = Cells(i, 3).Value
        End If

        totalVolume = totalVolume + Cells(i, 7).Value

        If Cells(i + 1, 1).Value <> ticker Then
            tickerCount = tickerCount + 1
            Cells(tickerCount + 1, 9) = ticker

            closePrice = Cells(i, 6)

            annualChange = closePrice - openPrice

           'Conditional formatting
            Cells(tickerCount + 1, 10).Value = annualChange
            
            'Format positive change in green
            If annualChange > 0 Then
                Cells(tickerCount + 1, 10).Interior.ColorIndex = 4
            'Format negative change in red
            ElseIf annualChange < 0 Then
                Cells(tickerCount + 1, 10).Interior.ColorIndex = 3
            End If

            'Need to read up more on Else/ElseIf, why Else here and ElseIf next set
            If openPrice = 0 Then
                percentChange = 0
            Else
                percentChange = (annualChange / openPrice)
            End If

            'Conditional formatting
            Cells(tickerCount + 1, 11).Value = Format(percentChange, "Percent")

            'Format positive change in green
            If percentChange > 0 Then
                Cells(tickerCount + 1, 11).Interior.ColorIndex = 4
            'Format negative change in red
            ElseIf percentChange < 0 Then
                Cells(tickerCount + 1, 11).Interior.ColorIndex = 3
            End If

            'New ticker, reset openPrice
            openPrice = 0

            'Display totalVolume for each ticker.
            Cells(tickerCount + 1, 12).Value = totalVolume

            'New ticker, reset totalVolume
            totalVolume = 0
        End If

    Next i

'Challenge Declarations
Dim mostpercentIncrease As Double
Dim mpiTicker As String
Dim mostpercentDecrease As Double
Dim mpdTicker As String
Dim highestVolume As Double
Dim hvTicker As String

    'Display challenge information
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    'Identify last row
    lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row

    'Initialize and set challenge variables-work on uniform naming convention for future-not fond of these
    mpiTicker = Cells(2, 9).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    mostpercentIncrease = Cells(2, 11).Value
    mostpercentDecrease = Cells(2, 11).Value
    highestVolume = Cells(2, 12).Value


    'Loop through tickers.
    For i = 2 To lastRow

        'Find Greatest % Increase
        If Cells(i, 11).Value > mostpercentIncrease Then
            mostpercentIncrease = Cells(i, 11).Value
            mpiTicker = Cells(i, 9).Value
        End If

        'Find Greatest % Decrease
        If Cells(i, 11).Value < mostpercentDecrease Then
            mostpercentDecrease = Cells(i, 11).Value
            mpdTicker = Cells(i, 9).Value
        End If

        'Find Greatest Total Volume
        If Cells(i, 12).Value > highestVolume Then
            highestVolume = Cells(i, 12).Value
            hvTicker = Cells(i, 9).Value
        End If

    Next i

    'Display/format challenge values
    Range("P2").Value = mpiTicker
    Range("P3").Value = mpdTicker
    Range("P4").Value = hvTicker
    Range("Q2").Value = Format(mostpercentIncrease, "Percent")
    Range("Q3").Value = Format(mostpercentDecrease, "Percent")
    Range("Q4").Value = highestVolume
    
    'Autofit columns-thank you google search stackoverflow
    Columns("A:Q").Select

    Selection.EntireColumn.AutoFit

Next ws

End Sub

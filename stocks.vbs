Sub stocks():
    ' Declare variables
    Dim ws As Worksheet
    Dim totalVolume As Currency
    Dim ticker As String
    Dim tickerCount, lastRow As Long
    Dim i, j As Long

    Dim openingPrice, closingPrice, yearlyChange, percentChange As Double

    Dim greatIncTicker, greatDecTicker, greatTotVolTicker As String
    Dim greatInc, greatDec As Double
    Dim greatTotVol As Currency

    For Each ws In Worksheets
        ' initialize variables
        lastRow = ws.UsedRange.Rows.Count
        tickerCount = 0

        ' create headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"

        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        ' algorithm assumes data is grouped by ticker, then sorted by date
        For i = 2 To lastRow + 1

            ' compute running total of totalVolume if no ticker change
            If ws.Cells(i, 1) = ticker Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value

            ' toggles every time ticker changes
            Else
                ' makes sure at least one ticker has been looped through
                If tickerCount > 0 Then
                    ' record information after finished looping through ticker
                    ws.Cells(tickerCount + 1, 9).Value = ticker
                    ws.Cells(tickerCount + 1, 12).Value = totalVolume
                    ws.Cells(tickerCount + 1, 12).Style = "Currency"

                    closingPrice = ws.Cells(i - 1, 6).Value
                    yearlyChange = closingPrice - openingPrice
                    percentChange = yearlyChange / openingPrice

                    ws.Cells(tickerCount + 1, 10).Value = yearlyChange
                    ws.Cells(tickerCount + 1, 10).NumberFormat = "0.00000"
                    ws.Cells(tickerCount + 1, 11).Value = percentChange
                    ws.Cells(tickerCount + 1, 11).NumberFormat = "0.00%"

                    ' if yearlyChange is negative: color the cell red; else color it green
                    ' (if there is zero change, I say that is not a bad thing so I color it green)
                    If ws.Cells(tickerCount + 1, 10).Value < 0 Then
                        ws.Cells(tickerCount + 1, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Cells(tickerCount + 1, 10).Interior.Color = RGB(0, 255, 0)
                    End If

                    ' checks if current percentChange is larger or smaller than last stored percentChange
                    ' if equal, do nothing
                    If percentChange > greatInc Then
                        greatInc = percentChange
                        greatIncTicker = ticker
                    ElseIf percentChange < greatDec Then
                        greatDec = percentChange
                        greatDecTicker = ticker
                    End If

                    ' checks if current totalVolume is greater than last stored totalVolume
                    If totalVolume > greatTotVol Then
                        greatTotVol = totalVolume
                        greatTotVolTicker = ticker
                    End If
                
                End If

                ' if ticker changes, re-initialize totalVolume, update ticker and tickerCount, and store new openingPrice
                totalVolume = ws.Cells(i, 7).Value
                ticker = ws.Cells(i, 1).Value
                tickerCount = tickerCount + 1

                ' if openingPrice = 0 on 01/01, set openingPrice to first non-zero value
                ' prevents division by zero when computing percentChange
                If ws.Cells(i, 3).Value = 0 And i < lastRow + 1 Then
                    j = i + 1
                    Do While ws.Cells(j, 3).Value = 0
                        j = j + 1
                    Loop
                    openingPrice = ws.Cells(j, 3).Value
                Else
                    openingPrice = ws.Cells(i, 3).Value
                End If

            End If

        Next i

        ' record greatInc, greatDec, and greatTotVol
        ws.Cells(2, 16) = greatIncTicker
        ws.Cells(2, 17) = greatInc
        ws.Cells(2, 17).NumberFormat = "0.00%"

        ws.Cells(3, 16) = greatDecTicker
        ws.Cells(3, 17) = greatDec
        ws.Cells(3, 17).NumberFormat = "0.00%"

        ws.Cells(4, 16) = greatTotVolTicker
        ws.Cells(4, 17) = greatTotVol
        ws.Cells(4, 17).Style = "Currency"

    Next ws

End Sub

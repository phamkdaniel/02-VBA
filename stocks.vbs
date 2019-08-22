Sub stocks():
    ' Declare variables
    Dim totalVolume As Currency
    Dim ticker As String
    Dim tickerCount As Integer
    Dim lastRow As Long

    Dim openingPrice, closingPrice, yearlyChange, percantChange As Double

    Dim greatInc, greatDec, greatTotVol As Variant


    ' initialize variables
    lastRow = ActiveSheet.UsedRange.Rows.Count
    tickerCount = 0
    greatInc = Array("", 0)
    greatDec = Array("", 0)
    greatTotVol = Array("", 0)


    ' create headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"


    ' algorithm assumes data is sorted by ticker, then by date
    For i = 2 To lastRow + 1
        ' toggles every time ticker changes
        If Cells(i, 1) <> ticker Then

            ' checks if ticker is not initial value to avoid overwriting header
            If tickerCount <> 0 Then
                Cells(tickerCount + 1, 9).Value = ticker
                Cells(tickerCount + 1, 12).Value = totalVolume
                closingPrice = Cells(i - 1, 6).Value

                ' prevents initial check when closingPrice = 0
                If closingPrice <> 0 Then

                    yearlyChange = closingPrice - openingPrice

                    'checks if openingPrice = 0 to prevent division by 0
                    If openingPrice = 0 Then
                        percentChange = 0
                    else
                        percentChange = yearlyChange / openingPrice
                    End If

                    Cells(tickerCount + 1, 10).Value = yearlyChange
                    Cells(tickerCount + 1, 11).Value = percentChange
                
                    ' if yearlyChange is negative: color the cell red; else color it green
                    ' (if there is zero change, I say that is not a bad thing so I colored it green)
                    If Cells(tickerCount + 1, 10).Value < 0 Then
                        Cells(tickerCount + 1, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                        Cells(tickerCount + 1, 10).Interior.Color = RGB(0, 255, 0)
                    End If

                    ' checks if current percentChange is larger or smaller than  
                    ' last stored percentChange and updates greatInc or greatDec
                    If percentChange > greatInc(1) Then
                        greatInc(1) = percentChange
                        greatInc(0) = ticker
                    ElseIf percentChange < greatDec(1) Then
                        greatDec(1) = percentChange
                        greatDec(0) = ticker
                    End If

                    ' checks if current totalVolume is greater than last stored totalVolume
                    If totalVolume > greatTotVol(1) Then
                        greatTotVol(1) = totalVolume
                        greatTotVol(0) = ticker
                    End If

                End If
            
            End If

            ' if ticker changes, re-initialize totalVolume, update ticker and ticker row, and store openingPrice
            totalVolume = Cells(i, 7).Value
            ticker = Cells(i, 1).Value
            tickerCount = tickerCount + 1
            openingPrice = Cells(i, 3).value
        Else
            ' computes running total of totalVolums if no ticker change
            totalVolume = totalVolume + Cells(i, 7).Value
        End If

    Next i


    ' populates greatInc, greatDec, and greatTotVol
    Range("P2").Value = greatInc(0)
    Range("Q2").Value = greatInc(1)
    Range("P3").Value = greatDec(0)
    Range("Q3").Value = greatDec(1)
    Range("P4").Value = greatTotVol(0)
    Range("Q4").Value = greatTotVol(1)

End Sub

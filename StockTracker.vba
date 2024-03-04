Sub StockTracker():

    ' Loop variables
    Dim wsCount As Integer
    Dim i As Integer
    Dim lastRow As LongLong

    ' Data variables
    Dim currentStock As Integer
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim volume As LongLong
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim greatestChangeTicker As String
    Dim smallestChangeTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestPercentChange As Double
    Dim smallestPercentChange As Double
    Dim greatestTotalVolume As Double

    ' Initial values
    currentStock = 2
    volume = 0
    greatestPercentChange = 0
    smallestPercentChange = 0
    greatestTotalVolume = 0

    ' Set WS_Count equal to the number of worksheets in the active workbook.
    ' Source: https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
    wsCount = ActiveWorkbook.Worksheets.Count

    For i = 1 To wsCount
        ' Set some of our variables to their default values
        volume = 0
        currentStock = 2
        greatestChangeTicker = ""
        greatestVolumeTicker = ""
        smallestChangeTicker = ""
        greatestPercentChange = 0
        smallestPercentChange = 0
        greatestTotalVolume = 0

        ' Sets the headers for each column/row we want
        Worksheets(i).Range("I1").Value = "Ticker"
        Worksheets(i).Range("J1").Value = "Yearly Change"
        Worksheets(i).Range("K1").Value = "Percent Change"
        Worksheets(i).Range("L1").Value = "Total Stock Volume"
        Worksheets(i).Range("P1").Value = "Ticker"
        Worksheets(i).Range("Q1").Value = "Value"
        Worksheets(i).Range("O2").Value = "Greatest % Increase"
        Worksheets(i).Range("O3").Value = "Greatest % Decrease"
        Worksheets(i).Range("O4").Value = "Greatest Total Volume"

        ' Find the number of the last row in the given worksheet (source: Anuj via bootcamp Slack)
        lastRow = Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row

        For j = 2 To lastRow
            ' Record the opening price for the first stock
            If j = 2 Then
                openingPrice = Worksheets(i).Cells(j, 3).Value
            End If

            ' Add a day's volume to the total
            volume = volume + Worksheets(i).Cells(j, 7).Value

            ' Check if this is the last row of a given stock
            If Worksheets(i).Cells(j, 1).Value <> Worksheets(i).Cells(j+1, 1).Value Then
                ' Record the closing price
                closingPrice = Worksheets(i).Cells(j, 6).Value

                ' Calculate yearly change and percent change
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice

                'Check yearly change against the greatest and smallest values encountered so far
                If percentChange > greatestPercentChange Then
                    greatestChangeTicker = Worksheets(i).Cells(j, 1).Value
                    greatestPercentChange = percentChange
                ElseIf percentChange < smallestPercentChange Then
                    smallestChangeTicker = Worksheets(i).Cells(j, 1).Value
                    smallestPercentChange = percentChange
                End If

                ' Do the same for volume
                If volume > greatestTotalVolume Then
                    greatestVolumeTicker = Worksheets(i).Cells(j, 1).Value
                    greatestTotalVolume = volume
                End If

                ' Print our individual data to the worksheet
                Worksheets(i).Cells(currentStock, 9).Value = Worksheets(i).Cells(j, 1).Value
                Worksheets(i).Cells(currentStock, 10).Value = yearlyChange
                Worksheets(i).Cells(currentStock, 11).Value = percentChange
                Worksheets(i).Cells(currentStock, 12).Value = volume

                'Do some cleanup
                currentStock = currentStock + 1
                openingPrice = Worksheets(i).Cells(j + 1, 3).Value
                volume = 0
            End If
        Next j

        ' Print our overall data to the worksheet
        Worksheets(i).Range("P2").Value = greatestChangeTicker
        Worksheets(i).Range("Q2").Value = greatestPercentChange
        Worksheets(i).Range("P3").Value = smallestChangeTicker
        Worksheets(i).Range("Q3").Value = smallestPercentChange
        Worksheets(i).Range("P4").Value = greatestVolumeTicker
        Worksheets(i).Range("Q4").Value = greatestTotalVolume
    Next i
End Sub

Sub Challenge2()

    'Set variables
    Dim WS As Worksheet

    ' Loop through all worksheets
    For Each WS In ActiveWorkbook.Worksheets
	
	' Find the last row in the worksheet
    Dim LastRow As Long
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
		
    ' Set the summary table variables
	Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim SummaryRow As Long

        OpenPrice = WS.Cells(2, 3).Value
        TotalStockVolume = 0
        SummaryRow = 2

        ' Loop through the rows in the worksheet
        For i = 2 To LastRow
            ' Check the ticker
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                ' Store the ticker
                Ticker = WS.Cells(i, 1).Value

                ' add the ticker total
                ClosePrice = WS.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If

                ' Create summary table
                WS.Cells(SummaryRow, "I").Value = Ticker
                WS.Cells(SummaryRow, "J").Value = YearlyChange
                WS.Cells(SummaryRow, "K").Value = PercentChange
                WS.Cells(SummaryRow, "K").NumberFormat = "0.00%"
                WS.Cells(SummaryRow, "L").Value = TotalStockVolume

                ' Print titles to the summary table
                WS.Cells(1, "I").Value = "Ticker"
                WS.Cells(1, "J").Value = "Yearly Change"
                WS.Cells(1, "K").Value = "Percent Change"
                WS.Cells(1, "L").Value = "Total Stock Volume"

                ' Reset the ticker total
                OpenPrice = WS.Cells(i + 1, 3).Value
                TotalStockVolume = 0
                SummaryRow = SummaryRow + 1
            Else
                ' Save the ticker value
                TotalStockVolume = TotalStockVolume + WS.Cells(i, 7).Value
            End If
        Next i

        ' Find the last row of the ticker column
        TickerLastRow = WS.Cells(Rows.Count, "I").End(xlUp).Row
		
    	 ' Set the conditional formating for the cells
	Dim j As Long
        For j = 2 To TickerLastRow
            If WS.Cells(j, "J").Value >= 0 Then
                WS.Cells(j, "J").Interior.ColorIndex = 10 ' Green
            Else
                WS.Cells(j, "J").Interior.ColorIndex = 3 ' Red
            End If
        Next j
		
	'create the statistics table
	Dim t As Long
        WS.Cells(2, "O").Value = "Greatest % Increase"
        WS.Cells(3, "O").Value = "Greatest % Decrease"
        WS.Cells(4, "O").Value = "Greatest Total Volume"
        WS.Cells(1, "P").Value = "Ticker"
        WS.Cells(1, "Q").Value = "Value"
		
		' Find the biggest values of the tickers
        Dim MaxPercentChange As Double
        Dim MinPercentChange As Double
        Dim MaxTotalStockVolume As Double
        Dim MaxPercentChangeTicker As String
        Dim MinPercentChangeTicker As String
        Dim MaxTotalStockVolumeTicker As String
        
        MaxPercentChange = Application.WorksheetFunction.Max(WS.Range("K2:K" & TickerLastRow))
        MinPercentChange = Application.WorksheetFunction.Min(WS.Range("K2:K" & TickerLastRow))
        MaxTotalStockVolume = Application.WorksheetFunction.Max(WS.Range("L2:L" & TickerLastRow))

        For t = 2 To TickerLastRow
            If WS.Cells(t, "K").Value = MaxPercentChange Then
                MaxPercentChangeTicker = WS.Cells(t, "I").Value
            ElseIf WS.Cells(t, "K").Value = MinPercentChange Then
                MinPercentChangeTicker = WS.Cells(t, "I").Value
            ElseIf WS.Cells(t, "L").Value = MaxTotalStockVolume Then
                MaxTotalStockVolumeTicker = WS.Cells(t, "I").Value
            End If
        Next t
        
        ' Update the greatest values in the summary table
        WS.Cells(2, "P").Value = MaxPercentChangeTicker
        WS.Cells(2, "Q").Value = MaxPercentChange
        WS.Cells(2, "Q").NumberFormat = "0.00%"
        WS.Cells(3, "P").Value = MinPercentChangeTicker
        WS.Cells(3, "Q").Value = MinPercentChange
        WS.Cells(3, "Q").NumberFormat = "0.00%"
        WS.Cells(4, "P").Value = MaxTotalStockVolumeTicker
        WS.Cells(4, "Q").Value = MaxTotalStockVolume
    Next WS
End Sub
Sub TickerSymbol()

    ' Create needed variables for ticker symbol
    Dim ticker_symbol As String

    Dim LastRow As LongLong

    ' Create variable for total stock volume and the row
    Dim total_stock_volume As LongLong
    Dim TotalStockRow As LongLong
    TotalStockRow = 2

     'Loop through all of the sheets in the wb
     For Each ws In Worksheets

      'Find the last row of each sheet
      LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
      
      ' Reset the value for the total stock row
      total_stock_volume = 0
      TotalStockRow = 2

        ' Loop through all of the ticker symbols
        For i = 2 To LastRow

            ' Check to see if ticker symbol is still the same during loop
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set Ticker Symbol
                ticker_symbol = ws.Cells(i, 1).Value

                ' Add to Total Stock Volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

                ' Print the ticker symbol in the summary chart
                ws.Cells(TotalStockRow, "J").Value = ticker_symbol

                ' Print total stock volume into summary chart
                ws.Cells(TotalStockRow, "M").Value = total_stock_volume

                ' Add headers for new columns
                ws.Cells(1, 10).Value = "Ticker"
                ws.Cells(1, 13).Value = "Total Stock Volume"

                ' Add a row to summary table for next entry
                TotalStockRow = TotalStockRow + 1

                ' Reset total stock volume
                total_stock_volume = 0

            'If the cell following the previous one is the same ticker symbol add to total stock volume
            Else

                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

            End If

        Next i


     Next ws
      

    End Sub




   

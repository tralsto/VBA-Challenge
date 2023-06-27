Sub CalculatedValues()

 ' Create variables of the summary table for the loop
 Dim ticker_symbol As String
 Dim Yearly_Change As Double
 Dim Percent_Change As Double
 Dim total_stock_volume As LongLong

 ' Create variables for new calculated values table
 Dim Greatest_Percent_Inc As Double
 Dim Greatest_Percent_Dec As Double
 Dim Greatest_Total_Vol As Double
 Dim Calculated_Values_Row As Double

    ' Loop through each sheet in the workbook
    For Each ws In Worksheets

     ' Define the last row
     LastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

     ' Reset Value for new table
     Calculated_Values_Row = 2

        ' Loop through the summary table
        For i = 2 To LastRow
         
         'Conditional for the different ticker symbols
         If ws.Cells(i + 2, 10).Value <> ws.Cells(i, 10).Value Then

         ' Set column values
         ticker_symbol = ws.Cells(i, 10).Value
         Yearly_Change = ws.Cells(i, 11).Value
         Percent_Change = ws.Cells(i, 12).Value
         total_stock_volume = ws.Cells(i, 13).Value

         ' Print headers and labels of new calculated values table
         ws.Cells(1, 17).Value = "Ticker"
         ws.Cells(1, 18).Value = "Value"
         ws.Cells(2, 16).Value = "Greatest % Increase"
         ws.Cells(3, 16).Value = "Greatest % Decrease"
         ws.Cells(4, 16).Value = "Greatest Total Volume"

<<<<<<< HEAD
         ' Print ticker symbols of calculated values in new table
=======
         ' Print ticker symbols of calculated values in the new table
>>>>>>> origin/main
         ticker_symbol = ws.Cells(Calculated_Values_Row, "Q").Value

         'Calculate Greatest Percent Increase
         Greatest_Percent_Inc = WorksheetFunction.Max(ws.Range("L:L"))

         ' Calculate Greatest Percent Decrease
         Greatest_Percent_Dec = WorksheetFunction.Min(ws.Range("L:L"))

         ' Calculate Greatest Total Volume
         Greatest_Total_Vol = WorksheetFunction.Max(ws.Range("M:M"))

         ' Print Values into Calculated Table
         ws.Cells(2, 18).Value = Greatest_Percent_Inc
         ws.Cells(3, 18).Value = Greatest_Percent_Dec
         ws.Cells(4, 18).Value = Greatest_Total_Vol


         End If

        Next i
      
      ' Format Greatest Percent Increase & Decrease to include %
      ws.Cells(2, 18).NumberFormat = "0.00%"
      ws.Cells(3, 18).NumberFormat = "0.00%"

    Next ws

End Sub

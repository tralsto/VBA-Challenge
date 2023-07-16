Sub StockData()

 ' Create needed variables
 Dim ws As Worksheet
 Dim LastRow As Long
 Dim ticker_symbol As String
 Dim Opening_Price As Double
 Dim Closing_Price As Double
 Dim Yearly_Change As Double
 Dim Percent_Change As Double
 Dim total_stock_volume As Double
 Dim Summary_Table_Row As Double
 Dim Greatest_Percent_Inc As Double
 Dim Greatest_Percent_Dec As Double
 Dim Greatest_Total_Vol As Double

  'Loop through all of the sheets in the wb
  For Each ws In Worksheets

    'Find the last row of each sheet
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
    ' Reset the value for the total stock row
    Yearly_Change = 0
    total_stock_volume = 0
    Summary_Table_Row = 2
  
    'First Opening Price
    Opening_Price = ws.Cells(2,3).Value

      ' Loop through all of the ticker symbols
      For i = 2 To LastRow

       total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

        ' Check to see if ticker symbol is still the same during loop
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

          ' Set column values
          ticker_symbol = ws.Cells(i, 1).Value
         
          Closing_Price = ws.Cells(i, 6).Value
         
          Yearly_Change = Closing_Price - Opening_Price

          ' Print the values in the first summary chart
          ws.Cells(Summary_Table_Row, "J").Value = ticker_symbol
          ws.Cells(Summary_Table_Row, "K").Value = Yearly_Change

          ' Format Yearly Change column so >0 will be green and <0 will be red
          If Yearly_Change >= 0 Then
            ws.Cells(Summary_Table_Row, "K").Interior.ColorIndex = 4

          ElseIf Yearly_Change < 0 Then
            ws.Cells(Summary_Table_Row, "K").Interior.ColorIndex = 3

          End If

          'Determine Percent Change from the Yearly Change data

          If Opening_Price <> 0 Then 
          Percent_Change = (Yearly_Change / Opening_Price)
          Else Percent_Change = 0 
          End If

          ' Print percent change values into summary chart
          ws.Cells(Summary_Table_Row, "L").Value = Percent_Change

          ' Print total stock volume into summary chart
          ws.Cells(Summary_Table_Row, "M").Value = total_stock_volume

          ' Add a row to summary table for next entry
          Summary_Table_Row = Summary_Table_Row + 1
          Opening_Price = ws.Cells(i+1, 3).Value

          ' Reset total stock volume
          total_stock_volume = 0

        End If
      Next i

    
      ' Loop through the summary table
      For i = 2 To LastRow
        
       Yearly_Change = ws.Cells(i, 11).Value
       Percent_Change = ws.Cells(i, 12).Value
       total_stock_volume = ws.Cells(i, 13).Value

       'Calculate Greatest Percent Increase
       Greatest_Percent_Inc = WorksheetFunction.Max(ws.Range("L:L"))
       Inc_Number = WorksheetFunction.match(WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0)

       ' Calculate Greatest Percent Decrease
       Greatest_Percent_Dec = WorksheetFunction.Min(ws.Range("L:L"))
       Dec_Number = WorksheetFunction.match(WorksheetFunction.Min(ws.Range("L:L")), ws.Range("L:L"), 0)

       ' Calculate Greatest Total Volume
       Greatest_Total_Vol = WorksheetFunction.Max(ws.Range("M:M"))
       Vol_Number = WorksheetFunction.match(WorksheetFunction.Max(ws.Range("M:M")), ws.Range("M:M"), 0)

       ' Print Values into Calculated Table
       ws.Cells(2, 18).Value = Greatest_Percent_Inc
       ws.Cells(2, 17).Value = ws.Cells(Inc_Number, 10)
       ws.Cells(3, 18).Value = Greatest_Percent_Dec
       ws.Cells(3, 17).Value = ws.Cells(Dec_Number, 10)
       ws.Cells(4, 18).Value = Greatest_Total_Vol
       ws.Cells(4, 17).Value = ws.Cells(Vol_Number, 10)

      Next i

   'Add headers to summary tables
   ws.Cells(1, 10).Value = "Ticker"
   ws.Cells(1, 11).Value = "Yearly Change"
   ws.Cells(1, 12).Value = "Percent Change"
   ws.Cells(1, 13).Value = "Total Stock Volume"
   ws.Cells(1, 17).Value = "Ticker"
   ws.Cells(1, 18).Value = "Value"
   ws.Cells(2, 16).Value = "Greatest % Increase"
   ws.Cells(3, 16).Value = "Greatest % Decrease"
   ws.Cells(4, 16).Value = "Greatest Total Volume"

   ' Format cells to include %
   ws.Columns("L").NumberFormat = "0.00%"
   ws.Cells(2, 18).NumberFormat = "0.00%"
   ws.Cells(3, 18).NumberFormat = "0.00%"

  Next ws
End Sub
Sub YearlyPercentChange()

 ' Create variables for the loop
 Dim ws As Worksheet
 Dim ticker_symbol As String
 Dim Opening_Price As Double
 Dim Closing_Price As Double
 Dim LastRowOpen As Long
 Dim LastRowClose As Long

 ' Create variables for the new summary column
 Dim Yearly_Change As Double
 Dim Percent_Change As Double
 Dim Summary_Table_Row As Double

  ' Loop through each sheet of the workbook
  For Each ws In Worksheets

   ' Set last row for each Opening and Closing Price Columns
   LastRowOpen = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
   LastRowClose = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row

   ' Reset value for new columns
   Yearly_Change = 0
   Summary_Table_Row = 2

    ' Loop through stock market data
    For i = 2 To LastRowOpen

      ' Set Ticker Column conditional to begin loop and register when next ticker has been reached
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ' Set column values
        ticker_symbol = ws.Cells(i, 1).Value
        Opening_Price = ws.Cells(i, 3).Value
        Closing_Price = ws.Cells(i, 6).Value

        'Add headers to summary table
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"

        ' Print ticker symbols in Summary Table
        ws.Cells(Summary_Table_Row, "J").Value = ticker_symbol

        ' Determine Yearly Change by taking the Closing Price from Opening Price
        Yearly_Change = (Closing_Price - Opening_Price)

        ' Print Yearly Change values into summary chart
        ws.Cells(Summary_Table_Row, "K").Value = Yearly_Change
 
        ' Format Yearly Change column so >0 will be green and <0 will be red
        If Yearly_Change >= 0 Then

          ws.Cells(Summary_Table_Row, "K").Interior.ColorIndex = 4

        ElseIf Yearly_Change < 0 Then

          ws.Cells(Summary_Table_Row, "K").Interior.ColorIndex = 3

        End If

        'Determine Percent Change from the Yearly Change data
        Percent_Change = (Yearly_Change / Opening_Price) * 100

        ' Print percent change values into summary chart
        ws.Cells(Summary_Table_Row, "L").Value = Percent_Change

        ' Add a row to the summary table
        Summary_Table_Row = Summary_Table_Row + 1

      End If

  
    Next i
   
    ' Format Percent Change cell to include %
    ws.Columns("L").NumberFormat = "0.00%"
   

  Next ws

End Sub
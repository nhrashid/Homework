
Sub StockMarket()

  ' Set an initial variable for holding the ticker name
  Dim Ticker_Symbol As String

  ' Set an initial variable for holding the total
  Dim Ticker_Total As Double
  Ticker_Total = 0
  

  ' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

' Loop through all Worksheets
 For Each ws In Worksheets
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Total Stock Volume"
  lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 ' Loop through all Ticker symbols

  For i = 2 To lastRow
  
 ' Check if we are still within the same Ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker symbol
      Ticker_Symbol = Cells(i, 1).Value
     
      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 3).Value

      ' Print the Ticker symbol in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Symbol

      ' Print the ticker Amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Ticker_Total = 0
    
     
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 3).Value

    End If

  Next i
 Next ws
End Sub


Sub mediumAssignment()

  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per stock Ticker
  Dim Volume As Double
  Volume = 0
  
  Dim OpenValue As Double
  OpenValue = 0
  Dim CloseValue As Double
  CloseValue = 0
  Dim YearDifference As Double
  YearDifference = 0
  

  ' Keep track of the location for each stock Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Setting data output names in table
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"

  ' Loop through all stock purchases
  For i = 2 To 705714
 
    ' Check if the next row has the appropriate closing value


    ' Check if we are not on the same ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set yearly opening value

        OpenValue = Cells(i, 3).Value

      ' Set the Ticker name
        Ticker_Name = Cells(i, 1).Value

      ' Add to the Ticker Total
        Volume = Volume + Cells(i, 7).Value

      ' Yearly difference
        YearDifference = OpenValue - CloseValue

        ' Print the stock Ticker in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker_Name
        
        Range("J" & Summary_Table_Row).Value = YearDifference
        If YearDifference = OpenValue Then
            Range("K" & Summary_Table_Row).Value = 0
        Else
            Range("K" & Summary_Table_Row).Value = 100 * (YearDifference / OpenValue)
        End If
            
      ' Print the Ticker Amount to the Summary Table
        Range("L" & Summary_Table_Row).Value = Volume



      ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the values
        Volume = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Ticker Total
      Volume = Volume + Cells(i, 3).Value
      CloseValue = Cells(i, 6).Value

    End If

  Next i

End Sub




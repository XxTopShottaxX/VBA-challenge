Sub StockTicker()

' Set values
    Dim Lastrow As Long
    Dim vol As Double
    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double


' Set column Names
    Cells(1, 9).Value = "Stock_Tick"
    Cells(1, 10).Value = "Yearly_change"
    Cells(1, 12).Value = "Total Stock Vol"
    Cells(1, 11).Value = "Yearly_percentage"
    Lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    year_open = Cells(2, 3)
    Summary_Table_Row = 2

For i = 2 To Lastrow

      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          year_close = Cells(i, 6).Value
          Yearly_change = year_close - year_open


          ticker = Cells(i, 1).Value
          vol = vol + Cells(i, 7).Value


          Cells(Summary_Table_Row, 10).Value = Yearly_change
          
          Cells(Summary_Table_Row, 9).Value = ticker
          
          Cells(Summary_Table_Row, 11).Value = year_percent
          
          Cells(Summary_Table_Row, 12).Value = vol

          Summary_Table_Row = Summary_Table_Row + 1

          vol = 0


        Yearly_change = year_close - year_open

    If year_open <> 0 And Yearly_close <> 0 Then
        year_percent = (year_close / year_open) * 100
    Else
        year_percent = 0
    End If




      Else

          vol = vol + Cells(i, 7).Value
          
          
          


      End If


    




    Next i

End Sub


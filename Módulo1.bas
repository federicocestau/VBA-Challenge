Attribute VB_Name = "Módulo1"
Sub STOCK()
Attribute STOCK.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ticker As String
Dim day_first As Double
Dim day_last As Double
Dim yr_change As Double
Dim percent_change As Double
Dim tot_stock_vol As LongLong
Dim summary_table_row As Integer
Dim Last_Row As Long
Dim Last_Row2 As Long
day_first = 0
day_last = 0
yr_change = 0
percent_change = 0
tot_stock_vol = 0
summary_table_row = 2
day_first = Cells(summary_table_row, 3).Value
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
Last_Row2 = Cells(Rows.Count, 7).End(xlUp).Row

For i = 2 To Last_Row

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
ticker = Cells(i, 1).Value
day_last = Cells(i, 6).Value
yr_change = day_last - day_first
If day_last > 0 And day_first > 0 Then

percent_change = ((day_last - day_first) / day_first)
End If
tot_stock_vol = tot_stock_vol + Cells(i, 7).Value
Range("I" & summary_table_row).Value = ticker
Range("J" & summary_table_row).Value = yr_change
If yr_change > 0 Then
Range("J" & summary_table_row).Interior.ColorIndex = 4
Else
Range("J" & summary_table_row).Interior.ColorIndex = 3
End If
Range("K" & summary_table_row).Value = percent_change
Range("K" & summary_table_row).NumberFormat = "0.00%"
Range("l" & summary_table_row).Value = tot_stock_vol
summary_table_row = summary_table_row + 1
day_first = 0
day_last = 0
yr_change = 0
percent_change = 0
tot_stock_vol = 0
If Cells(i + 1, 3).Value > 0 Then

day_first = Cells(i + 1, 3).Value

End If

Else
tot_stock_vol = tot_stock_vol + Cells(i, 7).Value

End If
Next i

Range("O2:O3").NumberFormat = "0.00%"

max_PercentChange = WorksheetFunction.Max(Range("k2:k" & Last_Row2))
max_TickerName = WorksheetFunction.Match(max_PercentChange, Range("k2:k" & Last_Row2), 0)
Range("N2") = Cells(max_TickerName + 1, 9)
Range("O2") = max_PercentChange

max_PercentDecrease = WorksheetFunction.Min(Range("k2:k" & Last_Row2))
Min_Tickername = WorksheetFunction.Match(max_PercentDecrease, Range("k2:k" & Last_Row2), 0)
Range("N3") = Cells(Min_Tickername + 1, 9)
Range("O3") = max_PercentDecrease

Max_TotalVolume = WorksheetFunction.Max(Range("L2:L" & Last_Row2))
Max_TickerTotalVolume = WorksheetFunction.Match(Max_TotalVolume, Range("L2:L" & Last_Row2), 0)
Range("N4") = Cells(Max_TickerTotalVolume + 1, 9)
Range("O4") = Max_TotalVolume



End Sub



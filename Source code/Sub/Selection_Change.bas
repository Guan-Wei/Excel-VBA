' 6.3.3 SelectionChange事件
' From <Excel 2016 Power Programming with VBA><中文版> Book
' Put this VBA code in your sheet(工作表) of excel.


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  Cells.Interior.ColorIndex = xlNone
  With ActiveCell
    .EntireRow.Interior.Color = RGB(219, 229, 241)
    .EntireColumn.Interior.Color = RGB(219, 229, 241)
  End With
End Sub

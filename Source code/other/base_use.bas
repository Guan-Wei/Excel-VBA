
' 1.Range
Worksheets("工作表1").Range("A1").Value = 2

' 2.Cells
Worksheets("工作表1").Cells(1, 1) = 3

' 3.offset(rowoffset,columnoffset)
Worksheets("工作表1").Cells(1, 1).offset(1,0) = 3 ' down
Worksheets("工作表1").Cells(1, 1).offset(0,1) = 3 ' right

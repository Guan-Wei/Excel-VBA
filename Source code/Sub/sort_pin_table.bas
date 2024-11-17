
'=============================================================
'sort pin
'=============================================================
Sub sort_pin_table()
'
'format: 1.pin name, 2. I/O, 3.netname.
'
Dim sort_range As String
Dim sort_item As Variant

sort_range = "$A$2:$A$358"
sort_item = Array("F9", "C12", "F13")

    ActiveSheet.Range(sort_range).AutoFilter Field:=1, Criteria1:=Array(sort_item), Operator:=xlFilterValues
    
End Sub

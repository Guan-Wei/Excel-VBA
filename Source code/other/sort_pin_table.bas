'=============================================================
'sort pin
'=============================================================
'
' -----------------------------
'  Ball name | I/O   | Netname |
' -----------------------------
'  A13       |input  | p3v3_pgd|
' -----------------------------
'  B12       |output | p3v3_en |
' -----------------------------

'       ^
'       |
'   Sort This)
'=============================================================
Sub sort_pin_table()
'
' 1. modify sort_range.
' 2. modify sort_item.
Dim sort_range As String
Dim sort_item As Variant

sort_range = "$A$2:$A$358"
sort_item = Array("W2", "W6", "Y6", "W20")

    ActiveSheet.Range(sort_range).AutoFilter Field:=1, Criteria1:=Array(sort_item), Operator:=xlFilterValues
    
End Sub

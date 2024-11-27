'=============================================================
'sort net name
'=============================================================
'
' -----------------------------
'  Ball name | I/O   | Netname |
' -----------------------------
'  A13       |input  | p3v3_pgd|
' -----------------------------
'  B12       |output | p3v3_en |
' -----------------------------

'                         ^
'                         |
'                      (Sort This)
'=============================================================
Sub sort_netname_table()
'
' 1. modify sort_range.
' 2. modify sort_item.
Dim sort_range As String
Dim sort_item As Variant

sort_range = "$A$2:$C$358"
sort_item = Array("p3v3_pgd", "p3v3_en")

    ActiveSheet.Range(sort_range).AutoFilter Field:=3, Criteria1:=Array(sort_item), Operator:=xlFilterValues
    
End Sub

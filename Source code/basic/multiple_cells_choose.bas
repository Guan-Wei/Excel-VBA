Option Explicit

'==============
'Explain
'==============
'
' ------------
'  Ball name |
' ------------
'  A13       |  <- modify offset(0,0) to choose netname.
' ------------
'  B12       |  <- modify offset(1,0) to choose netname.
' ------------
'  B13       |  <- modify offset(2,0) to choose netname.
' ------------
'
'
' PS: netname_data is "A2:A4".
'==============



Function test(netname_data As Range) As String

Dim num_datas    As Integer
Dim out_signal   As String

'1.計數選擇個數
num_datas = netname_data.Rows.Count

'2.選取此集合中個別的值(假設選取3個)
'out_signal = netname_data.Resize(1, 1).Offset(0, 0).Value 'A13
out_signal = netname_data.Resize(1, 1).Offset(1, 0).Value 'B12 <- here
'out_signal = netname_data.Resize(1, 1).Offset(2, 0).Value 'B13

'3.輸出
test = out_signal

End Function

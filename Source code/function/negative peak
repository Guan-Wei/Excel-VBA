---------
Location: 
---------
module

'======================================================================
' Code begin.
'======================================================================
'負源尖端偵測
Function NEG_PEAK(number1 As Range) As String
  
  Dim i As Long
  Dim cnt_peak As Long
  Dim max_index_col As Long
  Dim start_address As String
  Dim pos As Long
  Dim stock_date As String '2021/11/01
  
  'initial
  NEG_PEAK = ""
  cnt_peak = 0
  max_index_col = Range(number1.Address).Columns.Count
  stock_date = "" '2021/11/01
  
  If max_index_col >= 3 Then
    pos = InStr(number1.Address, ":") ' initial
    start_address = Left(number1.Address, pos - 1) ' initial
    
    For i = 1 To max_index_col - 2 '去頭去尾
      'If True Then
      If (Range(start_address).Offset(0, i).Value <= Range(start_address).Offset(0, i - 1).Value) And _
         (Range(start_address).Offset(0, i).Value <= Range(start_address).Offset(0, i + 1).Value) Then
        cnt_peak = cnt_peak + 1
        'NEG_PEAK = cnt_peak'number of peak
        'NEG_PEAK = Range(start_address).Offset(0, i).Address 'last peak address
        'NEG_PEAK = NEG_PEAK & Range(start_address).Offset(0, i).Address & "_" 'last peak address
        'NEG_PEAK = NEG_PEAK & Range(start_address).Offset(0, i).Value & "_" 'last peak Value'2021/11/01
        '----
        stock_date = "" ' Null '2021/11/02
        'stock_date = "(" & Left(Range("C1").Offset(0, i).Value, Len(Range("C1").Offset(0, i).Value) - 4) & ")" ' .xls '2021/11/02
        '----
        NEG_PEAK = NEG_PEAK & Range(start_address).Offset(0, i).Value & stock_date & "_" 'last peak Value'2021/11/01
      End If
    Next i
  Else
    NEG_PEAK = cnt_peak
  End If
End Function

'======================================================================
' Code end.
'======================================================================


ex:

Cell:
=POS_PEAK($C372:$V372)

Cell result: (Base on your data)[Separate char: "_"]
88.8_87.7_90.1_89_91_



'---------
' Location: 
'---------
module

'---------
' code:
'---------
    
Option Explicit
Sub remove()
'======================================
'移除行
'======================================
Dim 設定限制空格範圍 As Integer
Dim index_endmodule As Integer
Dim 字串計數_i As Integer

index_endmodule = 600'總共行數
設定限制空格範圍 = 100
字串計數_i = index_endmodule - 1'縮行最後位置

Do Until 字串計數_i = 0
    If check_all_row_nll(Range("A1").Offset(字串計數_i, 0).address, 設定限制空格範圍) Then
        'MsgBox Str(num_of_row) + "行內都是空的."
        'Range("A1").Offset(字串計數_i, 0).Interior.Color = RGB(255, 255, 0)
        'Range("A1").Offset(字串計數_i, 0).Select
        Rows(字串計數_i + 1).Delete
    End If
    
    字串計數_i = 字串計數_i - 1
Loop

End Sub

Function check_all_row_nll(start_address As String, limit_number As Integer) As Boolean
    Dim i As Integer
    
    For i = 0 To limit_number - 1
        If Range(start_address).Offset(0, i).Value <> "" Then Exit For
    Next
    
    'MsgBox "次數:" + Str(i + 1)
    
    If i = limit_number Then
        check_all_row_nll = True
    Else
        check_all_row_nll = False
    End If

End Function

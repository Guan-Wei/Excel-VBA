Attribute VB_Name = "Module1"
Option Explicit

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
'==================================================================
Sub sample_test()
    Dim num_of_row As Integer

    num_of_row = 4

    If check_all_row_nll(Range("A1").address, num_of_row) Then
        MsgBox Str(num_of_row) + "行內都是空的."
    Else
        MsgBox Str(num_of_row) + "行內有東西."
    End If

End Sub

'==================================================================

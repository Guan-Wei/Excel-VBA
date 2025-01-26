```Basic

' ======================================================
' Description:
'     Inser file path. Get file name.
' ex:
'     --------------------------------------------
'     | In data                  | get file name |
'     --------------------------------------------
'     | D:\test\item1\test.pdf   |   test.pdf
'     --------------------------------------------
'                                       ^
'                                       |
'                                   GET This
'                                  (=file_name(A1))
' ======================================================

Function file_name(data_in As String) As String

    Dim Target As String
    Dim total_len As Integer
    Dim local_bit As Integer
    Dim buf_data As String

    ' Auto decide "/" or "\"
    If (InStr(data_in, "/") <> 0) Then
        Target = "/"
    ElseIf (InStr(data_in, "\") <> 0) Then
        Target = "\"
    Else
        Target = "/"
    End If

    'get string
    buf_data = "" 'initial
    total_len = Len(data_in)
    local_bit = InStr(data_in, Target)
    buf_data = Right(data_in, total_len - local_bit)

    'more than one "/"
    Do While (InStr(buf_data, Target) <> 0)
        total_len = Len(buf_data)
        local_bit = InStr(buf_data, Target)
        buf_data = Right(buf_data, total_len - local_bit)
    Loop

    file_name = buf_data

End Function
```

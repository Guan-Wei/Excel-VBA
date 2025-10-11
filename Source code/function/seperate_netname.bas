' =============================================
' Description:
'     After setting a special character, take the string on the right or left.
' ex:
'     -------------------------------
'     | In data        | get signal |
'     -----------------------------
'     | input_p3v3_pgd_n | p3v3      |
'     -------------------------------
'                  or
'     | input_p3v3_pgd_n | pgd      |
'     -------------------------------
'                  or ...
'                            ^
'                            |
'                        GET This
'                       ( modify in code: seperate_netname = sig_array(1))
'                       ( modify in code: seperate_netname = sig_array(2))
'                                          or ...
' =============================================

Function seperate_netname(netname As String) As String

Dim num_underline As Integer 'under line
Dim locate_underline As Integer
Dim num_loop As Integer
Dim sig_array() As String
Dim tmp_sig As String

'0.產生陣列
num_underline = Len(netname) - Len(Replace(netname, "_", "")) 'initial

ReDim sig_array(num_underline + 1)

'1.值存入陣列
' bug: no consider no underline

tmp_sig = netname
For num_loop = 0 To (num_underline - 1)
    sig_array(num_loop) = get_string_all(tmp_sig, "_", "left")
    tmp_sig = get_string_all(tmp_sig, "_", "right")

Next num_loop

'last one
sig_array(num_underline) = tmp_sig

'-------------------------------------------
'main: choose number of array
'-------------------------------------------
seperate_netname = sig_array(0) ' modify the index of array to show

End Function


'Seperate Special char
Function get_string_all(data_in As String, set_string As String, sel As String) As String
    Dim i, k As Integer
    i = 1
    k = Len(data_in)

    i = InStr(data_in, set_string)
    Select Case sel
        Case Is = "right"
            get_string_all = Right(data_in, k - i)
        Case Is = "left"
            get_string_all = Left(data_in, i - 1)
        Case Else
            MsgBox "第三個參數設定錯誤: 只有 'right', 'left'."

    End Select

End Function

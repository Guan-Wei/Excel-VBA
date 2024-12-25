' =============================================
' Description:
'     After setting a special character, take the string on the right.
' ex:
'     -------------------------------
'     | In data        | get signal |
'     -----------------------------
'     | input-p3v3_pgd | p3v3_pgd   |
'     -------------------------------
'                            ^
'                            |
'                        GET This
'                       (=get_string(A1,"-"))
' =============================================
Option Explicit

Function get_string(data_in As String, set_string As String) As String
  Dim i, j As Integer
  
  i = 1
  j = Len(data_in)
  i = InStr(data_in, set_string)
  get_string = Right(data_in, j - i)

End Function


' =============================================
' Description:
'     After setting a special character, take the string on the right or left.
' ex:
'     -------------------------------
'     | In data        | get signal |
'     -----------------------------
'     | input-p3v3_pgd | p3v3_pgd   |
'     -------------------------------
'                            ^
'                            |
'                        GET This
'                       ( =get_string_all(A1,"-","right") )
'
'     -------------------------------
'     | input-p3v3_pgd | input      |
'     -------------------------------
'                            ^
'                            |
'                        GET This
'                       ( =get_string_all(A1,"-","left") )
'
' =============================================

Option Explicit

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


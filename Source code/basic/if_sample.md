***[if -elseif -else]:***

``` bas
Sub if_test()
    If Range("A1").Value = "" Then
        MsgBox "空"

    ElseIf Range("A1").Value >= 60 Then
        MsgBox "及格"

    ElseIf Range("A1").Value < 60 Then
        MsgBox "不及格"

    Else
        MsgBox "其他"
    End If

End Sub
```

***[For - to - next]:***

``` base
Sub for_test()
    Dim i As Integer
    Dim j As Integer
    j = 0
    For i = 1 To 3
        j = j + 1
    Next i
    Range("A1").Value = j
End Sub
```

***[For - to - step - next]:***
``` base
Sub for_test2()
    Dim i As Integer
    Dim j As Integer
    j = 0
    For i = 1 To 3 Step 2 ' 1, 3
        j = j + i
    Next i
    Range("A2").Value = j
End Sub
```

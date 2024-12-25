
**CV_DO_PATH.bas**
```basic

'----------------------------------------
' converter path format to do file.
' backslash converter to forward slash.
'----------------------------------------
' ex:
'      Path          => converter Path
'      D:\test\t123  => D:/test/t123
'----------------------------------------
Function CV_DO_PATH(number1 As Range) As String

    CV_DO_PATH = Replace(number1.Value, "\", "/")

End Function

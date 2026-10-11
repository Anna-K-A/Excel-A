
# Полезности в Excel
# Условное форматирование. Выделяем одинаковые данные уникальными цветами. 
<br/>
<img src=img/PoparnoeFormatirovanie.png width=100>
<br/>
<br/>
Чтобы это реализовать, используем следующий код:

```
Sub Color()
'Updateby Extendoffice
    Dim xRg As Range
    Dim xTxt As String
    Dim xCell As Range
    Dim xChar As String
    Dim xCellPre As Range
    Dim xCIndex As Long
    Dim xCol As Collection
    Dim I As Long
    On Error Resume Next
    If ActiveWindow.RangeSelection.Count > 1 Then
      xTxt = ActiveWindow.RangeSelection.AddressLocal
    Else
      xTxt = ActiveSheet.UsedRange.AddressLocal
    End If
    Set xRg = Application.InputBox("please select the data range:", "Kutools for Excel", xTxt, , , , , 8)
    If xRg Is Nothing Then Exit Sub
    xRg.Interior.ColorIndex = xlNone
    xCIndex = 2
    Set xCol = New Collection
    For Each xCell In xRg
      On Error Resume Next
      xCol.Add xCell, xCell.Text
      If Err.Number = 457 Then
        xCIndex = xCIndex + 1
        Set xCellPre = xCol(xCell.Text)
        If xCellPre.Interior.ColorIndex = xlNone Then xCellPre.Interior.ColorIndex = xCIndex
        xCell.Interior.ColorIndex = xCellPre.Interior.ColorIndex
      ElseIf Err.Number = 9 Then
        MsgBox "Too many duplicate companies!", vbCritical, "Kutools for Excel"
        Exit Sub
      End If
      On Error GoTo 0
    Next
End Sub
```

<br/>
В коде есть строчка, обеспечивающая сброс заливки при повторном запуске программы.
<br/>
<br>Это нужно для того,чтобы при внесении исправлений в данные (например, было две 1, ячейки с ними стали красными. А мы одну единицу решили исправить на 2) код срабатывал корректно.
<br/>
<br>Вот эта часть кода:
<br/>
<br>xRg.Interior.ColorIndex = xlNone
<br/>


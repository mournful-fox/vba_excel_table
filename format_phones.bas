Sub ФорматироватьТелефоны()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim rawPhone As String
    Dim digits As String
    Dim formattedPhone As String
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row ' Последняя строка в столбце E
    For i = 2 To lastRow
        rawPhone = ws.Cells(i, 5).Value
        ' Удаляем всё, кроме цифр
        digits = ""
        Dim j As Long
        For j = 1 To Len(rawPhone)
            If Mid(rawPhone, j, 1) Like "#" Then
                digits = digits & Mid(rawPhone, j, 1)
            End If
        Next j
        ' Заменяем первую 8 на 7, если надо
        If Left(digits, 1) = "8" Then
            digits = "7" & Mid(digits, 2)
        End If
        ' Приводим к формату, если 11 цифр
        If Len(digits) = 11 Then
            formattedPhone = "+7(" & Mid(digits, 2, 3) & ")" & Mid(digits, 5, 3) & "-" & Mid(digits, 8, 2) & "-" & Mid(digits, 10, 2)
        Else
            formattedPhone = "Некорректный номер"
        End If
        ' Записываем в столбец K
        ws.Cells(i, 11).Value = formattedPhone
    Next i
End Sub

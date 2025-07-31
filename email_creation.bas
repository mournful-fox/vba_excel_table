Sub СформироватьПисьмаOutlook()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim i As Integer
    Dim lastRow As Integer
    Dim имя As String, email As String, дата As String, сумма As String, статус As String
    Dim тело As String
    Set OutlookApp = CreateObject("Outlook.Application")
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = 6 ' Только первые 5 (2-6 строки)
    'Для формирования писем для всех строк:
    'lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        статус = Trim(LCase(ws.Cells(i, 10).Value)) ' Столбец J = 10; приведение к нижнему регистру
        ' Формировать письмо только для клиентов со статусом "потенциальный"
        If статус = "потенциальный" Then
            имя = ws.Cells(i, 2).Value
            email = ws.Cells(i, 4).Value
            дата = ws.Cells(i, 7).Value
            сумма = ws.Cells(i, 9).Value
            тело = "Добрый день, " & имя & "!" & vbCrLf & _
                   "Ура, Вы с нами!" & vbCrLf & _
                   "Ваша дата регистрации: " & дата & vbCrLf & _
                   "Сумма покупок: " & сумма & vbCrLf & vbCrLf & _
                   "С уважением," & vbCrLf & _
                   "ООО ""Рога и копыта"""
            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = email
                .Subject = "Добро пожаловать!"
                .Body = тело
                .Display ' Открыть как черновик
                ' Если необходимо сразу отправлять, заменить .Display на:
                '.Send
            End With
        End If
    Next i
End Sub

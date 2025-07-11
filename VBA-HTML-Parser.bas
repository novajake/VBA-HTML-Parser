Sub ImportHTMLruTEXT()
    Dim htmlDoc As Object
    Dim htmlTable As Object
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim htmlContent As String
    Dim stream As Object
    
    ' Создаем объекты для работы с файлами и кодировкой
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Тип объекта
    stream.Charset = "utf-8" ' Указываем кодировку (P.S можно еще windows-1251)
    stream.Open
    stream.LoadFromFile "C:\Users\USER_NAME\FOLDER_NAME\FILE_NAME.html" ' Тут вписать путь до файла
    htmlContent = stream.ReadText
    stream.Close
    
    ' Создаем объект для парсинга HTML
    Set htmlDoc = CreateObject("htmlfile")
    htmlDoc.Open
    htmlDoc.Write htmlContent
    htmlDoc.Close
    
    ' Берем первую таблицу
    Set htmlTable = htmlDoc.getElementsByTagName("table")(0)
    
    ' Создаем лист в Excel
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Импорт HTML"
    
    ' Копируем данные в Excel
    For i = 0 To htmlTable.Rows.Length - 1
        For j = 0 To htmlTable.Rows(i).Cells.Length - 1
            ' Чистим текст и заменяем переносы
            Dim cellText As String
            cellText = Replace(htmlTable.Rows(i).Cells(j).innerText, vbLf, " ")
            ws.Cells(i + 1, j + 1) = cellText
        Next j
    Next i
    
    ' Подбираем ширину колонок под текст / данные
    ws.Columns.AutoFit
    
    MsgBox "Готово! HTML таблица импортирована на лист '" & ws.Name & "'", vbInformation
End Sub

Option Explicit On

Imports System.Web.Script.Serialization
Imports Inventor
Imports Newtonsoft.Json

Public Class PartsWebForm
    Inherits System.Web.UI.Page

    '0. загрузка страницы
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'найти текущий сеанс Inventor (если Inventor не запущен - запустить)
        Try
            'пытаемся получить ссылку на запущенный Inventor
            _invApp = Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
        Catch ex As Exception
            'если не удалось получить ссылку (например, Inventor не запущен), то код ниже попытается создать новый сеанс Inventor.
            Try
                Dim invAppType As Type = Type.GetTypeFromProgID("Inventor.Application")
                _invApp = Activator.CreateInstance(invAppType)
                _invApp.Visible = True
                _isAppAutoStarted = True
            Catch ex2 As Exception
                MsgBox(ex2.ToString(), vbSystemModal, "Ошибка")
            End Try
        End Try

        If _invApp Is Nothing Then
            MsgBox("Не удалось ни найти, ни создать сеанс Inventor", vbSystemModal, "Ошибка")
            Server.Transfer("ErrorWebForm.aspx")
        End If
    End Sub

    'функция по нажатию кнопки экспорт таблицы в excel
    'нельзя вынести в модуль, много привязок к элементам конкретной страницы
    Protected Sub btnExportTable_Click(sender As Object, e As EventArgs) Handles btnExportTable.Click
        If lblCountOfRows.Text = "0" Or _finalMessageString = "" Then
            MsgBox("Сначала необходимо получить данные из деталей", vbSystemModal, "Ошибка")
            Exit Sub
        End If

        Dim StrHtmlGenerate As StringBuilder = New StringBuilder()
        Dim StrExport As StringBuilder = New StringBuilder()
        'старт html
        StrExport.Append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'><head><title>Time</title>")
        StrExport.Append("<body lang=EN-US style='mso-element:header' id=h1><span style='mso--code:DATE'></span><div class=Section1>")
        StrExport.Append("<DIV  style='font-size:12px;'>")
        StrExport.Append(_finalMessageString & "<br /><br />") 'дополнительная текстовая инф-я (не таблица)
        StrExport.Append(tableOfResults.InnerHtml) 'таблица (содержимое из div, id которого = tableOfResults)
        StrExport.Append("</div></body></html>")
        'наименование выходного файла
        Dim strFile As String = "Results from grader (" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & ").xls"
        Dim strcontentType As String = "application/excel"
        Response.ClearContent()
        Response.ClearHeaders()
        Response.BufferOutput = True
        Response.ContentType = strcontentType
        Response.AddHeader("Content-Disposition", "attachment; filename=" + strFile)
        Response.Write(StrExport.ToString())
        Response.Flush()
        Response.Close()
        Response.End()
    End Sub

    'функция по нажатию кнопки очистить таблицу
    Protected Sub btnClearTable_Click(sender As Object, e As EventArgs) Handles btnClearTable.Click
        _listOfStndAsmCriteria.Clear() 'список для хранения данных самой стандартной сборки
        _listOfStndPartAndDrawCriteria.Clear() 'список для хранения данных деталей и чертежей стандартной сборки
        _listOfChekAsmCriteria.Clear() 'список для хранения данных самой проверяемой сборки
        _listOfChekPartAndDrawCriteria.Clear() 'список для хранения данных деталей и чертежей проверяемой сборки
        _listOfResults.Clear() 'список результатов сравнения, выводится в таблицу на веб-форму
        _stndDocumentSaveLocation = "" 'путь хранения документа стандартной сборки
        _chekDocumentSaveLocation = "" 'путь хранения документа проверяемой сборки
        lblCountOfRows.Text = "0" 'очистка элемента вывода количества строк
        tableOfResults.InnerHtml = "" 'очистка таблицы
    End Sub

    'функция по нажатию кнопки загрузить файлы на сервер
    Private Sub SubmitToServer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubmitToServer.ServerClick
        _stndDocumentSaveLocation = ""
        _chekDocumentSaveLocation = ""
        _invApp.Documents.CloseAll()

        '1. загрузить файл эталонной детали на сервер
        If Not locationOfStandardDocument.PostedFile Is Nothing And locationOfStandardDocument.PostedFile.ContentLength > 0 Then
            Dim fn As String = System.IO.Path.GetFileName(locationOfStandardDocument.PostedFile.FileName)
            _stndDocumentSaveLocation = Server.MapPath("Data") & "\" & fn
            Try
                locationOfStandardDocument.PostedFile.SaveAs(_stndDocumentSaveLocation)
            Catch Exc As Exception
                MsgBox("Ошибка: " & Exc.Message, vbSystemModal, "Ошибка")
                Exit Sub
            End Try
        Else
            MsgBox("Не выбран файл эталонной детали для загрузки", vbSystemModal, "Ошибка")
            Exit Sub
        End If

        '2. загрузить файл проверяемой детали на сервер
        If Not locationOfCheckedDocument.PostedFile Is Nothing And locationOfCheckedDocument.PostedFile.ContentLength > 0 Then
            Dim fn As String = System.IO.Path.GetFileName(locationOfCheckedDocument.PostedFile.FileName)
            _chekDocumentSaveLocation = Server.MapPath("Data") & "\" & fn
            Try
                locationOfCheckedDocument.PostedFile.SaveAs(_chekDocumentSaveLocation)
            Catch Exc As Exception
                MsgBox("Ошибка: " & Exc.Message, vbSystemModal, "Ошибка")
                Exit Sub
            End Try
        Else
            MsgBox("Не выбран файл проверяемой детали для загрузки", vbSystemModal, "Ошибка")
            Exit Sub
        End If

        'продолжить работу, только если оба пути к файлам заполнены
        If _stndDocumentSaveLocation IsNot "" And _chekDocumentSaveLocation IsNot "" Then
            '3. очистка 4х листов для хранения данных всех сравниваемых документов
            'для деталей, далее, фактически будут использоваться только 2 листа (..PartAndDrawCriteria), а остальные оставться пустыми
            _listOfStndAsmCriteria.Clear()
            _listOfStndPartAndDrawCriteria.Clear()
            _listOfChekAsmCriteria.Clear()
            _listOfChekPartAndDrawCriteria.Clear()

            '4. работа с эталонной деталью (и всеми входящими в нее документами)
            WorkWithPart(_listOfStndAsmCriteria, _listOfStndPartAndDrawCriteria, _stndDocumentSaveLocation)

            '5. работа с проверяемой деталью (и всеми входящими в нее документами)
            WorkWithPart(_listOfChekAsmCriteria, _listOfChekPartAndDrawCriteria, _chekDocumentSaveLocation)

            'MsgBox("Дошли до проверка: длины листов")
            'проверка: длины листов эталонные_детали и проверяемые_детали должны быть равны, иначе работа прекращается
            If Not _listOfStndPartAndDrawCriteria.Count = _listOfChekPartAndDrawCriteria.Count Then
                MsgBox("Ошибка. Количество деталей в сборках не совпадает, сравнение не может быть проведено.", vbSystemModal, "Ошибка")
                Exit Sub
            End If

            '6. сравнение полученных списков
            _listOfResults.Clear() 'сначала необходимо очистить список результатов
            'потом идет работа с деталями
            compare(_listOfStndPartAndDrawCriteria, _listOfChekPartAndDrawCriteria)

            '7. вывод результатов: создание и заполнение таблицы данными листа результатов (тип Result)
            fillResultsTable()

            '8. вывод текстовых результатов
            finalMessage()

            '9. сохранение данных из списков в 2 json файла (записать в json, сериализация)
            Dim ser As New JavaScriptSerializer()
            Dim results As String = ser.Serialize(_listOfStndPartAndDrawCriteria)
            System.IO.File.WriteAllText(Server.MapPath("Data") & "\" & "StandardPart.json", results)
            results = ser.Serialize(_listOfChekPartAndDrawCriteria)
            System.IO.File.WriteAllText(Server.MapPath("Data") & "\" & "CheckedPart.json", results)
        End If
    End Sub

    'работа с деталью
    Private Sub WorkWithPart(ByVal listOfAsmCriteria As List(Of Criteria), ByVal listOfPartAndDrawCriteria As List(Of Criteria), ByVal partSaveLocation As String)
        'открыть документ детали
        Dim partDoc As PartDocument
        Try
            partDoc = _invApp.Documents.Open(partSaveLocation)
        Catch Exc As Exception
            MsgBox("Не удалось открыть документ детали. Ошибка: " & Exc.Message, vbSystemModal, "Ошибка")
            Exit Sub
        End Try

        'продолжать работу с Inventor можно, если открыт 1 документ, и тип открытого документа - Part
        If (_invApp.Documents.Count > 0) And (_invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject) Then
            'работа с деталью и ее чертежом - заполнение данных для них
            'десериализация файла json для детали и чертежа, создание списка критериев на один документ
            deserialization(_name_of_list_of_criteria_for_part_and_drawing, listOfPartAndDrawCriteria, Server.MapPath("Data"))

            'непосредственное получение данных детали из Inventor
            getPartAndDrawingData(partDoc, listOfPartAndDrawCriteria)

            'критерии в списках заполнены. закрыть конкретную деталь
            partDoc.Close()
        Else
            MsgBox("Не удалось открыть документ детали.", vbSystemModal, "Ошибка")
            Exit Sub
        End If
    End Sub

    'вспомогательная функция: создание и заполнение html-табоицы
    'нельзя вынести в модуль, много привязок к элементам конкретной страницы
    Private Sub fillResultsTable()
        lblCountOfRows.Text = "0" 'очистка элемента вывода количества строк

        tableOfResults.InnerHtml = "" 'очистка таблицы
        'добавление кода будет проводиться к div-у с ID "tableOfResults"
        tableOfResults.InnerHtml += "<table border='1' bordercolor='#616161' cellpadding=4 cellspacing=0 style='font: sans-serif;' style='font-weight: normal;'>" 'открывающий тэг таблицы

        'add table columns header
        tableOfResults.InnerHtml += "<thead style='background: #0D47A1; color:  white; font-weight: normal;'>"
        tableOfResults.InnerHtml += "<th>Название критерия</th>"
        tableOfResults.InnerHtml += "<th>Вес критерия</th>"
        tableOfResults.InnerHtml += "<th>Допустимое отклонение, точность (%)</th>"
        tableOfResults.InnerHtml += "<th>Значение из эталонной сборки</th>"
        tableOfResults.InnerHtml += "<th>Значение из проверяемой сборки</th>"
        tableOfResults.InnerHtml += "<th>Отклонение проверяемой величины (%)</th>"
        tableOfResults.InnerHtml += "</thead>"

        'start new row
        'заполнение из листа _listOfResults
        For i = 0 To _listOfResults.Count - 1
            'если ответ не правильный - выделить красным
            If _listOfResults(i).is_correct = 0 Then
                tableOfResults.InnerHtml += "<tr style='background: #FFEBEE'>"
            End If
            'если ответ правильный в пределах диапазона - выделить светло-зеленым
            If _listOfResults(i).is_correct = 1 Then
                tableOfResults.InnerHtml += "<tr style='background: #DCEDC8'>"
            End If
            'если ответ правильный - выделить зеленым
            If _listOfResults(i).is_correct = 2 Then
                tableOfResults.InnerHtml += "<tr style='background: #C8E6C9'>"
            End If

            tableOfResults.InnerHtml += "<td style='font-weight: bold'>" & _listOfResults(i).name & "</td>"
            tableOfResults.InnerHtml += "<td>" & _listOfResults(i).weight & "</td>"
            tableOfResults.InnerHtml += "<td>" & _listOfResults(i).tolerance & "</td>"
            tableOfResults.InnerHtml += "<td style='font-weight: bold'>" & _listOfResults(i).standard_value & "</td>"
            tableOfResults.InnerHtml += "<td style='font-weight: bold'>" & _listOfResults(i).checked_value & "</td>"
            tableOfResults.InnerHtml += "<td>" & _listOfResults(i).delta & "</td>"

            tableOfResults.InnerHtml += "</tr>"
        Next

        tableOfResults.InnerHtml += "</table>" 'закрвыающий тэг таблицы

        lblCountOfRows.Text = _listOfResults.Count.ToString() 'заполнение элемента вывода количества строк
    End Sub

End Class
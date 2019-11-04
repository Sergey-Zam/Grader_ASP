Option Explicit On

Imports System
Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Web.Script.Serialization
Imports Inventor
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports System.Data

Public Class MainWebForm
    Inherits System.Web.UI.Page

    'глобальные переменные
    Dim _invApp As Inventor.Application = Nothing 'приложение Inventor
    Dim _isAppAutoStarted As Boolean = False 'был ли данный сеанс Inventor создан программой
    Dim _listOfStndAsmCriteria As New List(Of Criteria)() 'список для хранения данных самой стандартной сборки
    Dim _listOfStndPartAndDrawCriteria As New List(Of Criteria)() 'список для хранения данных деталей и чертежей стандартной сборки
    Dim _listOfChekAsmCriteria As New List(Of Criteria)() 'список для хранения данных самой проверяемой сборки
    Dim _listOfChekPartAndDrawCriteria As New List(Of Criteria)() 'список для хранения данных деталей и чертежей проверяемой сборки
    Dim _listOfResults As New List(Of Result)() 'список результатов сравнения, выводится в таблицу на веб-форму
    Dim _stndAsmSaveLocation As String = "" 'путь хранения документа стандартной сборки
    Dim _chekAsmSaveLocation As String = "" 'путь хранения документа проверяемой сборки

    '0. загрузка страницы
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Страница доступна в первый раз.
        If Not IsPostBack Then
            MsgBox("Внимание! При продолжении работы, программа попытается найти и запустить Inventor, этот процесс может занять несколько минут")
        End If

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
                MsgBox(ex2.ToString())
            End Try
        End Try

        If _invApp Is Nothing Then
            MsgBox("Не удалось ни найти, ни создать сеанс Inventor")
            Server.Transfer("ErrorWebForm.aspx")
        End If
    End Sub

    'функция по нажатию кнопки экспорт таблицы в excel
    Protected Sub btnExportTable_Click(sender As Object, e As EventArgs) Handles btnExportTable.Click
        If lblCountOfRows.Text = "0" Then
            MsgBox("Сначала необходимо получить данные из сборок")
            Exit Sub
        End If

        Dim StrHtmlGenerate As StringBuilder = New StringBuilder()
        Dim StrExport As StringBuilder = New StringBuilder()
        StrExport.Append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'><head><title>Time</title>")
        StrExport.Append("<body lang=EN-US style='mso-element:header' id=h1><span style='mso--code:DATE'></span><div class=Section1>")
        StrExport.Append("<DIV  style='font-size:12px;'>")
        StrExport.Append(tableOfResults.InnerHtml)
        StrExport.Append("</div></body></html>")
        Dim strFile As String = "results_from_grader.xls"
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
        _stndAsmSaveLocation = "" 'путь хранения документа стандартной сборки
        _chekAsmSaveLocation = "" 'путь хранения документа проверяемой сборки
        lblCountOfRows.Text = "0" 'очистка элемента вывода количества строк
        tableOfResults.InnerHtml = "" 'очистка таблицы
    End Sub

    'функция по нажатию кнопки загрузить файлы на сервер
    Private Sub SubmitToServer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubmitToServer.ServerClick
        _stndAsmSaveLocation = ""
        _chekAsmSaveLocation = ""

        '1. загрузить файл эталонной сборки на сервер
        If Not locationOfStandardAssembly.PostedFile Is Nothing And locationOfStandardAssembly.PostedFile.ContentLength > 0 Then
            Dim fn As String = System.IO.Path.GetFileName(locationOfStandardAssembly.PostedFile.FileName)
            _stndAsmSaveLocation = Server.MapPath("Data") & "\" & fn
            Try
                locationOfStandardAssembly.PostedFile.SaveAs(_stndAsmSaveLocation)
            Catch Exc As Exception
                MsgBox("Ошибка: " & Exc.Message)
                Exit Sub
            End Try
        Else
            MsgBox("Не выбран файл эталонной сборки для загрузки")
            Exit Sub
        End If

        '2. загрузить файл проверяемой сборки на сервер
        If Not locationOfCheckedAssembly.PostedFile Is Nothing And locationOfCheckedAssembly.PostedFile.ContentLength > 0 Then
            Dim fn As String = System.IO.Path.GetFileName(locationOfCheckedAssembly.PostedFile.FileName)
            _chekAsmSaveLocation = Server.MapPath("Data") & "\" & fn
            Try
                locationOfCheckedAssembly.PostedFile.SaveAs(_chekAsmSaveLocation)
            Catch Exc As Exception
                MsgBox("Ошибка: " & Exc.Message)
                Exit Sub
            End Try
        Else
            MsgBox("Не выбран файл проверяемой сборки для загрузки")
            Exit Sub
        End If

        'продолжить работу, только если оба пути к файлам заполнены
        If _stndAsmSaveLocation IsNot "" And _chekAsmSaveLocation IsNot "" Then
            WorkWithFiles()
        End If
    End Sub


    Private Sub WorkWithFiles()
        '3. очистка 4х листов для хранения данных всех сравниваемых документов
        _listOfStndAsmCriteria.Clear()
        _listOfStndPartAndDrawCriteria.Clear()
        _listOfChekAsmCriteria.Clear()
        _listOfChekPartAndDrawCriteria.Clear()

        '4. работа с эталонной сборкой (и всеми входящими в нее документами)
        WorkWithAsm(_listOfStndAsmCriteria, _listOfStndPartAndDrawCriteria, _stndAsmSaveLocation)

        '5. работа с проверяемой сборкой (и всеми входящими в нее документами)
        WorkWithAsm(_listOfChekAsmCriteria, _listOfChekPartAndDrawCriteria, _chekAsmSaveLocation)

        '6. сравнение списков, подсчет и вывод результатов
        CompareListsAndOutputResult(_listOfStndAsmCriteria, _listOfStndPartAndDrawCriteria, _listOfChekAsmCriteria, _listOfChekPartAndDrawCriteria)

        '7. сохранение данных из списков в 4 json файла (записать в json, сериализация)
        Dim ser As New JavaScriptSerializer()
        Dim results As String = ser.Serialize(_listOfStndAsmCriteria)
        System.IO.File.WriteAllText(Server.MapPath("Data") & "\" & "StandardAssembly.json", results)

        results = ser.Serialize(_listOfStndPartAndDrawCriteria)
        System.IO.File.WriteAllText(Server.MapPath("Data") & "\" & "StandardPart.json", results)

        results = ser.Serialize(_listOfChekAsmCriteria)
        System.IO.File.WriteAllText(Server.MapPath("Data") & "\" & "CheckedAssembly.json", results)

        results = ser.Serialize(_listOfChekPartAndDrawCriteria)
        System.IO.File.WriteAllText(Server.MapPath("Data") & "\" & "CheckedPart.json", results)
    End Sub

    'работа со сборкой
    Private Sub WorkWithAsm(ByVal listOfAsmCriteria As List(Of Criteria), ByVal listOfPartAndDrawCriteria As List(Of Criteria), ByVal assemblySaveLocation As String)
        '1. открыть документ сборки
        Dim asmDoc As AssemblyDocument
        Try
            asmDoc = _invApp.Documents.Open(assemblySaveLocation)
        Catch Exc As Exception
            MsgBox("Не удалось открыть документ сборки. Ошибка: " & Exc.Message)
            Exit Sub
        End Try

        'продолжать работу с Inventor можно, если открыт 1 документ, и тип открытого документа - Assembly
        If (_invApp.Documents.Count > 0) And (_invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
            '2. работа со сборкой - заполнение критериев
            FillAssembly(asmDoc, listOfAsmCriteria)

            '3. работа с деталями и чертежами - заполнение критериев
            FillPartAndDraw(asmDoc, listOfPartAndDrawCriteria)

            '4. критерии в списках заполнены. закрыть все открытые документы
            _invApp.Documents.CloseAll()
        Else
            MsgBox("Не удалось открыть документ сборки.")
            Exit Sub
        End If
    End Sub

    Private Sub CompareListsAndOutputResult(ByVal listOfStandardAssemblyCriteria As List(Of Criteria), ByVal listOfStandardPartCriteria As List(Of Criteria), ByVal listOfCheckedAssemblyCriteria As List(Of Criteria), ByVal listOfCheckedPartCriteria As List(Of Criteria))
        'проверка: длины листов эталонная_сборка и проверяемая_сборка должны быть равны
        If Not _listOfStndAsmCriteria.Count = _listOfChekAsmCriteria.Count Then
            MsgBox("Ошибка. Количество сборок не совпадает, сравнение не может быть проведено.")
            Exit Sub
        End If

        'проверка: длины листов эталонные_детали и проверяемые_детали должны быть равны
        If Not _listOfStndPartAndDrawCriteria.Count = _listOfChekPartAndDrawCriteria.Count Then
            MsgBox("Ошибка. Количество деталей в сборках не совпадает, сравнение не может быть проведено.")
            Exit Sub
        End If

        'сравнение полученных четырех списков
        compareLists()

        'создание и заполнение таблицы данными листа результатов (тип Result)
        'и прочий вывод
        outputResult()
    End Sub

    'заполнение сборок данными
    Private Sub FillAssembly(asmDoc As AssemblyDocument, ByVal listOfAsmCriteria As List(Of Criteria))
        'сборка одна, циклы не нужны.
        'десериализация файла json для сборок, создание списка криериев на один документ
        Dim jsonString = System.IO.File.ReadAllText(Server.MapPath("Data") & "\" & "list_of_criteria_for_assembly.json")
        Dim values = JsonConvert.DeserializeObject(Of List(Of Criteria))(jsonString)
        For Each value In values
            listOfAsmCriteria.Add(New Criteria())
            listOfAsmCriteria.Last.id = value.id
            listOfAsmCriteria.Last.name = value.name
            listOfAsmCriteria.Last.weight = value.weight
            listOfAsmCriteria.Last.tolerance = value.tolerance
            listOfAsmCriteria.Last.value = value.value
        Next

        'непосредственное получение данных сборки из Inventor
        getAsmData(asmDoc, listOfAsmCriteria)
    End Sub

    'заполнение деталей и чертежей данными
    Private Sub FillPartAndDraw(asmDoc As AssemblyDocument, ByVal listOfPartAndDrawCriteria As List(Of Criteria))
        'деталей в сборке несколько, нужен цикл
        'пройти по всем деталям, которые есть в сборке
        For i As Integer = asmDoc.AllReferencedDocuments.Count To 1 Step -1
            'текущая деталь, для которой буду получены данные
            Dim currentPartDoc As PartDocument = asmDoc.AllReferencedDocuments(i)

            'десериализация файла json для сборок, создание списка криериев на один документ
            Dim jsonString = System.IO.File.ReadAllText(Server.MapPath("Data") & "\" & "list_of_criteria_for_part_and_drawing.json")
            Dim values = JsonConvert.DeserializeObject(Of List(Of Criteria))(jsonString)
            For Each value In values
                listOfPartAndDrawCriteria.Add(New Criteria())
                listOfPartAndDrawCriteria.Last.id = value.id
                listOfPartAndDrawCriteria.Last.name = value.name
                listOfPartAndDrawCriteria.Last.weight = value.weight
                listOfPartAndDrawCriteria.Last.tolerance = value.tolerance
                listOfPartAndDrawCriteria.Last.value = value.value
            Next

            'непосредственное получение данных детали  из Inventor
            getPartData(currentPartDoc, listOfPartAndDrawCriteria)

            'для текущей детали определить, есть ли для нее чертеж. если есть, получить данные и чертежа
            Dim drawingFullFileName As String = findDrawingFullFileNameForDocument(currentPartDoc) 'найти, если возможно, путь к чертежу детали
            'если путь к чертежу найден, инициализировать переменную чертежа и открыть чертеж
            If Not String.IsNullOrEmpty(drawingFullFileName) Then
                Dim drawingDoc As Document = _invApp.Documents.Open(drawingFullFileName) 'открыть чертеж

                'непосредственное получение данных чертежа из Inventor
                getDrawingData(drawingDoc, listOfPartAndDrawCriteria)
            End If
        Next

    End Sub

    'непосредственное получение данных сборки из Inventor
    Private Sub getAsmData(ByVal asmDoc As AssemblyDocument, ByVal listOfAsmCriteria As List(Of Criteria))
        'получение первого и последнего индекса интервала листа, с которым идет работа 
        Dim first As Integer = listOfAsmCriteria.Count - 11 '11-число записей в list_of_criteria_for_assembly.json
        Dim last As Integer = listOfAsmCriteria.Count - 1 'здесь всегда 1

        'заполнение value-получить данные сборки
        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = asmDoc.PropertySets
        Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information") ' Get the Inventor Summary Information property set.
        Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information") ' Get the Inventor Document Summary Information property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties") ' Get the Design Tracking Properties property set.

        '"Автор"
        writeToList("A1", oPropSetISI.Item("Author").Value, listOfAsmCriteria, first, last)

        '"Имя документа"
        writeToList("A2", oPropSetDTP.Item("Part Number").Value, listOfAsmCriteria, first, last)

        '"Название сборки"
        writeToList("A3", oPropSetDTP.Item("Description").Value, listOfAsmCriteria, first, last)

        '"Материал"
        writeToList("A4", oPropSetDTP.Item("Material").Value, listOfAsmCriteria, first, last)

        '"Дата создания фаила"
        Dim filespec As String = asmDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        writeToList("A5", f.DateCreated.ToString, listOfAsmCriteria, first, last)

        '"Дата изменения фаила"
        writeToList("A6", f.DateLastModified.ToString, listOfAsmCriteria, first, last)

        '"Количество деталей в сборке"
        writeToList("A7", asmDoc.AllReferencedDocuments.Count.ToString, listOfAsmCriteria, first, last)

        '"Все детали сборки закреплены (0 степеней свободы)"
        Dim occ As ComponentOccurrence
        Dim result As String = True
        For Each occ In asmDoc.ComponentDefinition.Occurrences 'occ - свойства part document (1..n) В assembly, их (документов) перебор
            result = occ.Grounded 'true - да, деталь закреплена, false - нет, деталь не закреплена
            If result = False Then
                Exit For
            End If
        Next
        writeToList("A8", result, listOfAsmCriteria, first, last)

        '"Масса сборки"     
        Dim massProps As MassProperties = asmDoc.ComponentDefinition.MassProperties
        writeToList("A9", massProps.Mass, listOfAsmCriteria, first, last)
        'дополнит. способ (к MassProperties):
        'Dim massProps As MassProperties = asmDoc.ComponentDefinition.MassProperties
        'Dim uom As UnitsOfMeasure = asmDoc.UnitsOfMeasure
        'Dim defaultLength As String = uom.GetStringFromType(uom.LengthUnits)
        'MsgBox(uom.GetStringFromValue(massProps.Volume, defaultLength & "^3"))

        '"Площадь сборки"
        writeToList("A10", massProps.Area * 100, listOfAsmCriteria, first, last)

        '"Объем сборки"
        writeToList("A11", massProps.Volume * 1000, listOfAsmCriteria, first, last)
    End Sub

    'непосредственное получение данных детали  из Inventor
    Private Sub getPartData(ByVal partDoc As PartDocument, ByVal listOfPartAndDrawCriteria As List(Of Criteria))
        'получение первого и последнего индекса интервала листа, с которым идет работа 
        Dim first As Integer = listOfPartAndDrawCriteria.Count - 15 '15-число записей в list_of_criteria_for_part_and_drawing.json
        Dim last As Integer = listOfPartAndDrawCriteria.Count - 1 'здесь всегда 1

        'заполнение value-получить данные детали
        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information") ' Get the Inventor Summary Information property set.
        Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information") ' Get the Inventor Document Summary Information property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties") ' Get the Design Tracking Properties property set.

        '"Имя документа"
        'или partDoc.DisplayName
        writeToList("P1", oPropSetDTP.Item("Part Number").Value, listOfPartAndDrawCriteria, first, last)

        '"Название детали"
        writeToList("P2", oPropSetDTP.Item("Description").Value, listOfPartAndDrawCriteria, first, last)

        '"Материал"
        writeToList("P3", oPropSetDTP.Item("Material").Value, listOfPartAndDrawCriteria, first, last)

        '"Дата создания фаила"
        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        writeToList("P4", f.DateCreated.ToString, listOfPartAndDrawCriteria, first, last)

        '"Дата изменения фаила"
        writeToList("P5", f.DateLastModified.ToString, listOfPartAndDrawCriteria, first, last)

        '"Деталь твердотельная (не поверхности)"
        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        Dim b As Boolean
        For Each SrfBod In SrfBods
            b = SrfBod.IsSolid '? значение последнего surface body ?
        Next
        writeToList("P6", b, listOfPartAndDrawCriteria, first, last)

        '"Деталь состоит из одного твердого тела"
        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        Dim result As String = "EMPTY VALUE"
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            result = True 'true - да, из одного
        Else
            result = False 'false - нет, не из одного
        End If
        writeToList("P7", result, listOfPartAndDrawCriteria, first, last)

        '"Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы"
        writeToList("P8", isOriginsInvisible(partDoc), listOfPartAndDrawCriteria, first, last) 'записать true - да, невидимый; false - видимый

        '"Все эскизы детали должны быть полностью определены"
        Dim isOk As Boolean = True
        Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
        'пройти по всем эскизам детали
        For Each sketch As Sketch In partDef.Sketches
            'является ли эскиз полностью определенным? если нет, то записывем ошибку
            If sketch.ConstraintStatus <> ConstraintStatusEnum.kFullyConstrainedConstraintStatus Then
                isOk = False
                Exit For
            End If
        Next
        writeToList("P9", isOk, listOfPartAndDrawCriteria, first, last) 'записать true - все эскизы детали полностью определены; false - хотя бы один эскиз детали не полностью определен

        '"Масса детали"
        Dim massProps As MassProperties = partDoc.ComponentDefinition.MassProperties
        writeToList("P10", massProps.Mass, listOfPartAndDrawCriteria, first, last)

        '"Площадь детали"
        writeToList("P11", massProps.Area * 100, listOfPartAndDrawCriteria, first, last)

        '"Объем детали"
        writeToList("P12", massProps.Volume * 1000, listOfPartAndDrawCriteria, first, last)
    End Sub

    ' непосредственное получение данных чертежа из Inventor
    Private Sub getDrawingData(ByVal drawingDoc As DrawingDocument, ByVal listOfPartAndDrawCriteria As List(Of Criteria))
        'получение первого и последнего индекса интервала листа, с которым идет работа 
        Dim first As Integer = listOfPartAndDrawCriteria.Count - 15 '15-число записей в list_of_criteria_for_part_and_drawing.json
        Dim last As Integer = listOfPartAndDrawCriteria.Count - 1 'здесь всегда 1

        'заполнение value-получить данные чертежа
        Dim oSheet As Sheet = drawingDoc.Sheets.Item(1) 'лист чертежа
        Dim oView As DrawingView = oSheet.DrawingViews.Item(1) 'вид листа     

        '"Выбор формата листа"
        Dim result As String = "EMPTY VALUE"
        If oSheet.Size = DrawingSheetSizeEnum.kA4DrawingSheetSize Then
            result = "А4"
        ElseIf oSheet.Size = DrawingSheetSizeEnum.kA3DrawingSheetSize Then
            result = "А3"
        Else
            result = "Другой формат"
        End If
        writeToList("D1", result, listOfPartAndDrawCriteria, first, last)

        '"Выбор масштаба главного вида"
        result = "EMPTY VALUE"
        Dim oPropSets As PropertySets = drawingDoc.PropertySets
        Dim oPropSetGOST As PropertySet = oPropSets.Item("Свойства ГОСТ")
        result = oPropSetGOST.Item("Масштаб").Value
        writeToList("D2", result, listOfPartAndDrawCriteria, first, last)

        '"Заполнение основной надписи"
        result = "EMPTY VALUE"
        Dim author As String = Nothing
        Dim designation As String = Nothing
        Dim header As String = Nothing

        Dim oTitleBlock As TitleBlock = oSheet.TitleBlock
        For Each tb As Inventor.TextBox In oTitleBlock.Definition.Sketch.TextBoxes
            If tb.Text = "<АВТОР>" Then
                author = oTitleBlock.GetResultText(tb)
            End If
            If tb.Text = "<ОБОЗНАЧЕНИЕ>" Then
                designation = oTitleBlock.GetResultText(tb)
            End If
            If tb.Text = "<ЗАГОЛОВОК>" Then
                header = oTitleBlock.GetResultText(tb)
            End If
        Next
        'если одна из строк пустая - ошибка, основная надпись не заполнена
        If (String.IsNullOrEmpty(author) Or String.IsNullOrEmpty(designation) Or String.IsNullOrEmpty(header)) Then
            result = False
        Else
            result = True
        End If
        writeToList("D3", result, listOfPartAndDrawCriteria, first, last)
    End Sub

    'вспомогательная функция запись данных из Inventor в List(Of Criteria)
    Private Sub writeToList(ByVal id As String, ByVal value As String, ByVal listOfCriteria As List(Of Criteria), ByVal first As Integer, ByVal last As Integer)
        'пройти по определенному участку листа критериев
        For i As Integer = first To last
            'если найден совпадающий id критерия, внести запись
            If id = listOfCriteria(i).id Then
                listOfCriteria(i).value = value
                Exit For
            End If
        Next
    End Sub

    'сравнение полученных четырех списков
    Private Sub compareLists()
        _listOfResults.Clear() 'сначала необходимо очистить список результатов

        'сначала идет работа со сборками
        compare(_listOfStndAsmCriteria, _listOfChekAsmCriteria)

        'потом идет работа с деталями
        compare(_listOfStndPartAndDrawCriteria, _listOfChekPartAndDrawCriteria)
    End Sub

    'вспомогательная функция сравнение значений конкретных списков
    Private Sub compare(ByVal listOfStndCriteria As List(Of Criteria), ByVal listOfChekCriteria As List(Of Criteria))
        For i = 0 To listOfStndCriteria.Count - 1
            'переписать уже имеющиеся значения списков в список результатов
            _listOfResults.Add(New Result)
            _listOfResults.Last.name = listOfStndCriteria(i).name
            _listOfResults.Last.weight = listOfStndCriteria(i).weight
            _listOfResults.Last.tolerance = listOfStndCriteria(i).tolerance
            _listOfResults.Last.standard_value = listOfStndCriteria(i).value
            _listOfResults.Last.checked_value = listOfChekCriteria(i).value

            'осьалось заполнить поля delta и is_correct

            'являются ли сравниваемые значения - числами?
            Dim stand_value As Double
            Dim check_value As Double
            If (Double.TryParse(_listOfResults.Last.standard_value, stand_value) And Double.TryParse(_listOfResults.Last.checked_value, check_value)) Then
                'являются

                'равны ли проверяемое и эталонные значения?
                If _listOfResults.Last.standard_value = _listOfResults.Last.checked_value Then
                    'если равны, поле разницы величин (delta) = 0, а корректность (is_correct) = 2 (полное совпадение)
                    _listOfResults.Last.delta = 0
                    _listOfResults.Last.is_correct = 2
                Else
                    'величины не равны. необходимо определить, какова разница между ними в % ?
                    Dim a As Double = _listOfResults.Last.standard_value
                    Dim b As Double = _listOfResults.Last.checked_value

                    'определить разницу и заполнить поле delta
                    If (a < b) Then
                        _listOfResults.Last.delta = ((b - a) / a) * 100
                    Else '(a > b)
                        _listOfResults.Last.delta = ((a - b) / a) * 100
                    End If

                    'теперь необходимо сравнить delta и tolerance.
                    If _listOfResults.Last.delta <= _listOfResults.Last.tolerance Then
                        'если delta <= tolerance, ответ правильный, но в пределах отклонения, is_correct = 1
                        _listOfResults.Last.is_correct = 1
                    Else
                        'в противном случае ответ не верный is_correct = 0
                        _listOfResults.Last.is_correct = 0
                    End If
                End If
            Else
                'не являются

                'равны ли проверяемое и эталонные значения?
                If _listOfResults.Last.standard_value = _listOfResults.Last.checked_value Then
                    'равны
                    _listOfResults.Last.delta = 0
                    _listOfResults.Last.is_correct = 2
                Else
                    'не равны
                    _listOfResults.Last.delta = 9999
                    _listOfResults.Last.is_correct = 0
                End If
            End If
        Next
    End Sub

    'вспомогательная функция создание и заполнение html-табоицы и прочий вывод
    Private Sub outputResult()
        lblCountOfRows.Text = "0" 'очистка элемента вывода количества строк

        tableOfResults.InnerHtml = "" 'очистка таблицы
        'добавление кода будет проводиться к div-у с ID "tableOfResults"
        tableOfResults.InnerHtml += "<table border=1 bordercolor=#0D47A1>" 'открывающий тэг таблицы

        'add table columns header
        tableOfResults.InnerHtml += "<thead style='background: #0D47A1; color:  white; font-weight: bold;'>"
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
                tableOfResults.InnerHtml += "<tr style='background: #EF9A9A'>"
            End If
            'если ответ правильный в пределах диапазона - выделить светло-зеленым
            If _listOfResults(i).is_correct = 1 Then
                tableOfResults.InnerHtml += "<tr style='background: #AED581'>"
            End If
            'если ответ правильный - выделить зеленым
            If _listOfResults(i).is_correct = 2 Then
                tableOfResults.InnerHtml += "<tr style='background: #81C784'>"
            End If

            tableOfResults.InnerHtml += "<td>" & _listOfResults(i).name & "</td>"
            tableOfResults.InnerHtml += "<td>" & _listOfResults(i).weight & "</td>"
            tableOfResults.InnerHtml += "<td>" & _listOfResults(i).tolerance & "</td>"
            tableOfResults.InnerHtml += "<td>" & _listOfResults(i).standard_value & "</td>"
            tableOfResults.InnerHtml += "<td>" & _listOfResults(i).checked_value & "</td>"
            tableOfResults.InnerHtml += "<td>" & _listOfResults(i).delta & "</td>"

            tableOfResults.InnerHtml += "</tr>"
        Next

        tableOfResults.InnerHtml += "</table>" 'закрвыающий тэг таблицы

        lblCountOfRows.Text = _listOfResults.Count.ToString() 'заполнение элемента вывода количества строк

        'вывод сообщения об окончании работы и результата
        Dim right As Integer = 0
        Dim wrong As Integer = 0
        Dim total_ball As Integer = 0
        Dim possible_ball As Integer = 0
        Dim percent As Double = 0
        Dim message As String = ""

        For i = 0 To _listOfResults.Count - 1
            possible_ball += _listOfResults(i).weight
            If (_listOfResults(i).is_correct = 1 Or _listOfResults(i).is_correct = 2) Then
                right += 1
                total_ball += _listOfResults(i).weight
            Else
                wrong += 1
            End If
        Next

        percent = total_ball / possible_ball

        message &= "Работа завершена. Ниже приведены результаты сравнения." & vbCrLf & vbCrLf
        message &= "Всего было сравнено параметров: " & _listOfResults.Count & vbCrLf
        message &= "Правильных параметров: " & right & vbCrLf
        message &= "Неправильных параметров: " & wrong & vbCrLf
        message &= "Набрано баллов (из возможных): " & total_ball & " / " & possible_ball & vbCrLf
        message &= "Баллы в процентах: " & percent
        MsgBox(message)
    End Sub

    'вспомогательная функция найти чертеж к документу: сборке (assembly) или детали (part). если чертеж не найден, возвращает пустую строку: ""
    Private Function findDrawingFullFileNameForDocument(ByVal doc As Document) As String
        Try
            Dim fullFilename As String = doc.FullFileName

            'переменная drawingFilename будет хранить полное имя чертежа для сборки / детали
            Dim drawingFilename As String = ""

            ' Extract the path from the full filename.
            Dim path As String = Microsoft.VisualBasic.Left$(fullFilename, InStrRev(fullFilename, "\"))

            ' Extract the filename from the full filename.
            Dim filename As String = Microsoft.VisualBasic.Right$(fullFilename, Len(fullFilename) - InStrRev(fullFilename, "\"))

            ' Replace the extension with "dwg"
            filename = Microsoft.VisualBasic.Left$(filename, InStrRev(filename, ".")) & "dwg"
            ' Find if the drawing exists.
            drawingFilename = _invApp.DesignProjectManager.ResolveFile(path, filename)

            ' Check the result.
            If drawingFilename = "" Then
                ' Try again with idw extension.
                filename = Microsoft.VisualBasic.Left$(filename, InStrRev(filename, ".")) & "idw"
                ' Find if the drawing exists.
                drawingFilename = _invApp.DesignProjectManager.ResolveFile(path, filename)
            End If

            ' Return result. Если не найден чертеж, вернет пустую строку ""
            Return drawingFilename
        Catch ex As Exception
            MsgBox("Ошибка: невозможно найти чертеж для документа" & vbCrLf & ex.ToString)
            Return ""
        End Try
    End Function

    'вспомогательная функция: проверить видимость 2d эскизов и объектов вспомогательной геометрии (плоскости, оси, точки). true - они все невидимы, false - есть как минимум 1 видимый объект
    Private Function isOriginsInvisible(ByVal oDoc As Document) As Boolean
        Dim isInvisible As Boolean = True

        ' получть все 2d эскизы детали и проверить их видимость
        Dim oSketches As PlanarSketches = oDoc.ComponentDefinition.Sketches
        For Each oSketch In oSketches
            If oSketch.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkPlanes collection (все плоскости документа)
        For Each oWorkPlane In oDoc.ComponentDefinition.WorkPlanes
            If oWorkPlane.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkAxes collection (все оси документа)
        For Each oWorkAxe In oDoc.ComponentDefinition.WorkAxes
            If oWorkAxe.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkPoints collection (все точки документа)
        For Each oWorkPoint In oDoc.ComponentDefinition.WorkPoints
            If oWorkPoint.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkSurfaces collection (все поверхности(?) документа)
        'For Each oWorkSurface In oDoc.ComponentDefinition.WorkSurfaces
        '    If oWorkSurface.Visible = False Then
        '        MsgBox(oWorkSurface.Name & " Visible false: ok")
        '    Else
        '        MsgBox(oWorkSurface.Name & "Visible true: not ok")
        '    End If
        'Next        
        Return isInvisible
    End Function
End Class
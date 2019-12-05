Option Explicit On

Imports System.Web.Script.Serialization
Imports Inventor
Imports Newtonsoft.Json

Module CommonFunctionsModule
    'ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ
    Public _listOfStndAsmCriteria As New List(Of Criteria)() 'список для хранения данных самой стандартной сборки
    Public _listOfStndPartAndDrawCriteria As New List(Of Criteria)() 'список для хранения данных деталей и чертежей стандартной сборки
    Public _listOfChekAsmCriteria As New List(Of Criteria)() 'список для хранения данных самой проверяемой сборки
    Public _listOfChekPartAndDrawCriteria As New List(Of Criteria)() 'список для хранения данных деталей и чертежей проверяемой сборки

    Public _stndDocumentSaveLocation As String = "" 'путь хранения документа стандартной сборки
    Public _chekDocumentSaveLocation As String = "" 'путь хранения документа проверяемой сборки
    Public _isAppAutoStarted As Boolean = False 'был ли данный сеанс Inventor создан программой
    Public _invApp As Inventor.Application = Nothing 'приложение Inventor
    Public _name_of_list_of_criteria_for_assembly As String = "list_of_criteria_for_assembly.json" 'имя list_of_criteria_for_assembly.json
    Public _name_of_list_of_criteria_for_part_and_drawing As String = "list_of_criteria_for_part_and_drawing.json" 'имя list_of_criteria_for_part_and_drawing.json
    Public _count_of_list_of_criteria_for_assembly As Integer = 27 'количество записей в list_of_criteria_for_assembly.json
    Public _count_of_list_of_criteria_for_part_and_drawing As Integer = 36 'количество записей в list_of_criteria_for_part_and_drawing.json
    Public _listOfResults As New List(Of Result)() 'список результатов сравнения, выводится в таблицу на веб-форму
    Public _finalMessageString As String = "" 'строка результатов, выводимая в конце работы (сравнения) программы

    'ФУНКЦИИ
    'непосредственное получение данных сборки из Inventor
    Public Sub getAsmData(ByVal asmDoc As AssemblyDocument, ByVal listOfAsmCriteria As List(Of Criteria))
        'получение первого и последнего индекса интервала листа, с которым идет работа 
        Dim first As Integer = listOfAsmCriteria.Count - _count_of_list_of_criteria_for_assembly
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

        '"Сколько раз применены операции"
        Dim features As Features = asmDoc.ComponentDefinition.Features
        writeToList("A12", features.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция Chamfer (фаска)"
        writeToList("A13", features.ChamferFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция CircularPattern (круговой массив)"
        writeToList("A14", features.CircularPatternFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция Client"
        writeToList("A15", features.ClientFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция Extrude (выдавливание)"
        writeToList("A16", features.ExtrudeFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция FaceOffset"
        writeToList("A17", features.FaceOffsetFeatures._Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция Fillet"
        writeToList("A18", features.FilletFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция Hole (отверстие)"
        writeToList("A19", features.HoleFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция MidSurface"
        writeToList("A20", features.MidSurfaceFeatures._Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция Mirror (отражение)"
        writeToList("A21", features.MirrorFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция MoveFace"
        writeToList("A22", features.MoveFaceFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция RectangularPattern (прямоугольный массив)"
        writeToList("A23", features.RectangularPatternFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция Revolve (вращение)"
        writeToList("A24", features.RevolveFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция SketchDrivenPattern"
        writeToList("A25", features.SketchDrivenPatternFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция Sweep"
        writeToList("A26", features.SweepFeatures.Count, listOfAsmCriteria, first, last)

        '"Сколько раз применена операция Thread (резьба)"
        writeToList("A27", features.ThreadFeatures.Count, listOfAsmCriteria, first, last)
    End Sub

    'непосредственное получение данных детали и ее чертежа из Inventor
    Public Sub getPartAndDrawingData(ByVal partDoc As PartDocument, ByVal listOfPartAndDrawCriteria As List(Of Criteria))
        'РАБОТА С ДЕТАЛЬЮ (PART)
        'получение первого и последнего индекса интервала листа, с которым идет работа 
        Dim first As Integer = listOfPartAndDrawCriteria.Count - _count_of_list_of_criteria_for_part_and_drawing
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

        '"Дата создания файла"
        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        writeToList("P4", f.DateCreated.ToString, listOfPartAndDrawCriteria, first, last)

        '"Дата изменения файла"
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

        '"Сколько раз применены операции"
        Dim oDef As PartComponentDefinition = partDoc.ComponentDefinition
        Dim features As PartFeatures = oDef.Features
        writeToList("P13", features.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция Chamfer (фаска)"
        writeToList("P14", features.ChamferFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция CircularPattern (круговой массив)"
        writeToList("P15", features.CircularPatternFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция Client"
        writeToList("P16", features.ClientFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция Extrude (выдавливание)"
        writeToList("P17", features.ExtrudeFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция FaceOffset"
        writeToList("P18", features.FaceOffsetFeatures._Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция Fillet"
        writeToList("P19", features.FilletFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция Hole (отверстие)"
        writeToList("P20", features.HoleFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция MidSurface"
        writeToList("P21", features.MidSurfaceFeatures._Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция Mirror (отражение)"
        writeToList("P22", features.MirrorFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция MoveFace"
        writeToList("P23", features.MoveFaceFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция RectangularPattern (прямоугольный массив)"
        writeToList("P24", features.RectangularPatternFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция Revolve (вращение)"
        writeToList("P25", features.RevolveFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция SketchDrivenPattern"
        writeToList("P26", features.SketchDrivenPatternFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция Sweep"
        writeToList("P27", features.SweepFeatures.Count, listOfPartAndDrawCriteria, first, last)

        '"Сколько раз применена операция Thread (резьба)"
        writeToList("P28", features.ThreadFeatures.Count, listOfPartAndDrawCriteria, first, last)

        'РАБОТА С ЧЕРТЕЖОМ (DRAWING)
        'для текущей детали определить, есть ли для нее чертеж. если есть, получить данные и чертежа
        Dim drawingFullFileName As String = findDrawingFullFileNameForDocument(partDoc, _invApp) 'найти, если возможно, путь к чертежу детали

        'если путь к чертежу не найден, выход из функции
        If String.IsNullOrEmpty(drawingFullFileName) Then
            Exit Sub
        End If

        'если чертеж найден:
        Dim drawingDoc As DrawingDocument = _invApp.Documents.Open(drawingFullFileName) 'открыть чертеж

        'заполнение value-получить данные чертежа
        Dim oSheet As Sheet = drawingDoc.Sheets.Item(1) 'лист чертежа
        Dim oView As DrawingView = oSheet.DrawingViews.Item(1) 'вид листа     

        '"Имя документа-чертежа"
        writeToList("D1", drawingDoc.DisplayName, listOfPartAndDrawCriteria, first, last)

        '"Выбор формата листа"
        result = "EMPTY VALUE"
        If oSheet.Size = DrawingSheetSizeEnum.kA4DrawingSheetSize Then
            result = "А4"
        ElseIf oSheet.Size = DrawingSheetSizeEnum.kA3DrawingSheetSize Then
            result = "А3"
        Else
            result = "Другой формат"
        End If
        writeToList("D2", result, listOfPartAndDrawCriteria, first, last)

        '"Выбор масштаба главного вида"
        result = "EMPTY VALUE"
        oPropSets = drawingDoc.PropertySets
        Dim oPropSetGOST As PropertySet = oPropSets.Item("Свойства ГОСТ")
        result = oPropSetGOST.Item("Масштаб").Value
        writeToList("D3", result, listOfPartAndDrawCriteria, first, last)

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
        writeToList("D4", result, listOfPartAndDrawCriteria, first, last)

        '"Единицы измерения углов"
        Select Case drawingDoc.UnitsOfMeasure.AngleUnits
            Case UnitsTypeEnum.kDegreeAngleUnits
                result = "Degree"
            Case UnitsTypeEnum.kGradAngleUnits
                result = "Grad"
            Case UnitsTypeEnum.kRadianAngleUnits
                result = "Radian"
            Case UnitsTypeEnum.kSteradianAngleUnits
                result = "Steradian"
            Case Else
                result = "Other"
        End Select
        writeToList("D5", result, listOfPartAndDrawCriteria, first, last)

        '"Единицы измерения длины"
        Select Case drawingDoc.UnitsOfMeasure.LengthUnits
            Case UnitsTypeEnum.kCentimeterLengthUnits
                result = "Centimeter"
            Case UnitsTypeEnum.kFootLengthUnits
                result = "Foot"
            Case UnitsTypeEnum.kInchLengthUnits
                result = "Inch"
            Case UnitsTypeEnum.kMeterLengthUnits
                result = "Meter"
            Case UnitsTypeEnum.kMicronLengthUnits
                result = "Micron"
            Case UnitsTypeEnum.kMileLengthUnits
                result = "Mile"
            Case UnitsTypeEnum.kMilLengthUnits
                result = "Mil"
            Case UnitsTypeEnum.kMillimeterLengthUnits
                result = "Millimeter"
            Case UnitsTypeEnum.kNauticalMileLengthUnits
                result = "NauticalMile"
            Case UnitsTypeEnum.kYardLengthUnits
                result = "Yard"
            Case Else
                result = "Other"
        End Select
        writeToList("D6", result, listOfPartAndDrawCriteria, first, last)

        '"Единицы измерения массы"
        Select Case drawingDoc.UnitsOfMeasure.MassUnits
            Case UnitsTypeEnum.kGramMassUnits
                result = "Gram"
            Case UnitsTypeEnum.kKilogramMassUnits
                result = "Kilogram"
            Case UnitsTypeEnum.kLbMassMassUnits
                result = "LbMass"
            Case UnitsTypeEnum.kOunceMassUnits
                result = "Ounce"
            Case UnitsTypeEnum.kSlugMassUnits
                result = "Slug"
            Case Else
                result = "Other"
        End Select
        writeToList("D7", result, listOfPartAndDrawCriteria, first, last)

        '"Единицы измерения времени"
        Select Case drawingDoc.UnitsOfMeasure.TimeUnits
            Case UnitsTypeEnum.kHourTimeUnits
                result = "Hour"
            Case UnitsTypeEnum.kMinuteTimeUnits
                result = "Minute"
            Case UnitsTypeEnum.kSecondTimeUnits
                result = "Second"
            Case Else
                result = "Other"
        End Select
        writeToList("D8", result, listOfPartAndDrawCriteria, first, last)

    End Sub

    'ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ

    'вспомогательная функция сравнение значений конкретных списков
    Public Sub compare(ByVal listOfStndCriteria As List(Of Criteria), ByVal listOfChekCriteria As List(Of Criteria))
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

    'вспомогательная функция десериализация из json файла в лист
    Public Sub deserialization(ByVal listOfCriteriaJsonName As String, ByVal listOfCriteria As List(Of Criteria), ByVal serverMapPath As String)
        Dim jsonString = System.IO.File.ReadAllText(serverMapPath & "\" & listOfCriteriaJsonName)
        Dim values = JsonConvert.DeserializeObject(Of List(Of Criteria))(jsonString)
        For Each value In values
            listOfCriteria.Add(New Criteria())
            listOfCriteria.Last.id = value.id
            listOfCriteria.Last.name = value.name
            listOfCriteria.Last.weight = value.weight
            listOfCriteria.Last.tolerance = value.tolerance
            listOfCriteria.Last.value = value.value
        Next
    End Sub

    'вспомогательная функция создание и вывод итоговой информации по проверке
    Public Sub finalMessage()
        'очистка строки результатов
        _finalMessageString = ""

        'вывод сообщения об окончании работы и результата
        Dim total_parameters = _listOfResults.Count
        Dim right As Integer = 0
        Dim wrong As Integer = 0
        Dim total_ball As Integer = 0
        Dim possible_ball As Integer = 0
        Dim percent As Double = 0

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

        _finalMessageString &= "Всего было сравнено параметров: " & total_parameters & vbCrLf
        _finalMessageString &= "Параметров совпало: " & right & vbCrLf
        _finalMessageString &= "Параметров не совпало: " & wrong & vbCrLf
        _finalMessageString &= "Набрано баллов (из возможных): " & total_ball & " / " & possible_ball & vbCrLf
        _finalMessageString &= "Баллы в процентах: " & (Math.Round(percent, 4) * 100) & "%"
        MsgBox(_finalMessageString, vbSystemModal, "Результаты сравнения")
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

    'вспомогательная функция найти чертеж к документу: сборке (assembly) или детали (part). если чертеж не найден, возвращает пустую строку: ""
    Private Function findDrawingFullFileNameForDocument(ByVal doc As Document, ByVal _invApp As Inventor.Application) As String
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
            MsgBox("Ошибка: невозможно найти чертеж для документа" & vbCrLf & ex.ToString, vbSystemModal, "Ошибка")
            Return ""
        End Try
    End Function

End Module

Attribute VB_Name = "Module1"
Sub CombineDataCreateTableAndSum()
    Dim wsTarget As Worksheet, wbSource As Workbook
    Dim lastRowTarget As Long, lastRowSource As Long, lastColSource As Long
    Dim folderPath As String, currentDayFolder As String, targetFileName As String
    Dim file As Object, folder As Object, fs As Object
    Dim sum As Double, sourceRange As Range, tbl As ListObject
    Dim rng As Range, targetCell As Range
    Dim sumRow As Long
    
    ' Получаем путь к папке из ячейки A1 на листе "Sheet1"
    folderPath = ThisWorkbook.Sheets("Sheet1").Range("A1").Value
    ' Создаем папку с именем текущей даты
    currentDayFolder = folderPath & "\" & Format(Now(), "yyyy-mm-dd")
    ' Создаем имя файла для сохранения итогового файла
    targetFileName = currentDayFolder & "\total.xlsx"
    
    ' Создаем папку с текущей датой
    On Error Resume Next
    MkDir currentDayFolder
    On Error GoTo 0
    
    ' Создаем новую рабочую книгу и лист для целевых данных
    Set wsTarget = Workbooks.Add.Sheets(1)
    
    ' Инициализируем объект файловой системы для перебора файлов в папке
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set folder = fs.GetFolder(folderPath)
    
    ' Перебираем файлы в папке и копируем данные в wsTarget
    For Each file In folder.Files
        If file.Name Like "*.xlsx" And Not file.Name Like "total.xlsx" Then
            ' Открываем исходный файл
            Set wbSource = Workbooks.Open(file.Path)
            With wbSource.Sheets(1)
                ' Определяем диапазон данных
                lastRowSource = .Cells(.Rows.Count, 1).End(xlUp).row
                lastColSource = .Cells(1, .Columns.Count).End(xlToLeft).Column
                Set sourceRange = .Range(.Cells(1, 1), .Cells(lastRowSource, lastColSource))
                ' Копируем шапку таблицы если это первый файл
                If wsTarget.Range("A1").Value = "" Then
                    sourceRange.Rows(1).Copy Destination:=wsTarget.Rows(1)
                End If
                ' Копируем данные
                lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).row + 1
                sourceRange.Offset(1, 0).Resize(lastRowSource - 1).Copy Destination:=wsTarget.Cells(lastRowTarget, 1)
            End With
            ' Закрываем исходный файл
            wbSource.Close False
        End If
    Next file
    
    ' Очищаем буфер обмена
    Application.CutCopyMode = False
    
    ' Обновляем последнюю строку на целевом листе
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).row
    
    ' Создаем таблицу на основе скопированных данных
    Set rng = wsTarget.Range("A1").CurrentRegion ' Текущий диапазон региона включает все данные
    Set tbl = wsTarget.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "TotalTable"
    tbl.TableStyle = "TableStyleLight9"
    tbl.HeaderRowRange.Font.Color = RGB(0, 0, 0)

    
    ' Вычисляем сумму по первому столбцу
    sumRow = lastRowTarget + 1 ' Строка для суммы расположена сразу под последней строкой данных
    sum = Application.WorksheetFunction.sum(wsTarget.Range(wsTarget.Cells(2, 1), wsTarget.Cells(lastRowTarget, 1)))
    Set targetCell = wsTarget.Cells(sumRow, 1)
    targetCell.Value = "Total Sum"
    targetCell.Offset(0, 1).Value = sum
    
    ' Применяем форматирование к ячейкам суммы
    targetCell.Font.Bold = True
    targetCell.Offset(0, 1).Font.Bold = True
    
    ' Заменяем точки на запятые во всех строках с данными, кроме строки с суммой
    Dim cell As Range
    For Each cell In tbl.DataBodyRange
        If Not IsNumeric(cell.Value) And InStr(cell.Value, ".") > 0 Then
            cell.Value = Replace(cell.Value, ".", ",")
        End If
    Next cell
    
    ' Сохраняем итоговую рабочую книгу
    wsTarget.Parent.SaveAs filename:=targetFileName
    wsTarget.Parent.Close SaveChanges:=False
End Sub


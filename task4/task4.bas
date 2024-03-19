Attribute VB_Name = "Module1"
Sub CombineDataCreateTableAndSum()
    Dim wsTarget As Worksheet, wbSource As Workbook
    Dim lastRowTarget As Long, lastRowSource As Long, lastColSource As Long
    Dim folderPath As String, currentDayFolder As String, targetFileName As String
    Dim file As Object, folder As Object, fs As Object
    Dim sum As Double, sourceRange As Range, tbl As ListObject
    Dim rng As Range, targetCell As Range
    Dim sumRow As Long
    
    ' �������� ���� � ����� �� ������ A1 �� ����� "Sheet1"
    folderPath = ThisWorkbook.Sheets("Sheet1").Range("A1").Value
    ' ������� ����� � ������ ������� ����
    currentDayFolder = folderPath & "\" & Format(Now(), "yyyy-mm-dd")
    ' ������� ��� ����� ��� ���������� ��������� �����
    targetFileName = currentDayFolder & "\total.xlsx"
    
    ' ������� ����� � ������� �����
    On Error Resume Next
    MkDir currentDayFolder
    On Error GoTo 0
    
    ' ������� ����� ������� ����� � ���� ��� ������� ������
    Set wsTarget = Workbooks.Add.Sheets(1)
    
    ' �������������� ������ �������� ������� ��� �������� ������ � �����
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set folder = fs.GetFolder(folderPath)
    
    ' ���������� ����� � ����� � �������� ������ � wsTarget
    For Each file In folder.Files
        If file.Name Like "*.xlsx" And Not file.Name Like "total.xlsx" Then
            ' ��������� �������� ����
            Set wbSource = Workbooks.Open(file.Path)
            With wbSource.Sheets(1)
                ' ���������� �������� ������
                lastRowSource = .Cells(.Rows.Count, 1).End(xlUp).row
                lastColSource = .Cells(1, .Columns.Count).End(xlToLeft).Column
                Set sourceRange = .Range(.Cells(1, 1), .Cells(lastRowSource, lastColSource))
                ' �������� ����� ������� ���� ��� ������ ����
                If wsTarget.Range("A1").Value = "" Then
                    sourceRange.Rows(1).Copy Destination:=wsTarget.Rows(1)
                End If
                ' �������� ������
                lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).row + 1
                sourceRange.Offset(1, 0).Resize(lastRowSource - 1).Copy Destination:=wsTarget.Cells(lastRowTarget, 1)
            End With
            ' ��������� �������� ����
            wbSource.Close False
        End If
    Next file
    
    ' ������� ����� ������
    Application.CutCopyMode = False
    
    ' ��������� ��������� ������ �� ������� �����
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).row
    
    ' ������� ������� �� ������ ������������� ������
    Set rng = wsTarget.Range("A1").CurrentRegion ' ������� �������� ������� �������� ��� ������
    Set tbl = wsTarget.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "TotalTable"
    tbl.TableStyle = "TableStyleLight9"
    tbl.HeaderRowRange.Font.Color = RGB(0, 0, 0)

    
    ' ��������� ����� �� ������� �������
    sumRow = lastRowTarget + 1 ' ������ ��� ����� ����������� ����� ��� ��������� ������� ������
    sum = Application.WorksheetFunction.sum(wsTarget.Range(wsTarget.Cells(2, 1), wsTarget.Cells(lastRowTarget, 1)))
    Set targetCell = wsTarget.Cells(sumRow, 1)
    targetCell.Value = "Total Sum"
    targetCell.Offset(0, 1).Value = sum
    
    ' ��������� �������������� � ������� �����
    targetCell.Font.Bold = True
    targetCell.Offset(0, 1).Font.Bold = True
    
    ' �������� ����� �� ������� �� ���� ������� � �������, ����� ������ � ������
    Dim cell As Range
    For Each cell In tbl.DataBodyRange
        If Not IsNumeric(cell.Value) And InStr(cell.Value, ".") > 0 Then
            cell.Value = Replace(cell.Value, ".", ",")
        End If
    Next cell
    
    ' ��������� �������� ������� �����
    wsTarget.Parent.SaveAs filename:=targetFileName
    wsTarget.Parent.Close SaveChanges:=False
End Sub


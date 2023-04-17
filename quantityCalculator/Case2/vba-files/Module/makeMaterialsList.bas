Attribute VB_Name = "makeMaterialsList"
'Option Explicit

'�C�X�U�u�{�ϥΪ�����Χ��� (TODO: �������ө�)
Sub makeMaterialsListByWork()
    '�إߧ��Ʋξ��
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "���ƾ�z"
    If Not isSheetExist(targetSheetName) Then
        Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        targetSheet.Name = targetSheetName
    Else
        Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    End If
    
    '��X"����"cell
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("����")

    workNo = 1
    isLastWork = False
    '���гB�z�U�Ӥu�{
    Do Until isLastWork
        Dim searchRange As Range
        Set searchRange = getWorkRange(workNo)
        If searchRange Is Nothing Then
            Exit Sub
        End If
        
        '�T�{��������ƽd��O�b����column��m
        Dim rangeDiff As Integer: rangeDiff = itemTitleCell.Column - searchRange.Column
        
        '�L�o����
        Dim noBackgroundColorRange As Range
        Dim notEmptyRange As Range
        Set noBackgroundColorRange = filterCells(searchRange.Offset(0, rangeDiff), xlFilterNoFill)
        Set notEmptyRange = filterCells(searchRange.Offset(0, rangeDiff), xlFilterValues)
        
        '�Q�Υ涰�A���o�Ҧ�����
        Dim materialsRange As Range
        Set materialsRange = Intersect(noBackgroundColorRange, notEmptyRange)
        
        '�N���ƽƻs����ƪ�A�ò������ƪ�����
        If Not materialsRange Is Nothing Then
            Dim copiedMaterialsRange As Range
            Set copiedMaterialsRange = copyRangeToSheet(materialsRange, targetSheet, 1)
            removeDuplicatesInColumn (copiedMaterialsRange)
        End If
    
    Loop
    
    
End Sub

'���o���w�u�{������d��
Function getWorkRange(workNo As Integer) As Range
    Dim workStartCell As Range
    Dim workEndCell As Range
    Set workStartCell = findCellByValue(num2Tc(workNo))
    Set workEndCell = findCellByValue(num2Tc(workNo + 1))
    
    If workStartCell Is Nothing Then
        Set getWorkRange = workStartCell
        Exit Function
    End If
    
    If workEndCell Is Nothing Then
        Dim lastRowCell As Range
        Set lastRowCell = getLastRow()
        Set getWorkRange = Range(workStartCell, lastRowCell)
    Else
        Set getWorkRange = Range(workStartCell, workEndCell.Offset(-1, 0))
    End If
End Function

'�򥻹L�o���
'XlAutoFilterOperator: xlFilterNoFill(�L�o�թ�)�BxlFilterValues(�L�o�D�ťո��)
Function filterCells(searchRange As Range, Optional filterOperator As XlAutoFilterOperator = xlFilterNoFill) As Range
    searchRange.AutoFilter Field:=1, Criteria1:="<>", Operator:=filterOperator
    Set filterCells = ActiveSheet.AutoFilter.Range.Offset(0, 0).SpecialCells(xlCellTypeVisible)  '�Q��SpecialCells�N�L�o�᪺����নrange
    ActiveSheet.AutoFilterMode = False  '�����L�o(excel�����L�o�Ϯ׷|����)
End Function

'�Nrange�ƻs����w�u�@���S�w��m�A�æ^�ǥؼФu�@�����o�ƻs���d��
Function copyRangeToSheet(sourceRange As Range, targetSheet As Worksheet, targetColumn As Integer) As Range
    sourceRange.Copy Destination:=targetSheet.Cells(1, targetColumn)
    Dim copiedRange As Range
    Set copiedRange = Intersect(targetSheet.UsedRange, targetSheet.Columns(targetColumn))
    Set copyRangeToSheet = copiedRange
End Function

'�������w�d�򪺭��Ƹ��(�Hcolumn���@���)
Function removeDuplicatesInColumn(targetRange As Range)
    targetRange.RemoveDuplicates Columns:=1, Header:=xlNo
End Function

Sub makeMaterialsListByWorkUnitTest()
    'Dim workStartCell As Range
    'Set workStartCell = getWorkRange(3)
    'workStartCell.Select
    
    'Dim lastRowCell As Variant
    'Set lastRowCell = getLastRow()
    'lastRowCell.Select
    
    Dim workRange As Range
    Set workRange = getWorkRange(1)
    'workRange.Offset(0, 1).Select
    'Debug.Print workRange.Column
    'Debug.Print "xxx"
    
    Dim noBackgroundColorRange As Range
    Set noBackgroundColorRange = filterCells(workRange.Offset(0, 1), xlFilterValues)
    
    Dim materialsRange As Range
    Set materialsRange = noBackgroundColorRange

    '�N���ƪ����ƥh��
    If Not materialsRange Is Nothing Then
        '���ƥ��ܧO�B
        Dim tempSheet As Worksheet
        Set tempSheet = ThisWorkbook.Sheets.Add
        materialsRange.Copy Destination:=tempSheet.Range("A1")
        
        '��������
        Dim copiedMaterialsRange As Variant
        Set copiedMaterialsRange = tempSheet.Range("A1").SpecialCells(xlCellTypeConstants)
        copiedMaterialsRange.RemoveDuplicates Columns:=1, Header:=xlNo
    End If
    
End Sub

Sub makeMaterialsListByWorkUnitTest2()
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "���ƾ�z"
    If Not isSheetExist(targetSheetName) Then
        Set targetSheet = ThisWorkbook.Sheets.Add
        targetSheet.Name = targetSheetName
    Else
        Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    End If
End Sub

Sub teettt()
    Dim workElementsSheet As Worksheet
    Set workElementsSheet = activateAndSelectSheet("��3_����ƶq�p���")
    
    '�إߧ��Ʋξ��
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "���ƾ�z"
    If Not isSheetExist(targetSheetName) Then
        Set targetSheet = ThisWorkbook.Sheets.Add
        targetSheet.Name = targetSheetName
        workElementsSheet.Activate
    Else
        Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    End If
    
    '��X"����"cell
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("����")

    Dim searchRange As Range
    Set searchRange = getWorkRange(1)
    If searchRange Is Nothing Then
        Exit Sub
    End If
    
    '�T�{��������ƽd��O�b����column��m
    Dim rangeDiff As Integer: rangeDiff = itemTitleCell.Column - searchRange.Column
    
    Dim copiedMaterialsRange As Range
    Set copiedMaterialsRange = copyRangeToSheet(searchRange.Offset(0, rangeDiff), targetSheet, 1)
    targetSheet.Activate
    
    '�L�o����
    Dim noBackgroundColorRange As Range
    Dim notEmptyRange As Range
    Set noBackgroundColorRange = filterCells(copiedMaterialsRange, xlFilterNoFill)
    Set copiedMaterialsRange2 = copyRangeToSheet(noBackgroundColorRange, targetSheet, 3)
    Set notEmptyRange = filterCells(copiedMaterialsRange, xlFilterValues)
    Set copiedMaterialsRange3 = copyRangeToSheet(notEmptyRange, targetSheet, 5)
    
    '�Q�Υ涰�A���o�Ҧ�����
    Dim materialsRange As Range
    Set materialsRange = Intersect(noBackgroundColorRange, notEmptyRange)
    Set copiedMaterialsRange = copyRangeToSheet(materialsRange, targetSheet, 7)
    
    '�N���ƽƻs����ƪ�A�ò������ƪ�����
    If Not materialsRange Is Nothing Then
        'Dim copiedMaterialsRange As Range
        'Set copiedMaterialsRange = copyRangeToSheet(materialsRange, targetSheet, 1)
        removeDuplicatesInColumn copiedMaterialsRange
    End If

End Sub

Sub testSpeciCell()
Cells(1, 1).SpecialCells(xlCellTypeVisible).Select
Cells(1, 1).SpecialCells(xlCellTypeConstants).Select

End Sub

Sub testCopy()
    Dim workElementsSheet As Worksheet
    Set workElementsSheet = activateAndSelectSheet("��3_����ƶq�p���")
    
    '�إߧ��Ʋξ��
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "���ƾ�z"
    If Not isSheetExist(targetSheetName) Then
        Set targetSheet = ThisWorkbook.Sheets.Add
        targetSheet.Name = targetSheetName
        workElementsSheet.Activate
    Else
        Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    End If
    
    '��X"����"cell
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("����")

    Dim searchRange As Range
    Set searchRange = getWorkRange(1)
    If searchRange Is Nothing Then
        Exit Sub
    End If
    
    '�T�{��������ƽd��O�b����column��m
    Dim rangeDiff As Integer: rangeDiff = itemTitleCell.Column - searchRange.Column

    searchRange.Offset(0, rangeDiff).Copy Destination:=targetSheet.Cells(1, 2)
    targetSheet.Activate
    Dim copiedRange As Range
    Set copiedRange = Intersect(targetSheet.UsedRange, targetSheet.Columns(2))
    copiedRange.Select
    Debug.Print targetSheet.UsedRange
    Debug.Print targetSheet.Columns(2)
    
    
End Sub

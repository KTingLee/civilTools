Attribute VB_Name = "makeElementsMaterialsByWork"
'Option Explicit

'�C�X�U�u�{�ϥΪ�����Χ��� (TODO: �������ө�)
Sub makeElementsMaterialsByWork()
    Dim workElementSheetParamCell As Range
    Set workElementSheetParamCell = getParamCell("elementsMaterialSheetParam", "�д��Ѥ���ƶq��W��")
    
    If Not isSheetExist(workElementSheetParamCell.Value) Then
        MsgBox ("�䤣�� " & workElementSheetParamCell.Value)
        Exit Sub
    End If

    Dim workElementsSheet As Worksheet
    Set workElementsSheet = activateAndSelectSheet(workElementSheetParamRange.Value)  '�Ѽ��x�s��

    '�إߧ��Ʋξ��
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "���ƾ�z"
    Set targetSheet = createSheet(targetSheetName)
    workElementsSheet.Activate
    
    '��X"����"cell
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("����")
    If itemTitleCell Is Nothing Then
        MsgBox ("�ʤ�'����'�x�s��A�L�k�w��")
        Exit Sub
    End If

    Dim workNo As Integer: workNo = 1
    Dim isLastWork As Boolean: isLastWork = False
    '���гB�z�U�Ӥu�{
    Do Until isLastWork
        Dim workRange As Range
        Set workRange = getWorkRange(workNo, itemTitleCell)
        If workRange Is Nothing Then
            isLastWork = True
            MsgBox ("�L�u�{" & num2Tc(workNo))
            Exit Sub
        End If
        
        '�����u�{��Ʀ�4��A�Ĥ@��O��Ӥ���t���Ƨ��ơA�ĤG��O�����Ƨ��ơA�ĤT�椸��
        Dim baseColumn As Integer: baseColumn = (workNo - 1) * 4 + 1
        
        Dim copiedWorkRange As Range
        Set copiedWorkRange = copyRangeToSheet(workRange, targetSheet, baseColumn)
        targetSheet.Activate  '������ƾ�z��A�~�వ�L�o
        
        '���o�����u�{�Ҧ����ƨåh������
        Dim materialsList As Range
        Set materialsList = getMaterialsList(copiedWorkRange)
        Set materialsList = copyRangeToSheet(materialsList, targetSheet, baseColumn + 1)  '�O�d�`����(�t���Ƨ���)�A�ñN�h���ƪ����G��b����
        removeDuplicatesInColumn materialsList
        
        '���o�����u�{�Ҧ�����
        Dim elementsList As Range
        Set elementsList = getWorkList(copiedWorkRange)
        Set elementsList = copyRangeToSheet(elementsList, targetSheet, baseColumn + 2)
        removeDuplicatesInColumn elementsList
        
        workNo = workNo + 1
        
        '���^�쥻���u�@��
        workElementsSheet.Activate
    Loop
End Sub

'���o���w�u�{������d��
Function getWorkRange(workNo As Integer, itemTitleCell As Range) As Range
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
    
    Dim rangeDiff As Integer: rangeDiff = itemTitleCell.Column - getWorkRange.Column
    Set getWorkRange = getWorkRange.Offset(0, rangeDiff)
End Function


'�򥻹L�o���
'XlAutoFilterOperator: xlFilterNoFill(�L�o�թ�)�BxlFilterValues(�L�o�D�ťո��)�BxlFilterCellColor(�L�o�C��)
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

'���o�����u�{�ϥΤ����`����(�ǤJ�����u�{�`�d��)
Function getMaterialsList(workElementMaterialsRange As Range) As Range
    '�L�o����
    Dim noBackgroundColorRange As Range
    Dim notEmptyRange As Range
    Set noBackgroundColorRange = filterCells(workElementMaterialsRange, xlFilterNoFill)
    Set notEmptyRange = filterCells(workElementMaterialsRange, xlFilterValues)
    
    '�Q�Υ涰�A���o�Ҧ�����
    Dim materialsRange As Range
    Set materialsRange = Intersect(noBackgroundColorRange, notEmptyRange)
    'removeDuplicatesInColumn materialsRange
    Set getMaterialsList = materialsRange
End Function

'���o�����u�{���Ҧ�����
Function getWorkList(workElementMaterialsRange As Range) As Range
    '�����x�s�榳����A�H�U���L�o�X�L����A�ûP��l�d��@�t��
    Dim noBackgroundColorRange As Range
    Set noBackgroundColorRange = filterCells(workElementMaterialsRange, xlFilterNoFill)
    Set getWorkList = Difference(workElementMaterialsRange, noBackgroundColorRange)
End Function





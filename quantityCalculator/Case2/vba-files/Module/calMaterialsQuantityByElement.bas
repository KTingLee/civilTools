Attribute VB_Name = "calMaterialsQuantityByElement"
Option Explicit

'���o�Ӥu�{�Ҧ����ƲM��
Function getMaterialsInCurrentWork() As Range
    Dim indexCell As Range
    Set indexCell = findCellByValue("�u�{����").Offset(1, 0)  '������w�I
    Set getMaterialsInCurrentWork = Range(indexCell, indexCell.End(xlDown))
End Function

'���o�����u�{������Χ����`�d��(�t���B�p�p�ƶq)
Function getElementsAndMaterialsRangeByWork(workNo As Integer, workElementsSheetName As String) As Variant
    '������l�u�@��A�̫�n���^��
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet

    '���줸��ƶq��
    Dim workElementsSheet As Worksheet
    Set workElementsSheet = activateAndSelectSheet(workElementsSheetName)
    
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("����")  '������A������w�I
    
    Dim quantityCell As Range
    Set quantityCell = findCellByValue("���")  '������A����B���Ƹ�T��w�I
    
    '���o�Ӥu�{���Ҧ�����P����
    Dim workRange As Range
    Set workRange = getWorkRange(workNo, itemTitleCell)
    If workRange Is Nothing Then
        MsgBox ("�нT�{�u�{�s���O�_�s�b")  'TODO: �Ȯɰ��@�ӿ��~�B�z�A���ᦳ�ߤO�b�Ӭݫ������n
        Set getElementsAndMaterialsRangeByWork = workRange
        Exit Function
    End If
    
    Dim rangeDiff As Integer: rangeDiff = quantityCell.Column - workRange.Column
    Set workRange = workRange.Resize(workRange.Rows.Count, workRange.Columns.Count + rangeDiff)  '��V�X�i�ܳ��d��
    
    Set getElementsAndMaterialsRangeByWork = workRange
    currentSheet.Activate
End Function

'�^�Ǥ���Χ��ƪ����
Function getObjectUnitCell(objectName As String, workRange As Range) As Range
    '��X����Χ��ƪ��x�s��(����خy�СA�@�ӬO�bworkRange�A�t�@�ӬO�b��lsheet�����y��)
    Dim objectCell As Range
    Set objectCell = findCellByValueInRange(objectName, workRange)
    
    If objectCell Is Nothing Then
        Set getObjectUnitCell = objectCell  'TODO: �Ȯɰ��@�ӿ��~�B�z�A���ᦳ�ߤO�b�Ӭݫ������n
        Exit Function
    End If
    
    '�z�L����B���ƥ������u�@��h�������A�]��workRange���t�����Y�C
    Dim indexCell As Range
    Set indexCell = findCellByValue("���", objectCell.Worksheet.Name)

    '�����q����B���ƥ������u�@��(�����)����
    Set getObjectUnitCell = objectCell.Worksheet.Cells(objectCell.Row, indexCell.Column)
End Function

'�^�Ǥ���Y�ӧ��Ƽƶq
Function getMaterialQuantityCell(materialName As String, elementName As String, workRange As Range) As Range
    Dim elementCell As Range
    Set elementCell = findCellByValueInRange(elementName, workRange)
    
    '�z�L����B���ƥ������u�@��h�������A�]��workRange���t�����Y�C
    Dim indexCell As Range
    Set indexCell = findCellByValue("�p�p", elementCell.Worksheet.Name)
    
    '�̷Ӥ���A���s�վ�workRange�d��A�קK���e�@�Ӥ��󪺧���
    Dim startCell As Range
    Dim endCell As Range
    Set startCell = workRange.Worksheet.Cells(elementCell.Row, elementCell.Column)  '�o�O�q����u�@��y�Ч�
    'Set endCell = workRange.Cells(workRange.Rows.Count, workRange.Columns.Count)  '�o�O�qworkRange�y�Ч� -> (�o�˥i��]�t�Ӥ���S��������)
    
    Dim searchRange As Range
    Set searchRange = workRange.Worksheet.Range(startCell, startCell.End(xlDown))

    Dim materialCell As Range
    Set materialCell = findCellByValueInRange(materialName, searchRange)
    
    If materialCell Is Nothing Then
        Set getMaterialQuantityCell = materialCell
        Exit Function
    End If
    
    Set getMaterialQuantityCell = searchRange.Worksheet.Cells(materialCell.Row, indexCell.Column)
End Function

'�u�{�ƶq�έp��: �إߤ���P���ƪ����Y��(��ƶq)
Sub calMaterialsQuantityByElement()
    Dim workElementSheetParamCell As Range
    Set workElementSheetParamCell = getParamCell("elementsMaterialSheetParam", "�д��Ѥ���ƶq��W��")
    
    Dim elementsQuantitySheetParamCell As Range
    Set elementsQuantitySheetParamCell = getParamCell("materialsQuantitySheetParam", "�г]�w�ؼФu�{�ƶq�έp��")
    
    Dim workNoParamCell As Range
    Set workNoParamCell = getParamCell("workNoParam", "�г]�w�ؼФu�{�s��")
    
    
    If Not isSheetExist(workElementSheetParamCell.Value) Or Not isSheetExist(elementsQuantitySheetParamCell.Value) Or IsEmpty(workNoParamCell) Then
        MsgBox ("�нT�{�����Τu�{�ƶq�έp��B�u�{�s���O�_���T")
        Exit Sub
    End If

    Dim elementsQuantitySheet As Worksheet
    Set elementsQuantitySheet = activateAndSelectSheet(elementsQuantitySheetParamCell.Value)  '����ؼФu�{�ƶq�έp��

    Dim indexCell As Range
    Set indexCell = findCellByValue("���", elementsQuantitySheet.Name)  '������w�I�A������k���}�l���O����W��
    
    Dim materialsList As Range
    Set materialsList = getMaterialsInCurrentWork
    materialsList.Select
    
    Dim workRange As Range
    Set workRange = getElementsAndMaterialsRangeByWork(workNoParamCell.Value, workElementSheetParamCell.Value)  '�ǤJ�u�{�s���B�����A�H����u�{������ƽd��
    If workRange Is Nothing Then
        Exit Sub  'TODO: �Ȯɰ��@�ӿ��~�B�z�A���ᦳ�ߤO�b�Ӭݫ������n
    End If
    
    '�B�z���Ƴ��
    Dim material As Range
    Dim materialName As String
    Dim unitCell As Range
    For Each material In materialsList
        materialName = material.Value
        Set unitCell = getObjectUnitCell(materialName, workRange)
        
        If unitCell Is Nothing Then
            MsgBox ("�b��e�u�{�s�������G�S���ϥ�: " & materialName)  'TODO: �Ȯɰ��@�ӿ��~�B�z�A���ᦳ�ߤO�b�Ӭݫ������n
            Exit Sub
        End If
        
        elementsQuantitySheet.Cells(material.Row, material.Column + 1).Formula = "=" & unitCell.Worksheet.Name & "!" & unitCell.Address
    Next
    
    '�v����B�z: �g�J������B�p��Ӥ���U�ӧ��ƪ��ƶq
    Dim element As Range
    Set element = Cells(indexCell.Row + 1, indexCell.Column + 1)  'indexCell.row+1 �O���F�O�I�A�קK while �L�k���X���x�s��
    
    Dim elementName As String
    Dim elementQuantityCell As Range
    Dim materialQuantityCell As Range
    Do While Not element.MergeCells  '�`�p���O�X���x�s��A�P�����x�s��榡���P
        '�B�z������
        elementName = element.Value
        Set unitCell = getObjectUnitCell(elementName, workRange)
        
        Set elementQuantityCell = element.Offset(1, 0)
        elementQuantityCell.NumberFormatLocal = "0" & """" & unitCell.Value & """"  'TODO: ���B�z�᭱��X��
        
        '�}�l�p�⤸�󪺧��Ƽƶq
        For Each material In materialsList
        
            materialName = material.Value
            Set materialQuantityCell = getMaterialQuantityCell(materialName, elementName, workRange)
            
            'Debug.Print "����ƶq��}" & elementQuantity.Address
            
            If materialQuantityCell Is Nothing Then
                'Debug.Print elementName & "�S������:  " & materialName
            Else
                elementsQuantitySheet.Cells(material.Row, element.Column).Formula = "=" & elementQuantityCell.Address & "*" & materialQuantityCell.Worksheet.Name & "!" & materialQuantityCell.Address
                'Debug.Print materialName & "���Ƽƶq��}:  " & materialQuantity.Address
            End If
            
        Next
        
        
        '�U�@�Ӥ���
        Set element = element.Offset(0, 1)
    Loop
    MsgBox ("����")
End Sub

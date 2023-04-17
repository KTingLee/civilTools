Attribute VB_Name = "calMaterialsQuantity"
'Option Explicit

'�̷Ӥ���W�٨��o�Ӥ��󤧫��w���Ƽƶq
Function getMaterialsQuantity(elementName As Variant, materialName As Variant, sourceSheetName As Variant) As Variant
    '���Ƹ�ƨӷ��u�@��
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets(sourceSheetName)
    
    '�T�{���Ƽƶq�s�����
    Dim quantityCell As Range
    Dim quantityColWord As String: quantityColWord = "�p�p"
    Set quantityCell = sourceSheet.Cells.Find( _
        What:=quantityColWord, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    '�������ҨϥΪ�����
    Dim elementCell As Range
    Set elementCell = sourceSheet.Cells.Find( _
        What:=elementName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )

    Dim materialsRange As Range
    Set materialsRange = sourceSheet.Range(elementCell, elementCell.End(xlDown))

    '�j�M�Ӥ���O�_�����w����
    Dim material As Range
    Set material = materialsRange.Find( _
        What:=materialName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    If material Is Nothing Then
        Set getMaterialsQuantity = material
    Else
        Set getMaterialsQuantity = sourceSheet.Cells(material.Row, quantityCell.Column)
    End If
End Function

'���o�Ҧ����ƲM��
Function getMaterialsList() As Range
    Dim indexWord As String: indexWord = "�u�{����"
    Set indexCell = findCellByValue(indexWord).Offset(1, 0)
    Set getMaterialsList = Range(indexCell, indexCell.End(xlDown))
End Function

Sub test()
    '�@��u�{�ƶq�έp��A������k���}�l���O����W��
    Dim indexWord As String: indexWord = "���"
    Set indexCell = findCellByValue(indexWord)
    
    Dim materialsList As Range
    Set materialsList = getMaterialsList
    
    '�v����p��U�ӧ��ƪ��ƶq
    Dim element As Range
    Set element = Cells(indexCell.Row + 1, indexCell.Column + 1)  'indexCell.row+1 �O���F�O�I�A�קK while �L�k���X���x�s��
    Do While Not element.MergeCells
        
        elementName = element.value
        Set elementQuantity = element.Offset(1, 0)
        
        For Each material In materialsList
            materialName = material.value
            Set materialQuantity = getMaterialsQuantity(elementName, materialName, "��5_����ƶq�p���")
            
            'Debug.Print "����ƶq��}" & elementQuantity.Address
            
            If materialQuantity Is Nothing Then
                'Debug.Print elementName & "�S������:  " & materialName
            Else
                Cells(material.Row, element.Column).Formula = "=" & elementQuantity.Address & "*" & materialQuantity.Worksheet.Name & "!" & materialQuantity.Address
                'Debug.Print materialName & "���Ƽƶq��}:  " & materialQuantity.Address
            End If
            
        Next
        
        
        '�U�@�Ӥ���
        Set element = element.Offset(0, 1)
    Loop  'Loop ����
    'If rng.MergeCells
    
    
    
End Sub

Sub test2()
'Set res = findCellByValue("�A�ؼҪ�", "��5_����ƶq�p���")
'Set res = findCellByValue("���")
'Debug.Print res.value

'Set res = getMaterialsList()
'For Each material In res
'    Debug.Print material.value
'Next

Set res = getMaterialsQuantity("�a�u���פg��(h=2.0m)", "�A�ؼҪ�", "��5_����ƶq�p���")
Debug.Print res.value

End Sub

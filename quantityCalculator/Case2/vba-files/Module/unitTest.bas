Attribute VB_Name = "unitTest"
Option Explicit

'����specialCells�\��
Sub testSpecialCell()
    Cells(1, 1).SpecialCells(xlCellTypeVisible).Select
    Cells(1, 1).SpecialCells(xlCellTypeConstants).Select
End Sub

Sub testCreateSheet()
    Dim workElementsSheet As Worksheet
    Set workElementsSheet = activateAndSelectSheet("��3_����ƶq�p���")

    '�إߧ��Ʋξ��
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "���ƾ�z"
    Set targetSheet = createSheet(targetSheetName)
    workElementsSheet.Activate
End Sub

'�L�o�������x�s��
Sub testFilterColor()
    Dim res As Range
    Set res = getWorkList(ActiveSheet.Range("$A$1:$A$87"))

    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "���ƾ�z"
    Set targetSheet = createSheet(targetSheetName)
    Set res = copyRangeToSheet(res, targetSheet, 5)
End Sub

'���լO�_���V�X�i�ܤp�p���
Sub testGetMaterialQuantityByElementInWork()
    Dim testRange As Range
    Set testRange = getMaterialQuantityByElementInWork(1, "123", "456", "��3_����ƶq�p���")
    testRange.Select
End Sub

'���թw�q�W�٦p��ϥ�
Sub testRangeName()
    Dim testRange As Range
    Set testRange = Range("elementsMaterialSheetParam")
    Debug.Print testRange.Value
End Sub

'���կ�_���Q���o����p�p
Sub testGetElementQuantity()
    Dim workRange As Range
    Set workRange = Range(Cells(7, 1), Cells(17, 6))
    
    Dim materialName As String: materialName = "���x"
    Dim elementName As String: elementName = "���x"
    Dim materialQuantityCell As Range
    Set materialQuantityCell = getMaterialQuantityCell(materialName, elementName, workRange)
    materialQuantityCell.Select
End Sub

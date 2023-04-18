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

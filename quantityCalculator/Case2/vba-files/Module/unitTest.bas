Attribute VB_Name = "unitTest"
Option Explicit

'測試specialCells功能
Sub testSpecialCell()
    Cells(1, 1).SpecialCells(xlCellTypeVisible).Select
    Cells(1, 1).SpecialCells(xlCellTypeConstants).Select
End Sub

Sub testCreateSheet()
    Dim workElementsSheet As Worksheet
    Set workElementsSheet = activateAndSelectSheet("表3_元件數量計算表")

    '建立材料統整表
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "材料整理"
    Set targetSheet = createSheet(targetSheetName)
    workElementsSheet.Activate
End Sub

'過濾有底色儲存格
Sub testFilterColor()
    Dim res As Range
    Set res = getWorkList(ActiveSheet.Range("$A$1:$A$87"))

    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "材料整理"
    Set targetSheet = createSheet(targetSheetName)
    Set res = copyRangeToSheet(res, targetSheet, 5)
End Sub

'測試是否能橫向擴展至小計欄位
Sub testGetMaterialQuantityByElementInWork()
    Dim testRange As Range
    Set testRange = getMaterialQuantityByElementInWork(1, "123", "456", "表3_元件數量計算表")
    testRange.Select
End Sub

'測試定義名稱如何使用
Sub testRangeName()
    Dim testRange As Range
    Set testRange = Range("elementsMaterialSheetParam")
    Debug.Print testRange.Value
End Sub

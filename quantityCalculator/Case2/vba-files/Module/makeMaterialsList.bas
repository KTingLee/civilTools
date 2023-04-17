Attribute VB_Name = "makeMaterialsList"
'Option Explicit

'列出各工程使用的元件及材料 (TODO: 之後應該拆)
Sub makeMaterialsListByWork()
    '建立材料統整表
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "材料整理"
    If Not isSheetExist(targetSheetName) Then
        Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        targetSheet.Name = targetSheetName
    Else
        Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    End If
    
    '找出"項目"cell
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("項目")

    workNo = 1
    isLastWork = False
    '反覆處理各個工程
    Do Until isLastWork
        Dim searchRange As Range
        Set searchRange = getWorkRange(workNo)
        If searchRange Is Nothing Then
            Exit Sub
        End If
        
        '確認選取的材料範圍是在項目column位置
        Dim rangeDiff As Integer: rangeDiff = itemTitleCell.Column - searchRange.Column
        
        '過濾材料
        Dim noBackgroundColorRange As Range
        Dim notEmptyRange As Range
        Set noBackgroundColorRange = filterCells(searchRange.Offset(0, rangeDiff), xlFilterNoFill)
        Set notEmptyRange = filterCells(searchRange.Offset(0, rangeDiff), xlFilterValues)
        
        '利用交集，取得所有材料
        Dim materialsRange As Range
        Set materialsRange = Intersect(noBackgroundColorRange, notEmptyRange)
        
        '將材料複製到材料表，並移除重複的材料
        If Not materialsRange Is Nothing Then
            Dim copiedMaterialsRange As Range
            Set copiedMaterialsRange = copyRangeToSheet(materialsRange, targetSheet, 1)
            removeDuplicatesInColumn (copiedMaterialsRange)
        End If
    
    Loop
    
    
End Sub

'取得指定工程的元件範圍
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

'基本過濾資料
'XlAutoFilterOperator: xlFilterNoFill(過濾白底)、xlFilterValues(過濾非空白資料)
Function filterCells(searchRange As Range, Optional filterOperator As XlAutoFilterOperator = xlFilterNoFill) As Range
    searchRange.AutoFilter Field:=1, Criteria1:="<>", Operator:=filterOperator
    Set filterCells = ActiveSheet.AutoFilter.Range.Offset(0, 0).SpecialCells(xlCellTypeVisible)  '利用SpecialCells將過濾後的資料轉成range
    ActiveSheet.AutoFilterMode = False  '關閉過濾(excel中的過濾圖案會消失)
End Function

'將range複製到指定工作表的特定位置，並回傳目標工作表中取得複製的範圍
Function copyRangeToSheet(sourceRange As Range, targetSheet As Worksheet, targetColumn As Integer) As Range
    sourceRange.Copy Destination:=targetSheet.Cells(1, targetColumn)
    Dim copiedRange As Range
    Set copiedRange = Intersect(targetSheet.UsedRange, targetSheet.Columns(targetColumn))
    Set copyRangeToSheet = copiedRange
End Function

'移除指定範圍的重複資料(以column為一整組)
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

    '將重複的材料去除
    If Not materialsRange Is Nothing Then
        '先備份至別處
        Dim tempSheet As Worksheet
        Set tempSheet = ThisWorkbook.Sheets.Add
        materialsRange.Copy Destination:=tempSheet.Range("A1")
        
        '移除重複
        Dim copiedMaterialsRange As Variant
        Set copiedMaterialsRange = tempSheet.Range("A1").SpecialCells(xlCellTypeConstants)
        copiedMaterialsRange.RemoveDuplicates Columns:=1, Header:=xlNo
    End If
    
End Sub

Sub makeMaterialsListByWorkUnitTest2()
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "材料整理"
    If Not isSheetExist(targetSheetName) Then
        Set targetSheet = ThisWorkbook.Sheets.Add
        targetSheet.Name = targetSheetName
    Else
        Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    End If
End Sub

Sub teettt()
    Dim workElementsSheet As Worksheet
    Set workElementsSheet = activateAndSelectSheet("表3_元件數量計算表")
    
    '建立材料統整表
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "材料整理"
    If Not isSheetExist(targetSheetName) Then
        Set targetSheet = ThisWorkbook.Sheets.Add
        targetSheet.Name = targetSheetName
        workElementsSheet.Activate
    Else
        Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    End If
    
    '找出"項目"cell
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("項目")

    Dim searchRange As Range
    Set searchRange = getWorkRange(1)
    If searchRange Is Nothing Then
        Exit Sub
    End If
    
    '確認選取的材料範圍是在項目column位置
    Dim rangeDiff As Integer: rangeDiff = itemTitleCell.Column - searchRange.Column
    
    Dim copiedMaterialsRange As Range
    Set copiedMaterialsRange = copyRangeToSheet(searchRange.Offset(0, rangeDiff), targetSheet, 1)
    targetSheet.Activate
    
    '過濾材料
    Dim noBackgroundColorRange As Range
    Dim notEmptyRange As Range
    Set noBackgroundColorRange = filterCells(copiedMaterialsRange, xlFilterNoFill)
    Set copiedMaterialsRange2 = copyRangeToSheet(noBackgroundColorRange, targetSheet, 3)
    Set notEmptyRange = filterCells(copiedMaterialsRange, xlFilterValues)
    Set copiedMaterialsRange3 = copyRangeToSheet(notEmptyRange, targetSheet, 5)
    
    '利用交集，取得所有材料
    Dim materialsRange As Range
    Set materialsRange = Intersect(noBackgroundColorRange, notEmptyRange)
    Set copiedMaterialsRange = copyRangeToSheet(materialsRange, targetSheet, 7)
    
    '將材料複製到材料表，並移除重複的材料
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
    Set workElementsSheet = activateAndSelectSheet("表3_元件數量計算表")
    
    '建立材料統整表
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "材料整理"
    If Not isSheetExist(targetSheetName) Then
        Set targetSheet = ThisWorkbook.Sheets.Add
        targetSheet.Name = targetSheetName
        workElementsSheet.Activate
    Else
        Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    End If
    
    '找出"項目"cell
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("項目")

    Dim searchRange As Range
    Set searchRange = getWorkRange(1)
    If searchRange Is Nothing Then
        Exit Sub
    End If
    
    '確認選取的材料範圍是在項目column位置
    Dim rangeDiff As Integer: rangeDiff = itemTitleCell.Column - searchRange.Column

    searchRange.Offset(0, rangeDiff).Copy Destination:=targetSheet.Cells(1, 2)
    targetSheet.Activate
    Dim copiedRange As Range
    Set copiedRange = Intersect(targetSheet.UsedRange, targetSheet.Columns(2))
    copiedRange.Select
    Debug.Print targetSheet.UsedRange
    Debug.Print targetSheet.Columns(2)
    
    
End Sub

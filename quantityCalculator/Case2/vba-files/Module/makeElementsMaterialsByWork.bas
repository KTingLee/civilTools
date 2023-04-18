Attribute VB_Name = "makeElementsMaterialsByWork"
'Option Explicit

'列出各工程使用的元件及材料 (TODO: 之後應該拆)
Sub makeElementsMaterialsByWork()
    Dim workElementSheetParamCell As Range
    Set workElementSheetParamCell = getParamCell("elementsMaterialSheetParam", "請提供元件數量表名稱")
    
    If Not isSheetExist(workElementSheetParamCell.Value) Then
        MsgBox ("找不到 " & workElementSheetParamCell.Value)
        Exit Sub
    End If

    Dim workElementsSheet As Worksheet
    Set workElementsSheet = activateAndSelectSheet(workElementSheetParamRange.Value)  '參數儲存格

    '建立材料統整表
    Dim targetSheet As Worksheet
    Dim targetSheetName As String: targetSheetName = "材料整理"
    Set targetSheet = createSheet(targetSheetName)
    workElementsSheet.Activate
    
    '找出"項目"cell
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("項目")
    If itemTitleCell Is Nothing Then
        MsgBox ("缺少'項目'儲存格，無法定位")
        Exit Sub
    End If

    Dim workNo As Integer: workNo = 1
    Dim isLastWork As Boolean: isLastWork = False
    '反覆處理各個工程
    Do Until isLastWork
        Dim workRange As Range
        Set workRange = getWorkRange(workNo, itemTitleCell)
        If workRange Is Nothing Then
            isLastWork = True
            MsgBox ("無工程" & num2Tc(workNo))
            Exit Sub
        End If
        
        '單類工程資料佔4行，第一行是整個元件含重複材料，第二行是不重複材料，第三行元件
        Dim baseColumn As Integer: baseColumn = (workNo - 1) * 4 + 1
        
        Dim copiedWorkRange As Range
        Set copiedWorkRange = copyRangeToSheet(workRange, targetSheet, baseColumn)
        targetSheet.Activate  '切到材料整理表，才能做過濾
        
        '取得單類工程所有材料並去除重複
        Dim materialsList As Range
        Set materialsList = getMaterialsList(copiedWorkRange)
        Set materialsList = copyRangeToSheet(materialsList, targetSheet, baseColumn + 1)  '保留總材料(含重複材料)，並將去重複的結果放在旁邊
        removeDuplicatesInColumn materialsList
        
        '取得單類工程所有元件
        Dim elementsList As Range
        Set elementsList = getWorkList(copiedWorkRange)
        Set elementsList = copyRangeToSheet(elementsList, targetSheet, baseColumn + 2)
        removeDuplicatesInColumn elementsList
        
        workNo = workNo + 1
        
        '切回原本的工作表
        workElementsSheet.Activate
    Loop
End Sub

'取得指定工程的元件範圍
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


'基本過濾資料
'XlAutoFilterOperator: xlFilterNoFill(過濾白底)、xlFilterValues(過濾非空白資料)、xlFilterCellColor(過濾顏色)
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

'取得單類工程使用元件之總材料(傳入單類工程總範圍)
Function getMaterialsList(workElementMaterialsRange As Range) As Range
    '過濾材料
    Dim noBackgroundColorRange As Range
    Dim notEmptyRange As Range
    Set noBackgroundColorRange = filterCells(workElementMaterialsRange, xlFilterNoFill)
    Set notEmptyRange = filterCells(workElementMaterialsRange, xlFilterValues)
    
    '利用交集，取得所有材料
    Dim materialsRange As Range
    Set materialsRange = Intersect(noBackgroundColorRange, notEmptyRange)
    'removeDuplicatesInColumn materialsRange
    Set getMaterialsList = materialsRange
End Function

'取得單類工程的所有元件
Function getWorkList(workElementMaterialsRange As Range) As Range
    '元件儲存格有底色，以下先過濾出無底色，並與原始範圍作差集
    Dim noBackgroundColorRange As Range
    Set noBackgroundColorRange = filterCells(workElementMaterialsRange, xlFilterNoFill)
    Set getWorkList = Difference(workElementMaterialsRange, noBackgroundColorRange)
End Function





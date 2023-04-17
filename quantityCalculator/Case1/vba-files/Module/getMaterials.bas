Attribute VB_Name = "getMaterials"
'Option Explicit

'取得各元件的材料
Sub getMaterialsFromElements()
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets("表4_工程數量統計表1")
    
    'todo: 後續開一張settings表格
    sourceSheetName = "表5_元件數量計算表"
    Worksheets(sourceSheetName).Activate
    
    '找出目前資料的最末 row
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '找出"項目"所在的 row, column
    Dim searchValue As String
    Dim foundCell As Range
    searchValue = "項目"
    Set foundCell = Cells.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole _
        , SearchOrder:=xlByRows, SearchDirection:=xlNext)
    
    '選取材料並過濾掉空白(階段一，過濾白底資料)
    Range(foundCell, Cells(lastRow, foundCell.Column)).AutoFilter Field:=1, Criteria1:="<>", Operator:=xlFilterNoFill  '過濾出白底的資料
    Set noBackgroundColorRange = ActiveSheet.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)  '利用SpecialCells將過濾後的資料轉成range

    '過濾非空白資料(階段二，過濾非空白資料)
    Range(foundCell, Cells(lastRow, foundCell.Column)).AutoFilter Field:=1, Criteria1:="<>", Operator:=xlFilterValues  '過濾非空白資料
    Set notEmptyRange = ActiveSheet.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)
    
    ActiveSheet.AutoFilterMode = False  '關閉過濾(excel中的過濾圖案會消失)
    
    '利用交集，取得所有材料
    Dim materialsRange As Range
    Set materialsRange = Intersect(noBackgroundColorRange, notEmptyRange)
    
    '將重複的材料去除
    If Not materialsRange Is Nothing Then
        '先備份至別處
        Set tempSheet = ThisWorkbook.Sheets.Add
        materialsRange.Copy Destination:=tempSheet.Range("A1")
        '移除重複
        Set copiedMaterialsRange = tempSheet.Range("A1").SpecialCells(xlCellTypeConstants)
        copiedMaterialsRange.RemoveDuplicates Columns:=1, Header:=xlNo
        '將結果複製到目標工作表
        copiedMaterialsRange.Copy
        targetSheet.Range("B6").PasteSpecial Paste:=xlPasteValues
        
        Application.DisplayAlerts = False ' 禁止顯示刪除警告
        tempSheet.Delete
        Application.DisplayAlerts = True ' 啟用顯示警告
    End If
    
    

End Sub


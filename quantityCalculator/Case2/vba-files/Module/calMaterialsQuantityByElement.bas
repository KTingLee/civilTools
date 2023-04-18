Attribute VB_Name = "calMaterialsQuantityByElement"
Option Explicit

'取得該工程所有材料清單
Function getMaterialsInCurrentWork() As Range
    Dim indexCell As Range
    Set indexCell = findCellByValue("工程項目").Offset(1, 0)  '材料錨定點
    Set getMaterialsInCurrentWork = Range(indexCell, indexCell.End(xlDown))
End Function

'取得單類工程的元件及材料總範圍(含單位、小計數量)
Function getElementsAndMaterialsRangeByWork(workNo As Integer, workElementsSheetName As String) As Variant
    '紀錄原始工作表，最後要切回來
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet

    '切到元件數量表
    Dim workElementsSheet As Worksheet
    Set workElementsSheet = activateAndSelectSheet(workElementsSheetName)
    
    Dim itemTitleCell As Range
    Set itemTitleCell = findCellByValue("項目")  '元件表中，材料錨定點
    
    Dim quantityCell As Range
    Set quantityCell = findCellByValue("單位")  '元件表中，元件、材料資訊錨定點
    
    '取得該工程的所有元件與材料
    Dim workRange As Range
    Set workRange = getWorkRange(workNo, itemTitleCell)
    If workRange Is Nothing Then
        MsgBox ("請確認工程編號是否存在")  'TODO: 暫時做一個錯誤處理，之後有心力在來看怎麼改比較好
        Set getElementsAndMaterialsRangeByWork = workRange
        Exit Function
    End If
    
    Dim rangeDiff As Integer: rangeDiff = quantityCell.Column - workRange.Column
    Set workRange = workRange.Resize(workRange.Rows.Count, workRange.Columns.Count + rangeDiff)  '橫向擴展至單位範圍
    
    Set getElementsAndMaterialsRangeByWork = workRange
    currentSheet.Activate
End Function

'回傳元件或材料的單位
Function getObjectUnitCell(objectName As String, workRange As Range) As Range
    '找出元件或材料的儲存格(分兩種座標，一個是在workRange，另一個是在原始sheet中的座標)
    Dim objectCell As Range
    Set objectCell = findCellByValueInRange(objectName, workRange)
    
    If objectCell Is Nothing Then
        Set getObjectUnitCell = objectCell  'TODO: 暫時做一個錯誤處理，之後有心力在來看怎麼改比較好
        Exit Function
    End If
    
    '透過元件、材料本身的工作表去找單位欄位，因為workRange不含單位標頭列
    Dim indexCell As Range
    Set indexCell = findCellByValue("單位", objectCell.Worksheet.Name)

    '直接從元件、材料本身的工作表(元件表)找單位
    Set getObjectUnitCell = objectCell.Worksheet.Cells(objectCell.Row, indexCell.Column)
End Function

'回傳元件某個材料數量
Function getMaterialQuantityCell(materialName As String, elementName As String, workRange As Range) As Range
    Dim elementCell As Range
    Set elementCell = findCellByValueInRange(elementName, workRange)
    
    '透過元件、材料本身的工作表去找單位欄位，因為workRange不含單位標頭列
    Dim indexCell As Range
    Set indexCell = findCellByValue("小計", elementCell.Worksheet.Name)
    
    '依照元件，重新調整workRange範圍，避免找到前一個元件的材料
    Dim startCell As Range
    Dim endCell As Range
    Set startCell = workRange.Worksheet.Cells(elementCell.Row, elementCell.Column)  '這是從元件工作表座標找
    'Set endCell = workRange.Cells(workRange.Rows.Count, workRange.Columns.Count)  '這是從workRange座標找 -> (這樣可能包含該元件沒有的材料)
    
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

'工程數量統計表: 建立元件與材料的關係式(算數量)
Sub calMaterialsQuantityByElement()
    Dim workElementSheetParamCell As Range
    Set workElementSheetParamCell = getParamCell("elementsMaterialSheetParam", "請提供元件數量表名稱")
    
    Dim elementsQuantitySheetParamCell As Range
    Set elementsQuantitySheetParamCell = getParamCell("materialsQuantitySheetParam", "請設定目標工程數量統計表")
    
    Dim workNoParamCell As Range
    Set workNoParamCell = getParamCell("workNoParam", "請設定目標工程編號")
    
    
    If Not isSheetExist(workElementSheetParamCell.Value) Or Not isSheetExist(elementsQuantitySheetParamCell.Value) Or IsEmpty(workNoParamCell) Then
        MsgBox ("請確認元件表或工程數量統計表、工程編號是否正確")
        Exit Sub
    End If

    Dim elementsQuantitySheet As Worksheet
    Set elementsQuantitySheet = activateAndSelectSheet(elementsQuantitySheetParamCell.Value)  '切到目標工程數量統計表

    Dim indexCell As Range
    Set indexCell = findCellByValue("單位", elementsQuantitySheet.Name)  '元件錨定點，單位欄位右側開始都是元件名稱
    
    Dim materialsList As Range
    Set materialsList = getMaterialsInCurrentWork
    materialsList.Select
    
    Dim workRange As Range
    Set workRange = getElementsAndMaterialsRangeByWork(workNoParamCell.Value, workElementSheetParamCell.Value)  '傳入工程編號、元件表，以選取工程元件材料範圍
    If workRange Is Nothing Then
        Exit Sub  'TODO: 暫時做一個錯誤處理，之後有心力在來看怎麼改比較好
    End If
    
    '處理材料單位
    Dim material As Range
    Dim materialName As String
    Dim unitCell As Range
    For Each material In materialsList
        materialName = material.Value
        Set unitCell = getObjectUnitCell(materialName, workRange)
        
        If unitCell Is Nothing Then
            MsgBox ("在當前工程編號中似乎沒有使用: " & materialName)  'TODO: 暫時做一個錯誤處理，之後有心力在來看怎麼改比較好
            Exit Sub
        End If
        
        elementsQuantitySheet.Cells(material.Row, material.Column + 1).Formula = "=" & unitCell.Worksheet.Name & "!" & unitCell.Address
    Next
    
    '逐元件處理: 寫入元件單位、計算該元件各個材料的數量
    Dim element As Range
    Set element = Cells(indexCell.Row + 1, indexCell.Column + 1)  'indexCell.row+1 是為了保險，避免 while 無法找到合併儲存格
    
    Dim elementName As String
    Dim elementQuantityCell As Range
    Dim materialQuantityCell As Range
    Do While Not element.MergeCells  '總計欄位是合併儲存格，與元件儲存格格式不同
        '處理元件單位
        elementName = element.Value
        Set unitCell = getObjectUnitCell(elementName, workRange)
        
        Set elementQuantityCell = element.Offset(1, 0)
        elementQuantityCell.NumberFormatLocal = "0" & """" & unitCell.Value & """"  'TODO: 單位處理後面抽出來
        
        '開始計算元件的材料數量
        For Each material In materialsList
        
            materialName = material.Value
            Set materialQuantityCell = getMaterialQuantityCell(materialName, elementName, workRange)
            
            'Debug.Print "元件數量位址" & elementQuantity.Address
            
            If materialQuantityCell Is Nothing Then
                'Debug.Print elementName & "沒有材料:  " & materialName
            Else
                elementsQuantitySheet.Cells(material.Row, element.Column).Formula = "=" & elementQuantityCell.Address & "*" & materialQuantityCell.Worksheet.Name & "!" & materialQuantityCell.Address
                'Debug.Print materialName & "材料數量位址:  " & materialQuantity.Address
            End If
            
        Next
        
        
        '下一個元件
        Set element = element.Offset(0, 1)
    Loop
    MsgBox ("完成")
End Sub

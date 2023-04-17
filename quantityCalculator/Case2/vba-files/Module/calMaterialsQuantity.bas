Attribute VB_Name = "calMaterialsQuantity"
'Option Explicit

'依照元件名稱取得該元件之指定材料數量
Function getMaterialsQuantity(elementName As Variant, materialName As Variant, sourceSheetName As Variant) As Variant
    '材料資料來源工作表
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets(sourceSheetName)
    
    '確認材料數量存放的欄位
    Dim quantityCell As Range
    Dim quantityColWord As String: quantityColWord = "小計"
    Set quantityCell = sourceSheet.Cells.Find( _
        What:=quantityColWord, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    '選取元件所使用的材料
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

    '搜尋該元件是否有指定材料
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

'取得所有材料清單
Function getMaterialsList() As Range
    Dim indexWord As String: indexWord = "工程項目"
    Set indexCell = findCellByValue(indexWord).Offset(1, 0)
    Set getMaterialsList = Range(indexCell, indexCell.End(xlDown))
End Function

Sub test()
    '一般工程數量統計表，單位欄位右側開始都是元件名稱
    Dim indexWord As String: indexWord = "單位"
    Set indexCell = findCellByValue(indexWord)
    
    Dim materialsList As Range
    Set materialsList = getMaterialsList
    
    '逐元件計算各個材料的數量
    Dim element As Range
    Set element = Cells(indexCell.Row + 1, indexCell.Column + 1)  'indexCell.row+1 是為了保險，避免 while 無法找到合併儲存格
    Do While Not element.MergeCells
        
        elementName = element.value
        Set elementQuantity = element.Offset(1, 0)
        
        For Each material In materialsList
            materialName = material.value
            Set materialQuantity = getMaterialsQuantity(elementName, materialName, "表5_元件數量計算表")
            
            'Debug.Print "元件數量位址" & elementQuantity.Address
            
            If materialQuantity Is Nothing Then
                'Debug.Print elementName & "沒有材料:  " & materialName
            Else
                Cells(material.Row, element.Column).Formula = "=" & elementQuantity.Address & "*" & materialQuantity.Worksheet.Name & "!" & materialQuantity.Address
                'Debug.Print materialName & "材料數量位址:  " & materialQuantity.Address
            End If
            
        Next
        
        
        '下一個元件
        Set element = element.Offset(0, 1)
    Loop  'Loop 結尾
    'If rng.MergeCells
    
    
    
End Sub

Sub test2()
'Set res = findCellByValue("乙種模版", "表5_元件數量計算表")
'Set res = findCellByValue("單位")
'Debug.Print res.value

'Set res = getMaterialsList()
'For Each material In res
'    Debug.Print material.value
'Next

Set res = getMaterialsQuantity("懸臂式擋土牆(h=2.0m)", "乙種模版", "表5_元件數量計算表")
Debug.Print res.value

End Sub

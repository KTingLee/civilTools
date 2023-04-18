Attribute VB_Name = "commonTools"
Dim Json As Object

'eValue，不知道為何不能直接命名為 eval
Function eValue(x As String) As Double
  eValue = Evaluate(x)
End Function

'找出目前資料的最末 row 值
Function getLastRowNum() As Integer
    Dim lastRow As Long
    getLastRow = Cells(Rows.Count, 1).End(xlUp).Row
End Function

'找出目前資料的最末 row cell
Function getLastRow() As Range
    Set getLastRow = Cells(Rows.Count, 1).End(xlUp)
End Function

'自定義find(工作表全域搜尋)
Function findCellByValue(keyWord As Variant, Optional sheetName As Variant) As Range
    Dim searchSheet As Variant
    If IsMissing(sheetName) Then  '注意，isMissing 主要用於 Variant 參數
        Set searchSheet = ActiveSheet
    Else
        Set searchSheet = ThisWorkbook.Sheets(sheetName)
    End If
    
    Set findCellByValue = searchSheet.Cells.Find( _
        What:=keyWord, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
End Function

'自定義find(工作表依範圍搜尋)
Function findCellByValueInRange(keyWord As Variant, searchRange As Range) As Range
    Set findCellByValueInRange = searchRange.Find( _
        What:=keyWord, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
End Function


'檢查目標工作表是否存在
Function isSheetExist(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    isSheetExist = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            isSheetExist = True
            Exit Function
        End If
    Next
End Function

'建立工作表(並回傳該工作表)
Function createSheet(sheetName As String) As Worksheet
    If Not isSheetExist(sheetName) Then
        Set createSheet = ThisWorkbook.Sheets.Add
        createSheet.Name = sheetName
    Else
        Set createSheet = ThisWorkbook.Worksheets(sheetName)
    End If
End Function

'切換至指定工作表(並回傳該工作表)
Function activateAndSelectSheet(sheetName As String) As Worksheet
    Set activateAndSelectSheet = ThisWorkbook.Sheets(sheetName)
    activateAndSelectSheet.Activate
End Function

'數字轉中文(Traditional Chinese)
'NOTE: Remember import JSON.vba, then open the ref setting "Microsoft Scripting Runtime"
Function num2Tc(num As Integer) As String
    JsonConverter.JsonOptions.AllowUnquotedKeys = True
    Set chineseNumberJson = JsonConverter.ParseJson("{0: '零', 1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九'}")
    
    Set integerSplitObj = integerSplit(num)
    restNum = integerSplitObj("restNum")
    lastNum = integerSplitObj("lastNum")
    result = chineseNumberJson(CStr(lastNum))
    
    'num = 10~19，特別處理
    If restNum = 1 Then
        If lastNum = 0 Then
            result = "十"
        Else
            result = "十" & result
        End If
        num2Tc = result
        Exit Function
    End If

    'num > 19，以遞迴處理方式逐一轉換為中文
    Do While restNum >= 10
        Set integerSplitObj = integerSplit(restNum)
        restNum = integerSplitObj("restNum")
        lastNum = integerSplitObj("lastNum")
        result = chineseNumberJson(CStr(lastNum)) & result
    Loop
    If restNum <> 0 Then
        result = chineseNumberJson(CStr(restNum)) & result
    End If
    num2Tc = result
End Function

'將數字拆出個位數及剩餘數字(i.e. 將數字除以10取商與餘數)
Function integerSplit(ByVal num As Integer) As Object
    JsonConverter.JsonOptions.AllowUnquotedKeys = True
    restNum = num \ 10
    lastNum = num Mod 10
    
    result = "{" & _
        "restNum:" & restNum & "," & _
        "lastNum:" & lastNum & _
    "}"
    
    Set integerSplit = JsonConverter.ParseJson(result)
End Function

'ChatGPT提供的差集函式
Function Difference(rng1 As Range, rng2 As Range) As Range
    Dim cell As Range, checkCell As Range
    Dim result As Range
    Dim overlap As Boolean
    
    For Each cell In rng1
        overlap = False
        For Each checkCell In rng2
            If cell.Address = checkCell.Address Then
                overlap = True
                Exit For
            End If
        Next checkCell
        If Not overlap Then
            If result Is Nothing Then
                Set result = cell
            Else
                Set result = Application.Union(result, cell)
            End If
        End If
    Next cell
    
    Set Difference = result
End Function

'ChatGPT提供的差集函式2
Function Difference2(rng1 As Range, rng2 As Range) As Range
    Dim cell As Range, rngTemp As Range
    Set rngTemp = rng1.Cells(1, 1).EntireRow.Columns(1)  '感覺這個entireRow有點不妙
    For Each cell In rng1
        If Intersect(cell, rng2) Is Nothing Then
            Set rngTemp = Union(rngTemp, cell)
        End If
    Next cell
    Set Difference2 = Intersect(rngTemp, rng1)
End Function

'ChatGPT提供的參數檢查函數
Function getParamCell(ByVal cellName As String, ByVal errMsg As String) As Range
    Dim paramCell As Range
    Set paramCell = Range(cellName)
    If IsEmpty(paramCell) Then
        MsgBox (errMsg)
    End If
    Set getParamCell = paramCell
End Function

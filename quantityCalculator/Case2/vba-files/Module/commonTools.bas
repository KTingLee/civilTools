Attribute VB_Name = "commonTools"
Dim Json As Object

'eValueぃ笵ぃ钡㏑ eval
Function eValue(x As String) As Double
  eValue = Evaluate(x)
End Function

'тヘ玡戈程ソ row 
Function getLastRowNum() As Integer
    Dim lastRow As Long
    getLastRow = Cells(Rows.Count, 1).End(xlUp).Row
End Function

'тヘ玡戈程ソ row cell
Function getLastRow() As Range
    Set getLastRow = Cells(Rows.Count, 1).End(xlUp)
End Function

'﹚竡find
Function findCellByValue(keyWord As Variant, Optional sheetName As Variant) As Range
    Dim searchSheet As Variant
    If IsMissing(sheetName) Then  '猔種isMissing 璶ノ Variant 把计
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

'浪琩ヘ夹琌
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

'ち传﹚
Function activateAndSelectSheet(sheetName As String) As Worksheet
    Set activateAndSelectSheet = ThisWorkbook.Sheets(sheetName)
    activateAndSelectSheet.Activate
End Function

'计锣いゅ(Traditional Chinese)
'NOTE: Remember import JSON.vba, then open the ref setting "Microsoft Scripting Runtime"
Function num2Tc(num As Integer) As String
    JsonConverter.JsonOptions.AllowUnquotedKeys = True
    Set chineseNumberJson = JsonConverter.ParseJson("{0: '箂', 1: '', 2: '', 3: '', 4: '', 5: 'き', 6: 'せ', 7: '', 8: '', 9: ''}")
    
    Set integerSplitObj = integerSplit(num)
    restNum = integerSplitObj("restNum")
    lastNum = integerSplitObj("lastNum")
    result = chineseNumberJson(CStr(lastNum))
    
    'num = 10~19疭矪瞶
    If restNum = 1 Then
        If lastNum = 0 Then
            result = ""
        Else
            result = "" & result
        End If
        num2Tc = result
        Exit Function
    End If

    'num > 19患癹矪瞶よΑ硋锣传いゅ
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

'盢计╊计の逞緇计(i.e. 盢计埃10坝籔緇计)
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


Sub test33()
'Debug.Print num2Tc(21)
'Set res = integerSplit(21)
'Debug.Print res("restNum")
res = getLastRow()
Debug.Print res
End Sub

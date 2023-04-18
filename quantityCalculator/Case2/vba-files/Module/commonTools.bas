Attribute VB_Name = "commonTools"
Dim Json As Object

'eValue�A�����D���󤣯ઽ���R�W�� eval
Function eValue(x As String) As Double
  eValue = Evaluate(x)
End Function

'��X�ثe��ƪ��̥� row ��
Function getLastRowNum() As Integer
    Dim lastRow As Long
    getLastRow = Cells(Rows.Count, 1).End(xlUp).Row
End Function

'��X�ثe��ƪ��̥� row cell
Function getLastRow() As Range
    Set getLastRow = Cells(Rows.Count, 1).End(xlUp)
End Function

'�۩w�qfind(�u�@�����j�M)
Function findCellByValue(keyWord As Variant, Optional sheetName As Variant) As Range
    Dim searchSheet As Variant
    If IsMissing(sheetName) Then  '�`�N�AisMissing �D�n�Ω� Variant �Ѽ�
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

'�۩w�qfind(�u�@��̽d��j�M)
Function findCellByValueInRange(keyWord As Variant, searchRange As Range) As Range
    Set findCellByValueInRange = searchRange.Find( _
        What:=keyWord, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
End Function


'�ˬd�ؼФu�@��O�_�s�b
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

'�إߤu�@��(�æ^�ǸӤu�@��)
Function createSheet(sheetName As String) As Worksheet
    If Not isSheetExist(sheetName) Then
        Set createSheet = ThisWorkbook.Sheets.Add
        createSheet.Name = sheetName
    Else
        Set createSheet = ThisWorkbook.Worksheets(sheetName)
    End If
End Function

'�����ܫ��w�u�@��(�æ^�ǸӤu�@��)
Function activateAndSelectSheet(sheetName As String) As Worksheet
    Set activateAndSelectSheet = ThisWorkbook.Sheets(sheetName)
    activateAndSelectSheet.Activate
End Function

'�Ʀr�त��(Traditional Chinese)
'NOTE: Remember import JSON.vba, then open the ref setting "Microsoft Scripting Runtime"
Function num2Tc(num As Integer) As String
    JsonConverter.JsonOptions.AllowUnquotedKeys = True
    Set chineseNumberJson = JsonConverter.ParseJson("{0: '�s', 1: '�@', 2: '�G', 3: '�T', 4: '�|', 5: '��', 6: '��', 7: '�C', 8: '�K', 9: '�E'}")
    
    Set integerSplitObj = integerSplit(num)
    restNum = integerSplitObj("restNum")
    lastNum = integerSplitObj("lastNum")
    result = chineseNumberJson(CStr(lastNum))
    
    'num = 10~19�A�S�O�B�z
    If restNum = 1 Then
        If lastNum = 0 Then
            result = "�Q"
        Else
            result = "�Q" & result
        End If
        num2Tc = result
        Exit Function
    End If

    'num > 19�A�H���j�B�z�覡�v�@�ഫ������
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

'�N�Ʀr��X�Ӧ�ƤγѾl�Ʀr(i.e. �N�Ʀr���H10���ӻP�l��)
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

'ChatGPT���Ѫ��t���禡
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

'ChatGPT���Ѫ��t���禡2
Function Difference2(rng1 As Range, rng2 As Range) As Range
    Dim cell As Range, rngTemp As Range
    Set rngTemp = rng1.Cells(1, 1).EntireRow.Columns(1)  '�Pı�o��entireRow���I����
    For Each cell In rng1
        If Intersect(cell, rng2) Is Nothing Then
            Set rngTemp = Union(rngTemp, cell)
        End If
    Next cell
    Set Difference2 = Intersect(rngTemp, rng1)
End Function

'ChatGPT���Ѫ��Ѽ��ˬd���
Function getParamCell(ByVal cellName As String, ByVal errMsg As String) As Range
    Dim paramCell As Range
    Set paramCell = Range(cellName)
    If IsEmpty(paramCell) Then
        MsgBox (errMsg)
    End If
    Set getParamCell = paramCell
End Function

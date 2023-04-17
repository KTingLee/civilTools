Attribute VB_Name = "commonTools"
'eValue�A�����D���󤣯ઽ���R�W�� eval
Function eValue(x As String) As Double
  eValue = Evaluate(x)
End Function

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

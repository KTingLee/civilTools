Attribute VB_Name = "getMaterials"
'Option Explicit

'���o�U���󪺧���
Sub getMaterialsFromElements()
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets("��4_�u�{�ƶq�έp��1")
    
    'todo: ����}�@�isettings���
    sourceSheetName = "��5_����ƶq�p���"
    Worksheets(sourceSheetName).Activate
    
    '��X�ثe��ƪ��̥� row
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '��X"����"�Ҧb�� row, column
    Dim searchValue As String
    Dim foundCell As Range
    searchValue = "����"
    Set foundCell = Cells.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole _
        , SearchOrder:=xlByRows, SearchDirection:=xlNext)
    
    '������ƨùL�o���ť�(���q�@�A�L�o�թ����)
    Range(foundCell, Cells(lastRow, foundCell.Column)).AutoFilter Field:=1, Criteria1:="<>", Operator:=xlFilterNoFill  '�L�o�X�թ������
    Set noBackgroundColorRange = ActiveSheet.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)  '�Q��SpecialCells�N�L�o�᪺����নrange

    '�L�o�D�ťո��(���q�G�A�L�o�D�ťո��)
    Range(foundCell, Cells(lastRow, foundCell.Column)).AutoFilter Field:=1, Criteria1:="<>", Operator:=xlFilterValues  '�L�o�D�ťո��
    Set notEmptyRange = ActiveSheet.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)
    
    ActiveSheet.AutoFilterMode = False  '�����L�o(excel�����L�o�Ϯ׷|����)
    
    '�Q�Υ涰�A���o�Ҧ�����
    Dim materialsRange As Range
    Set materialsRange = Intersect(noBackgroundColorRange, notEmptyRange)
    
    '�N���ƪ����ƥh��
    If Not materialsRange Is Nothing Then
        '���ƥ��ܧO�B
        Set tempSheet = ThisWorkbook.Sheets.Add
        materialsRange.Copy Destination:=tempSheet.Range("A1")
        '��������
        Set copiedMaterialsRange = tempSheet.Range("A1").SpecialCells(xlCellTypeConstants)
        copiedMaterialsRange.RemoveDuplicates Columns:=1, Header:=xlNo
        '�N���G�ƻs��ؼФu�@��
        copiedMaterialsRange.Copy
        targetSheet.Range("B6").PasteSpecial Paste:=xlPasteValues
        
        Application.DisplayAlerts = False ' �T����ܧR��ĵ�i
        tempSheet.Delete
        Application.DisplayAlerts = True ' �ҥ����ĵ�i
    End If
    
    

End Sub


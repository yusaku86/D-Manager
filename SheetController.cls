VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sheetController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// �V�[�g��Z���𑀍삷��N���X
Option Explicit

'// �擪(�ŏI)�s���擾
Public Function getRow(rowType As String, standardColumn As Long, targetSheet As Worksheet) As Long

    With targetSheet
    
        If rowType = "start" Then
            getRow = .Cells(1, standardColumn).End(xlDown).Row
        ElseIf rowType = "last" Then
            getRow = .Cells(Rows.Count, standardColumn).End(xlUp).Row
        End If
        
    End With
    
End Function

'// �擪(�ŏI)����擾
Public Function getColumn(columnType As String, standardRow As Long, targetSheet As Worksheet) As Long
    
    With targetSheet
    
        If columnType = "start" Then
            getColumn = .Cells(standardRow, 1).End(xlToRight).Column
        ElseIf columnType = "last" Then
            getColumn = .Cells(standardRow, Columns.Count).End(xlToLeft).Column
        End If

    End With
    
End Function

'// �V�[�g�����݂��Ȃ���΍쐬(����΃N���A)
Public Sub createSheet(sheetName As String)
    
    If searchSheet(sheetName) = True Then
        Sheets(sheetName).Cells.Clear
    Else
        Dim newSheet As Worksheet
        
        Set newSheet = Sheets.Add(after:=Sheets(Sheets.Count))
        newSheet.Name = sheetName
        
        Set newSheet = Nothing
    End If
    
End Sub

'// �V�[�g�����邩�̔���
Private Function searchSheet(sheetName As String) As Boolean

    Dim sheet As Worksheet
    
    For Each sheet In Sheets
        If sheet.Name = sheetName Then
            searchSheet = True
            Exit Function
        End If
    Next
    
    searchSheet = False
    
End Function

'/**
 '* ������𐔒l�ɕϊ�
 '**/
 Public Sub toNumber(targetColumn As Long, startRow As Long, lastRow As Long)

    Columns(targetColumn).Insert xlToRight
    
    Dim columnIndexName As String: columnIndexName = Split(Columns(targetColumn + 1).Address, "$")(2)
    
    Cells(startRow, targetColumn).Formula = "=IF(" & columnIndexName & startRow & "="""","""",IFERROR(VALUE(" & columnIndexName & startRow & ")," & columnIndexName & startRow & "))"
    Cells(startRow, targetColumn).AutoFill Range(Cells(startRow, targetColumn), Cells(lastRow, targetColumn)), xlFillValues
    
    Columns(targetColumn).Copy
    Columns(targetColumn).PasteSpecial xlPasteValues
    Columns(targetColumn + 1).Delete xlToLeft
    
 End Sub
 
 '/**
 '* �f�[�^����בւ�
 '* @params keyColumn ���בւ��̃L�[
 '**/
 Public Sub sortValues(keyColumn As Long, targetRange)
 
    With ActiveSheet.Sort
        With .SortFields
            .Clear
            .Add Key:=Columns(keyColumn), SortOn:=xlSortOnValues, Order:=xlAscending
        End With
        
        .SetRange targetRange
        .Header = xlYes
        .Apply
    End With
        
 End Sub
 '/**
 '* ���t�̏�����a��ɕύX
 '* @params startRow �J�n�s
 '* @params lastRow �ŏI�s
 '* @params targetColumn �Ώۗ�
 '* @params dateType ���A���̂ǂ��܂ŕ\�����邩�w��(month,day�̂ǂ��炩)
 '**/
Public Sub changeDateFormat(startRow As Long, lastRow As Long, targetColumn As Long, dateType As String)

    Columns(targetColumn).Insert (xlToLeft)
    
    Dim columnIndexName As String: columnIndexName = Split(Columns(targetColumn + 1).Address, "$")(2)
    
    If dateType = "day" Then
        Cells(startRow, targetColumn).Formula = "=TEXT(DATE(LEFT(" & columnIndexName & startRow & ",4),MID(" & columnIndexName & startRow & " ,6,2),RIGHT(" & columnIndexName & startRow & ",2)),""ggge�Nmm��dd��"")"
    ElseIf dateType = "month" Then
        Cells(startRow, targetColumn).Formula = "=TEXT(DATE(LEFT(" & columnIndexName & startRow & ",4),MID(" & columnIndexName & startRow & " ,6,2),RIGHT(" & columnIndexName & startRow & ",2)),""ggge�Nmm��"")"
    End If
    
    Cells(startRow, targetColumn).AutoFill Range(Cells(startRow, targetColumn), Cells(lastRow, targetColumn)), xlFillValues
    
    Columns(targetColumn).Copy
    Columns(targetColumn).PasteSpecial xlPasteValues
    Columns(targetColumn + 1).Delete xlToLeft
    
End Sub
'/**
'* �w��̗�̒l�𔼊p�ɕϊ�
'* @params targetColumn �Ώۗ�
'**/
Public Sub convertIntoLower(startRow As Long, lastRow As Long, targetColumn As Long)

    Columns(targetColumn).Insert xlToRight
    
    Dim columnIndexName As String: columnIndexName = Split(Columns(targetColumn + 1).Address, "$")(2)
    
    Cells(startRow, targetColumn).Formula = "=SUBSTITUTE(ASC(" & columnIndexName & startRow & "),""����"",""��t"")"
    Cells(startRow, targetColumn).AutoFill Range(Cells(startRow, targetColumn), Cells(lastRow, targetColumn)), xlFillValues
    
    Columns(targetColumn).Copy
    Columns(targetColumn).PasteSpecial xlPasteValues
    Columns(targetColumn + 1).Delete xlToLeft

End Sub

'/**
' * �t�B���^�[����������A���o�������̂��폜����(�t�B���^�[������������ԂŎg�p����)
' * @params standardCell: ��ƂȂ�Z��
' * @params offsetRow: ������s��
'**/
Public Sub deleteAfterFilter(standardCell As Range)

    Dim targetRange As Range: Set targetRange = standardCell.CurrentRegion
    
    targetRange.Offset(1).Resize(targetRange.Rows.Count - 1).Delete (xlUp)
    standardCell.AutoFilter
    
    Set targetRange = Nothing
    
End Sub

'/**
 '* �t�B���^�[����������A���o�������̂�ʃV�[�g�ɓ\��t��
 '* @params standardCell: ��ƂȂ�Z��
 '* @params targetSheet: �\��t����̃V�[�g
 '**/
 Private Sub pasteAfterFilter(standardCell As Range, standardSheet As Worksheet, targetSheet As Worksheet)
  
    Dim targetRange As Range: Set targetRange = standardCell.CurrentRegion
    targetRange.Offset(1).Resize(targetRange.Rows.Count - 1).Copy
    
    targetSheet.Cells(Me.getRow("last", 1, targetSheet) + 1, 1).PasteSpecial
    
    Set targetRange = Nothing
    
 End Sub

'// �ԗ��̕��ނ��Ƃɕ�����T�u���[�`��
Public Sub divideTruck(trucktype() As Variant, tmpSheet As Worksheet, targetSheet As Worksheet)

    tmpSheet.Activate
    
    Dim startRow As Long: startRow = Me.getRow("start", 1, tmpSheet)
    Dim lastRow As Long: lastRow = Me.getRow("last", 1, tmpSheet)
    Dim lastColumn As Long: lastColumn = Me.getColumn("last", startRow, tmpSheet)
    Dim targetRange As Range: Set targetRange = Range(Cells(startRow, 1), Cells(lastRow, lastColumn))
    
    '// �w��̕��ނŃt�B���^�[��������(truckType�����̑��̏ꍇ�̓t�B���^�[�����Ȃ�)
    If trucktype(1) = "���̑�" Then
        targetRange.AutoFilter 1, "<>"
    Else
        filterByTruckTypes trucktype, targetSheet, targetRange
    End If
    
    Me.sortValues 1, targetRange                           '// �Ԏ햼�ŕ��ёւ�
    
    '// �t�B���^�[�������ĉ�������������Ȃ������甲����
    If Cells(Rows.Count, 1).End(xlUp).Row = startRow Then
        Cells(startRow, 1).AutoFilter
        Exit Sub
    End If
    
    '//�Ԏ�̌��o���Z���쐬
    setTruckTypeHeader trucktype(), targetSheet
        
    '// ���o��������targetSheet�ɓ\��t�� & tmpSheet����폜
    Dim standardCell As Range: Set standardCell = Cells(Me.getRow("start", 1, tmpSheet), 1)
    Call pasteAfterFilter(standardCell, tmpSheet, targetSheet)
    Call deleteAfterFilter(standardCell)
    
    '// �Ԏ킲�Ƃ�1�ԉ��̍s��2�d����
    Dim tmpLastRow As Long: tmpLastRow = Me.getRow("last", 1, targetSheet)
    
    With targetSheet
        .Activate
        .Range(.Cells(tmpLastRow, 1), .Cells(tmpLastRow, lastColumn)).Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    
    '// �ʂ��ԍ��t�^
    Call countTrucks(Cells(startRow, 2), targetSheet)
    
End Sub

'/**
 '* �Ԏ�ōi��ہA�����Ԏ����̃J�e�S���[�Ƃ���ꍇ(3t��4t�Ȃ�)
 '* �Ԏ킲�ƂɃt�B���^�[�������Ĉ��������������̂̍ŏI���1�E�̗�̒l�� targeted �ɂ��āA�Ō��targeted�̂��̂Ńt�B���^�[��������
 '*/
Private Sub filterByTruckTypes(truckTypes() As Variant, targetSheet As Worksheet, targetRange As Range)

    Dim i As Long
    Dim tmpRange As Range
    
    For i = 0 To UBound(truckTypes)
        '// �Ԏ킪�󔒂������玟�̃��[�v��
        If truckTypes(i) = "" Then
            GoTo Continue
        End If
        
        targetRange.AutoFilter 1, "*" & StrConv(truckTypes(i), vbNarrow) & "*", xlOr, "*" & StrConv(truckTypes(i), vbWide) & "*"
        
        '// �t�B���^�[�������ĉ�������������Ȃ���Ύ��̃��[�v��
        If Cells(Rows.Count, 1).End(xlUp).Row = targetRange(1, 1).Row Then
            targetRange.AutoFilter
            GoTo Continue
        End If
        
        Set tmpRange = targetRange.CurrentRegion
        tmpRange.Offset(1, tmpRange.Columns.Count - 1).Resize(, 1).Value = "targeted"
        targetRange.AutoFilter
Continue:
    Next
    
    Set tmpRange = Nothing
    
    targetRange.AutoFilter targetRange.Columns.Count, "targeted"

    
End Sub

'// �Ԏ�̐擪�̌��o���ݒ�
Private Sub setTruckTypeHeader(trucktype() As Variant, targetSheet As Worksheet)

    Dim truckTypeHeader As String: truckTypeHeader = createTruckTypeHeader(trucktype)
    
    With targetSheet.Cells(Me.getRow("last", 1, targetSheet) + 1, 1)
        .Value = truckTypeHeader
        .Interior.Color = RGB(192, 192, 192)
        .HorizontalAlignment = xlCenter
        .Font.Size = 12
        .Font.Bold = True
    End With

End Sub

'// �Ԏ�̌��o�����쐬
Private Function createTruckTypeHeader(trucktype() As Variant)

    Dim truckTypeHeader As String
    Dim i As Long
    
    For i = 0 To UBound(trucktype)
        If trucktype(i) = "" Then
            GoTo Continue
        End If
        
        If i = UBound(trucktype) Then
            truckTypeHeader = truckTypeHeader & trucktype(i)
        Else
            truckTypeHeader = truckTypeHeader & trucktype(i) & "�E"
        End If

Continue:
    Next
    
    createTruckTypeHeader = truckTypeHeader
    
End Function

'// �Ԏ킲�Ƃɒʂ��ԍ���t�^
Private Sub countTrucks(standardCell As Range, targetSheet As Worksheet)

    Dim startRow As Long: startRow = Me.getRow("last", 2, targetSheet) + 2
    Dim lastRow As Long: lastRow = Me.getRow("last", 1, targetSheet)
    
    Dim i As Long
    Dim counter As Long: counter = 1

    For i = startRow To lastRow
        targetSheet.Cells(i, 2).Value = counter
        counter = counter + 1
    Next
    
    '// �Ԏ�̑䐔�̍��v���w�b�_�[�̉��ɓ���(���̕\�Ŏg�p���邽��)
    targetSheet.Cells(startRow - 1, 2).Value = counter - 1
    
End Sub

'// �ԗ��ꗗ�̉��Ɋe�Ԏ킲�Ƃ̑䐔��\�������\���쐬
Public Sub createNumberOfTrucksChart(targetSheet As Worksheet, configSheet As Worksheet)

    targetSheet.Activate

    Cells(Me.getRow("last", 1, targetSheet) + 2, 1).Value = "�Ԏ킲�Ƃ̑䐔"

    Dim i As Long
    
    Dim truckTypes() As Variant
    Dim truckTypeHeader As String
    
    Dim targetRow As Long
    
    Dim numberOfTrucks As Long
    Dim totalAmount As Long
    
    For i = 2 To Me.getRow("last", 1, configSheet)
        truckTypes = splitToArray(Split(configSheet.Cells(i, 1), ","))
        truckTypeHeader = createTruckTypeHeader(truckTypes)                     '// �Ԏ햼
        
        If truckTypeHeader = "" Then
            GoTo Continue
        End If
        
        '// �Ώۂ̌��o���������ꂽ�Z���̍s�ԍ�(������Ȃ��ꍇ�̓G���[����������̂ł��̕����̓G���[����)
        On Error Resume Next
        targetRow = 0
        
        targetRow = WorksheetFunction.Match(truckTypeHeader, targetSheet.Columns(1), 0)
        
        On Error GoTo 0
        
        If targetRow = 0 Then
            GoTo Continue
        End If
        
        numberOfTrucks = Cells(targetRow, 2).Value                                         '// �Ԏ킲�Ƃ̍��v�䐔
        
        With Cells(Me.getRow("last", 1, targetSheet) + 1, 1)
            .Value = truckTypeHeader
            .Offset(, 1).Value = numberOfTrucks & "��"
        End With
        
        Cells(targetRow, 2).Value = ""
        
        totalAmount = totalAmount + numberOfTrucks

Continue:
    Next
        
    With Cells(Me.getRow("last", 1, targetSheet) + 1, 1)
        .Interior.Color = RGB(144, 238, 144)
        .Value = "���v"
        
        With .Offset(, 1)
            .Value = totalAmount & "��"
            .Interior.Color = RGB(144, 238, 144)
        End With
    End With
            
End Sub

'// split�֐��ŕ������l��z��ɕϊ�
Public Function splitToArray(splitedValue As Variant)


    Dim returnArray() As Variant
    ReDim returnArray(0)
    
    Dim i As Long
    
    For i = 0 To UBound(splitedValue)
        ReDim Preserve returnArray(UBound(returnArray) + 1)
        returnArray(UBound(returnArray)) = splitedValue(i)
    Next
    
    splitToArray = returnArray

End Function

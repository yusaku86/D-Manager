Attribute VB_Name = "createChart"
' D-Manager����o�͂���Excel�����₷�����H
Option Explicit

'// ���C�����[�`��
Public Sub main()
    
    If MsgBox("�\�����H���܂��B��낵���ł���?", vbYesNo, ThisWorkbook.Name) = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        
    Cells.UnMerge
    Cells.Borders.LineStyle = xlLineStyleNone
    
    Dim sc As New sheetController
    Dim startRow As Long: startRow = sc.getRow("start", 1, Sheets("�R�݉^��"))
    Dim lastRow As Long: lastRow = sc.getRow("last", 1, Sheets("�R�݉^��"))
    Dim lastColumn As Long: lastColumn = sc.getColumn("last", startRow, Sheets("�R�݉^��"))
    
    Cells(1, 5).Value = ""
    Cells(1, 17).Value = ""
    Cells(startRow, 5).Value = "���"
    Cells(startRow, 27).Value = "����ʍs������(�J�n)"
    Cells(startRow, 28).Value = "����ʍs������(�I��)"
    Cells(startRow, 29).Value = "�ʍs���؊���(�J�n)"
    Cells(startRow, 30).Value = "�ʍs���؊���(�I��)"
    Cells(startRow, 31).Value = "filtered"
    
    Columns("F:G").Delete xlToLeft
    Columns(6).Cut
    Columns(1).Insert xlToRight
    
    '// �Ԏ�Ɂu�e�X�g�v�Ƃ��������񂪓����Ă�����̂��폜
    Range(Cells(startRow, 1), Cells(lastRow, lastColumn)).AutoFilter 1, "*�e�X�g*", xlOr, "*ý�*"
    If Cells(Rows.Count, 1).End(xlUp).Row > startRow Then
        sc.deleteAfterFilter Cells(startRow, 1)
    End If
    
    '// �ő�ύڗʂ𐔒l�ɕϊ� & �ő�ύڗʂŏ����ɕ��בւ�
    sc.toNumber 14, startRow, lastRow
    Columns(14).NumberFormatLocal = "#,###"
    sc.sortValues 14, Range(Cells(startRow, 1), Cells(lastRow, lastColumn))
    
    '// ���t�𐼗��a��ɕϊ�
    sc.changeDateFormat startRow + 1, lastRow, 9, "month"
    Cells(startRow, 9).Value = "���N�x�o�^�N��"
    sc.changeDateFormat startRow + 1, lastRow, 10, "day"
    Cells(startRow, 10).Value = "�o�^�N����"
    
    '// �Ԏ�𔼊p�ɓ���
    sc.convertIntoLower startRow + 1, lastRow, 1
    Cells(startRow, 1).Value = "�Ԏ�"
    
    Columns(2).Insert xlToRight
    Cells(startRow, 2).Value = "�䐔"
    
    '// YCL�̕���YCL�̃V�[�g�Ɉړ� & �u�ۑ��v�{�^���ǉ�
    sc.createSheet "YCL"
    
    Sheets("YCL").Activate
    Dim bc As New buttonController
    bc.addButton Sheets("YCL"), Sheets("YCL").Range(Cells(1, 1), Cells(2, 1)), "�ۑ�", "openForm"
    Set bc = Nothing
    
    Sheets("�R�݉^��").Activate
    
    Range(Cells(startRow, 1), Cells(lastRow, lastColumn)).AutoFilter 3, "YCL"
    Cells(startRow, 1).CurrentRegion.Copy Sheets("YCL").Cells(startRow, 1)
    sc.deleteAfterFilter Cells(startRow, 1)
    
    '// �R�݉^����YCL�ȊO�̂��̂��폜
    Range(Cells(startRow, 1), Cells(lastRow, lastColumn)).AutoFilter 3, "<>�R�݉^��"
    sc.deleteAfterFilter Cells(startRow, 1)
    
    '// �R�݉^�����̃f�[�^���ꎞ�I�ɕۑ�����V�[�g�u�R�݉^��tmp�v���쐬 & �f�[�^���u�R�݉^��tmp�v�փR�s�[
    sc.createSheet "�R�݉^��tmp"
    pasteToTmpSheet Sheets("�R�݉^��"), Sheets("�R�݉^��tmp"), startRow, lastRow
    
    '// YCL���̃f�[�^���ꎞ�I�ɕۑ�����V�[�g�uYCLtmp�v���쐬 & �f�[�^���uYCLtmp�v�փR�s�[
    sc.createSheet "YCLtmp"
    Call pasteToTmpSheet(Sheets("YCL"), Sheets("YCLtmp"), sc.getRow("start", 1, Sheets("YCL")), sc.getRow("last", 1, Sheets("YCL")))
    
    '// �Ԏ킲�Ƃɕ��ނ����\���쐬
    Call classifyTruck(Sheets("�R�݉^��"), Sheets("�R�݉^��tmp"), Sheets("�ݒ�(�R�݉^��)"), sc)
    Call classifyTruck(Sheets("YCL"), Sheets("YCLtmp"), Sheets("�ݒ�(YCL)"), sc)
    
    '// �w�b�_�[�̌r���𑾎��ɕύX
    Call setHeaderLine(Sheets("�R�݉^��"), sc, xlMedium)
    Call setHeaderLine(Sheets("YCL"), sc, xlMedium)
    
    Sheets("�R�݉^��").Activate
    Cells(1, 1).Select
        
    Set sc = Nothing
    
    Application.DisplayAlerts = True
    
    MsgBox "�������������܂����B", Title:=ThisWorkbook.Name
    
End Sub

'// �ꎞ�I�ɍ쐬�����V�[�g�Ɍ��̃V�[�g�̃f�[�^���R�s�[(���̃V�[�g�̃w�b�_�[�͎c��)
Private Sub pasteToTmpSheet(targetSheet As Worksheet, tmpSheet As Worksheet, startRow As Long, lastRow As Long)

    With targetSheet
        .Activate
        .Cells.Copy Destination:=tmpSheet.Cells(1, 1)
        .Range(Cells(startRow + 1, 1), Cells(lastRow, lastRow)).Clear
    End With
        
End Sub

'/**
' *�Ԏ킲�Ƃɕ��ނ���T�u���[�`��
' *@params targetSheet �\��t����̃V�[�g
' *@params tmpSheet    �f�[�^���Ԏ킲�Ƃɕ����邽�߂Ɉꎞ�I�ɍ쐬����V�[�g
' *@params configSheet �Ԏ�̐ݒ蓙�������ꂽ�V�[�g
'**/
Private Sub classifyTruck(targetSheet As Worksheet, tmpSheet As Worksheet, configSheet As Worksheet, sc As sheetController)

    With configSheet
        .Activate
        
        Dim tmpCell As Range
        Dim truckTypes() As Variant
         
        Dim i As Long
        Dim splitedTmpcell As Variant
        
        For Each tmpCell In Range(Cells(2, 1), Cells(sc.getRow("last", 1, configSheet), 1))
            splitedTmpcell = Split(tmpCell.Value, ",")
            ReDim truckTypes(0)
                
            '// �ݒ�V�[�g�ɓ��͂��ꂽ�Ԏ���i�[�����z����쐬
            For i = 0 To UBound(splitedTmpcell)
                ReDim Preserve truckTypes(UBound(truckTypes) + 1)
                truckTypes(UBound(truckTypes)) = splitedTmpcell(i)
            Next
            
            sc.divideTruck truckTypes, tmpSheet, targetSheet
        Next
    End With
    
    tmpSheet.Delete
    
    '//�ԔԁE�Ԗ��̑䐔�̌�Ɉړ�
    With targetSheet
        .Activate
        .Range("H:I").Cut
        .Cells(3).Insert xlToRight
    End With
    
    '// �Ԏ킲�Ƃ̑䐔�̕\���ԗ��ꗗ�̉��ɍ쐬
    sc.createNumberOfTrucksChart targetSheet, configSheet
    
    '// �Ԏ킲�Ƃɕ��ނ���ۂɎg�p����filterd��폜
    Dim startRow As Long: startRow = sc.getRow("start", 1, targetSheet)
    targetSheet.Columns(sc.getColumn("last", startRow, targetSheet)).Delete
    
    '// �Ԍ��L��������ɏ����t�������ݒ�
    Call setFormatCondition(targetSheet)
    
    '// �w�b�_�[�̃t�H���g�T�C�Y�ύX�E�������낦�E�V�[�g�S�̂̃t�H���g�����C���I�ɕύX�E�E�B���h�E�g�̌Œ�
    With targetSheet
        With .Rows(startRow)
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        End With
    
        .Range("A:A,C:I,P:P,V:AC").HorizontalAlignment = xlCenter
        .Range(.Cells(sc.getRow("start", 2, targetSheet) + 1, 2), .Cells(sc.getRow("last", 2, targetSheet), 2)).HorizontalAlignment = xlRight
    
        With .Cells
            .Font.Name = "���C���I"
            .EntireColumn.AutoFit
        End With
        
        .Cells(4, 4).Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
    End With
    
End Sub

'/**
' * �Ԍ��L��������ɏ����t�������ݒ�
' * �����؂ꁨ��,10���ȓ������F,30���ȓ�����
' */
Private Sub setFormatCondition(targetSheet As Worksheet)

    With targetSheet.Range("P:P").FormatConditions
        .Delete
        
        '// �����؂�
        Dim fcRed As FormatCondition
        Set fcRed = .Add(Type:=xlExpression, Formula1:="=DATEVALUE(P1) <TODAY()")
        fcRed.Interior.Color = RGB(178, 34, 34)
        Set fcRed = Nothing
        
        '// 10�ȓ��ɓ���
        Dim fcYellow As FormatCondition
        Set fcYellow = .Add(Type:=xlExpression, Formula1:="=AND(DATEVALUE(P1) >=TODAY(), DATEVALUE(P1) <=TODAY() + 10)")
        fcYellow.Interior.Color = RGB(255, 255, 102)
        Set fcYellow = Nothing
        
        '// 30���ȓ��ɓ���
        Dim fcGreen As FormatCondition
        Set fcGreen = .Add(Type:=xlExpression, Formula1:="=AND(DATEVALUE(P1) > TODAY() + 10, DATEVALUE(P1) <=TODAY() + 30)")
        fcGreen.Interior.Color = RGB(50, 205, 50)
        Set fcGreen = Nothing
    End With
End Sub


'/**
 '* �w�b�_�[�̌r���𑾐��ɐݒ�
 '**/
Private Sub setHeaderLine(targetSheet As Worksheet, sc As sheetController, lineWeight As Long)

    Dim startRow As Long: startRow = sc.getRow("start", 1, targetSheet)

    With targetSheet
        With .Range(.Cells(startRow, 1), .Cells(startRow, sc.getColumn("last", startRow, targetSheet)))
            .Borders.Weight = lineWeight
        End With
    End With

End Sub

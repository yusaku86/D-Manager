Attribute VB_Name = "saveFile"
'// �쐬�����\�𖼑O��t���ĕۑ�
Option Explicit

'// ���C�����[�`��
Public Sub saveChart(targetYear As String, targetMonth As String)

    Dim myFso As New FileSystemObject
    Dim folderPath As String: folderPath = getPath(ActiveSheet.Name)
    
    '// �ۑ���p�X�m�F
    If folderPath = "" Then
        MsgBox "�t�@�C���̕ۑ��悪�ݒ肳��Ă��܂���B" & vbLf & "�ݒ�(" & ActiveSheet.Name & ")�̃V�[�g�́u�ۑ���ύX�v���ݒ肵�Ă��������B", vbQuestion, ThisWorkbook.Name
        GoTo Break
    ElseIf myFso.FolderExists(folderPath) = False Then
        MsgBox "�ۑ���Ƃ��Đݒ肳��Ă���t�H���_�����݂��܂���B" & vbLf & "�ۑ����ύX���Ă��������B", vbQuestion, ThisWorkbook.Name
        GoTo Break
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim fullPath As String: fullPath = folderPath & "\" & targetYear & "\" & ActiveSheet.Name & "�ԗ��ꗗ" & targetYear & targetMonth & ".xlsx"
    
    '//�ۑ���̃t�H���_�ɑΏ۔N�̃t�H���_��������΍쐬
    Call createFolderIfNotExist(folderPath & "\" & targetYear, myFso)
    
    '// ���Ƀt�@�C��������ꍇ�㏑�����邩�m�F
    If myFso.FileExists(fullPath) Then
        If (MsgBox("���̏ꏊ�Ɋ���" & vbLf & vbLf & fullPath & vbLf & vbLf & "�Ƃ����t�@�C�������݂��܂����㏑�����܂���?", vbYesNo, ThisWorkbook.Name)) = vbNo Then
            GoTo Break
        End If
    End If

    Set myFso = Nothing
    
    '// �{�^���폜
    Dim bc As New buttonController
    bc.deleteButtons ActiveSheet
    
    '// �\����̃t�@�C���Ƃ��ĕۑ�
    ActiveSheet.Copy
    ActiveWorkbook.SaveAs fullPath, xlOpenXMLWorkbook
    ActiveWorkbook.Close
    
    '// �{�^������
    bc.addButton ActiveSheet, Range(Cells(1, 1), Cells(2, 1)), "�ۑ�", "openForm"
    If ActiveSheet.Name = "�R�݉^��" Then
        bc.addButton ActiveSheet, Range(Cells(1, 3), Cells(2, 3)), "�\���H", "main"
    End If
    
    Set bc = Nothing
    
    MsgBox "�ۑ����������܂����B", Title:=ThisWorkbook.Name
    
Break:
    Set myFso = Nothing

End Sub
'// �ۑ���擾
Private Function getPath(company As String) As String

    getPath = Sheets("�ݒ�(" & company & ")").Cells(2, 2).Value
    
End Function

'// �t�H���_��������΍쐬
Private Sub createFolderIfNotExist(path As String, myFso As FileSystemObject)

    With myFso
        If .FolderExists(path) = False Then
            .CreateFolder path
        End If
    End With
    
End Sub

'// �\�̕ۑ���ݒ�
Public Sub setPath()

    Dim path As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "�ۑ���t�H���_�̐ݒ�"
        If .Show = True Then
            path = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ActiveSheet.Cells(2, 2).Value = path
        
End Sub



'// ���[�U�[�t�H�[���N��
Public Sub openForm()

    formPeriod.Show

End Sub

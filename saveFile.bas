Attribute VB_Name = "saveFile"
'// �쐬�����\�𖼑O��t���ĕۑ�
Option Explicit

'// �\�𖼑O��t���ĕۑ�(���C�����[�`��)
Public Sub saveChart(ByVal targetYear As String, ByVal targetMonth As String)

    Dim myFso As New FileSystemObject
    Dim folderPath As String: folderPath = getPath(ActiveSheet.Name)
    Dim companyName As String: companyName = ActiveSheet.Name
    
    '// �ۑ���p�X�m�F
    If folderPath = "" Then
        MsgBox "�t�@�C���̕ۑ��悪�ݒ肳��Ă��܂���B" & vbLf & "�ݒ�(" & companyName & ")�̃V�[�g�́u�ۑ���ύX�v���ݒ肵�Ă��������B", vbQuestion, ThisWorkbook.Name
        GoTo Break
    ElseIf myFso.FolderExists(folderPath) = False Then
        MsgBox "�ۑ���Ƃ��Đݒ肳��Ă���t�H���_�����݂��܂���B" & vbLf & "�ۑ����ύX���Ă��������B", vbQuestion, ThisWorkbook.Name
        GoTo Break
    End If
    
    Application.DisplayAlerts = False
    
    Dim fullPath As String: fullPath = folderPath & "\" & targetYear & "\" & companyName & "�ԗ��ꗗ" & targetYear & targetMonth & ".xlsx"
    
    '//�ۑ���̃t�H���_�ɑΏ۔N�̃t�H���_��������΍쐬
    Call createFolderIfNotExist(folderPath & "\" & targetYear, myFso)
    
    '// ���Ƀt�@�C��������ꍇ�㏑�����邩�m�F
    If myFso.FileExists(fullPath) Then
        If (MsgBox("���̏ꏊ�Ɋ���" & vbLf & vbLf & fullPath & vbLf & vbLf & "�Ƃ����t�@�C�������݂��܂����㏑�����܂���?", vbYesNo, ThisWorkbook.Name)) = vbNo Then
            GoTo Break
        End If
    End If

    Set myFso = Nothing
    
    '// �\����̃t�@�C���Ƃ��ĕۑ�
    ActiveSheet.Cells.Copy
    Dim addedFile As Workbook: Set addedFile = Workbooks.Add
    
    With addedFile
        .Sheets(1).Cells(1, 1).PasteSpecial
        .Sheets(1).Name = companyName
    End With
    
    ActiveWorkbook.SaveAs fullPath, xlOpenXMLWorkbook
    addedFile.Close
        
    Set addedFile = Nothing
    
    MsgBox "�ۑ����������܂����B", Title:=ThisWorkbook.Name
    
Break:
    Set myFso = Nothing

End Sub

'// �ۑ���擾
Private Function getPath(ByVal company As String) As String

    getPath = Sheets("�ݒ�(" & company & ")").Cells(2, 2).Value
    
End Function

'// �t�H���_��������΍쐬
Private Sub createFolderIfNotExist(ByVal path As String, ByVal myFso As FileSystemObject)

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
        .InitialFileName = "G:"
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

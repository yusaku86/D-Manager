Attribute VB_Name = "saveFile"
'// 作成した表を名前を付けて保存
Option Explicit

'// 表を名前を付けて保存(メインルーチン)
Public Sub saveChart(ByVal targetYear As String, ByVal targetMonth As String)

    Dim myFso As New FileSystemObject
    Dim folderPath As String: folderPath = getPath(ActiveSheet.Name)
    Dim companyName As String: companyName = ActiveSheet.Name
    
    '// 保存先パス確認
    If folderPath = "" Then
        MsgBox "ファイルの保存先が設定されていません。" & vbLf & "設定(" & companyName & ")のシートの「保存先変更」より設定してください。", vbQuestion, ThisWorkbook.Name
        GoTo Break
    ElseIf myFso.FolderExists(folderPath) = False Then
        MsgBox "保存先として設定されているフォルダが存在しません。" & vbLf & "保存先を変更してください。", vbQuestion, ThisWorkbook.Name
        GoTo Break
    End If
    
    Application.DisplayAlerts = False
    
    Dim fullPath As String: fullPath = folderPath & "\" & targetYear & "\" & companyName & "車両一覧" & targetYear & targetMonth & ".xlsx"
    
    '//保存先のフォルダに対象年のフォルダが無ければ作成
    Call createFolderIfNotExist(folderPath & "\" & targetYear, myFso)
    
    '// 既にファイルがある場合上書きするか確認
    If myFso.FileExists(fullPath) Then
        If (MsgBox("この場所に既に" & vbLf & vbLf & fullPath & vbLf & vbLf & "というファイルが存在しますが上書きしますか?", vbYesNo, ThisWorkbook.Name)) = vbNo Then
            GoTo Break
        End If
    End If

    Set myFso = Nothing
    
    '// 表を一つのファイルとして保存
    ActiveSheet.Cells.Copy
    Dim addedFile As Workbook: Set addedFile = Workbooks.Add
    
    With addedFile
        .Sheets(1).Cells(1, 1).PasteSpecial
        .Sheets(1).Name = companyName
    End With
    
    ActiveWorkbook.SaveAs fullPath, xlOpenXMLWorkbook
    addedFile.Close
        
    Set addedFile = Nothing
    
    MsgBox "保存が完了しました。", Title:=ThisWorkbook.Name
    
Break:
    Set myFso = Nothing

End Sub

'// 保存先取得
Private Function getPath(ByVal company As String) As String

    getPath = Sheets("設定(" & company & ")").Cells(2, 2).Value
    
End Function

'// フォルダが無ければ作成
Private Sub createFolderIfNotExist(ByVal path As String, ByVal myFso As FileSystemObject)

    With myFso
        If .FolderExists(path) = False Then
            .CreateFolder path
        End If
    End With
    
End Sub

'// 表の保存先設定
Public Sub setPath()

    Dim path As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "保存先フォルダの設定"
        .InitialFileName = "G:"
        If .Show = True Then
            path = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ActiveSheet.Cells(2, 2).Value = path
        
End Sub

'// ユーザーフォーム起動
Public Sub openForm()

    formPeriod.Show

End Sub

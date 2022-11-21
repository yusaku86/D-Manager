VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPeriod 
   Caption         =   "年月入力"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3570
   OleObjectBlob   =   "formPeriod.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// 表を保存する際のファイル名に使用する年月を指定
Option Explicit

'// 「実行」ボタンクリック時
Private Sub cmdEnter_Click()

    Call saveChart(cmbYear.Value, cmbMonth.Value)
    Unload Me
    
End Sub

'// フォーム起動時
Private Sub UserForm_Initialize()
    
    Dim i As Long

    '// コンボボックスの値設定(現在から5年まえまで選択肢として追加)
    With cmbYear
        For i = Year(Now) - 5 To Year(Now)
            .AddItem i & "年"
        Next
        .Value = Year(Now) & "年"
    End With
    
    With cmbMonth
        For i = 1 To 12
            .AddItem i & "月"
        Next
        .Value = Month(Now) & "月"
    End With
End Sub

'// 「閉じる」ボタンクリック時
Private Sub cmdClose_Click()

    Unload Me

End Sub

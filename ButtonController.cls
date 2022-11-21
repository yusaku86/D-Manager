VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "buttonController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// ボタン操作を制御するクラス
Option Explicit

'// シートのボタン削除
Public Sub deleteButtons(ByVal targetSheet As Worksheet)

    Dim btn As Object
    
    For Each btn In targetSheet.Buttons
        btn.Delete
    Next
    
End Sub

'// ボタン追加
Public Sub addButton(ByVal targetSheet As Worksheet, ByVal targetRange As Range, ByVal caption As String, ByVal program As String)

    With targetSheet.Buttons.Add(targetRange.Left + 30.5625, targetRange.Top + 3.75, targetRange.Width / 2, targetRange.Height / 3 * 2)
        .Characters.text = caption
        .Placement = xlFreeFloating
        .OnAction = program
    End With
    
End Sub

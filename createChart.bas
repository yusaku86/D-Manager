Attribute VB_Name = "createChart"
' D-Managerから出力したExcelを見やすく加工
Option Explicit

'// メインルーチン
Public Sub main()
    
    If MsgBox("表を加工します。よろしいですか?", vbYesNo, ThisWorkbook.Name) = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        
    Cells.UnMerge
    Cells.Borders.LineStyle = xlLineStyleNone
    
    Dim sc As New sheetController
    Dim startRow As Long: startRow = sc.getRow("start", 1, Sheets("山岸運送"))
    Dim lastRow As Long: lastRow = sc.getRow("last", 1, Sheets("山岸運送"))
    Dim lastColumn As Long: lastColumn = sc.getColumn("last", startRow, Sheets("山岸運送"))
    
    Cells(1, 5).Value = ""
    Cells(1, 17).Value = ""
    Cells(startRow, 5).Value = "状態"
    Cells(startRow, 27).Value = "特殊通行許可期限(開始)"
    Cells(startRow, 28).Value = "特殊通行許可期限(終了)"
    Cells(startRow, 29).Value = "通行許可証期限(開始)"
    Cells(startRow, 30).Value = "通行許可証期限(終了)"
    Cells(startRow, 31).Value = "filtered"
    
    Columns("F:G").Delete xlToLeft
    Columns(6).Cut
    Columns(1).Insert xlToRight
    
    '// 車種に「テスト」という文字列が入っているものを削除
    Range(Cells(startRow, 1), Cells(lastRow, lastColumn)).AutoFilter 1, "*テスト*", xlOr, "*ﾃｽﾄ*"
    If Cells(Rows.Count, 1).End(xlUp).Row > startRow Then
        sc.deleteAfterFilter Cells(startRow, 1)
    End If
    
    '// 最大積載量を数値に変換 & 最大積載量で昇順に並べ替え
    sc.toNumber 14, startRow, lastRow
    Columns(14).NumberFormatLocal = "#,###"
    sc.sortValues 14, Range(Cells(startRow, 1), Cells(lastRow, lastColumn))
    
    '// 日付を西暦から和暦に変換
    sc.changeDateFormat startRow + 1, lastRow, 9, "month"
    Cells(startRow, 9).Value = "初年度登録年月"
    sc.changeDateFormat startRow + 1, lastRow, 10, "day"
    Cells(startRow, 10).Value = "登録年月日"
    
    '// 車種を半角に統一
    sc.convertIntoLower startRow + 1, lastRow, 1
    Cells(startRow, 1).Value = "車種"
    
    Columns(2).Insert xlToRight
    Cells(startRow, 2).Value = "台数"
    
    '// YCLの分をYCLのシートに移動 & 「保存」ボタン追加
    sc.createSheet "YCL"
    
    Sheets("YCL").Activate
    Dim bc As New buttonController
    bc.addButton Sheets("YCL"), Sheets("YCL").Range(Cells(1, 1), Cells(2, 1)), "保存", "openForm"
    Set bc = Nothing
    
    Sheets("山岸運送").Activate
    
    Range(Cells(startRow, 1), Cells(lastRow, lastColumn)).AutoFilter 3, "YCL"
    Cells(startRow, 1).CurrentRegion.Copy Sheets("YCL").Cells(startRow, 1)
    sc.deleteAfterFilter Cells(startRow, 1)
    
    '// 山岸運送とYCL以外のものを削除
    Range(Cells(startRow, 1), Cells(lastRow, lastColumn)).AutoFilter 3, "<>山岸運送"
    sc.deleteAfterFilter Cells(startRow, 1)
    
    '// 山岸運送分のデータを一時的に保存するシート「山岸運送tmp」を作成 & データを「山岸運送tmp」へコピー
    sc.createSheet "山岸運送tmp"
    pasteToTmpSheet Sheets("山岸運送"), Sheets("山岸運送tmp"), startRow, lastRow
    
    '// YCL分のデータを一時的に保存するシート「YCLtmp」を作成 & データを「YCLtmp」へコピー
    sc.createSheet "YCLtmp"
    Call pasteToTmpSheet(Sheets("YCL"), Sheets("YCLtmp"), sc.getRow("start", 1, Sheets("YCL")), sc.getRow("last", 1, Sheets("YCL")))
    
    '// 車種ごとに分類した表を作成
    Call classifyTruck(Sheets("山岸運送"), Sheets("山岸運送tmp"), Sheets("設定(山岸運送)"), sc)
    Call classifyTruck(Sheets("YCL"), Sheets("YCLtmp"), Sheets("設定(YCL)"), sc)
    
    '// ヘッダーの罫線を太字に変更
    Call setHeaderLine(Sheets("山岸運送"), sc, xlMedium)
    Call setHeaderLine(Sheets("YCL"), sc, xlMedium)
    
    Sheets("山岸運送").Activate
    Cells(1, 1).Select
        
    Set sc = Nothing
    
    Application.DisplayAlerts = True
    
    MsgBox "処理が完了しました。", Title:=ThisWorkbook.Name
    
End Sub

'// 一時的に作成したシートに元のシートのデータをコピー(元のシートのヘッダーは残す)
Private Sub pasteToTmpSheet(targetSheet As Worksheet, tmpSheet As Worksheet, startRow As Long, lastRow As Long)

    With targetSheet
        .Activate
        .Cells.Copy Destination:=tmpSheet.Cells(1, 1)
        .Range(Cells(startRow + 1, 1), Cells(lastRow, lastRow)).Clear
    End With
        
End Sub

'/**
' *車種ごとに分類するサブルーチン
' *@params targetSheet 貼り付け先のシート
' *@params tmpSheet    データを車種ごとに分けるために一時的に作成するシート
' *@params configSheet 車種の設定等が書かれたシート
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
                
            '// 設定シートに入力された車種を格納した配列を作成
            For i = 0 To UBound(splitedTmpcell)
                ReDim Preserve truckTypes(UBound(truckTypes) + 1)
                truckTypes(UBound(truckTypes)) = splitedTmpcell(i)
            Next
            
            sc.divideTruck truckTypes, tmpSheet, targetSheet
        Next
    End With
    
    tmpSheet.Delete
    
    '//車番・車名の台数の後に移動
    With targetSheet
        .Activate
        .Range("H:I").Cut
        .Cells(3).Insert xlToRight
    End With
    
    '// 車種ごとの台数の表を車両一覧の下に作成
    sc.createNumberOfTrucksChart targetSheet, configSheet
    
    '// 車種ごとに分類する際に使用したfilterd列削除
    Dim startRow As Long: startRow = sc.getRow("start", 1, targetSheet)
    targetSheet.Columns(sc.getColumn("last", startRow, targetSheet)).Delete
    
    '// 車検有効期限列に条件付き書式設定
    Call setFormatCondition(targetSheet)
    
    '// ヘッダーのフォントサイズ変更・中央ぞろえ・シート全体のフォントをメイリオに変更・ウィンドウ枠の固定
    With targetSheet
        With .Rows(startRow)
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        End With
    
        .Range("A:A,C:I,P:P,V:AC").HorizontalAlignment = xlCenter
        .Range(.Cells(sc.getRow("start", 2, targetSheet) + 1, 2), .Cells(sc.getRow("last", 2, targetSheet), 2)).HorizontalAlignment = xlRight
    
        With .Cells
            .Font.Name = "メイリオ"
            .EntireColumn.AutoFit
        End With
        
        .Cells(4, 4).Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
    End With
    
End Sub

'/**
' * 車検有効期限列に条件付き書式設定
' * 期限切れ→赤,10日以内→黄色,30日以内→緑
' */
Private Sub setFormatCondition(targetSheet As Worksheet)

    With targetSheet.Range("P:P").FormatConditions
        .Delete
        
        '// 期限切れ
        Dim fcRed As FormatCondition
        Set fcRed = .Add(Type:=xlExpression, Formula1:="=DATEVALUE(P1) <TODAY()")
        fcRed.Interior.Color = RGB(178, 34, 34)
        Set fcRed = Nothing
        
        '// 10以内に到来
        Dim fcYellow As FormatCondition
        Set fcYellow = .Add(Type:=xlExpression, Formula1:="=AND(DATEVALUE(P1) >=TODAY(), DATEVALUE(P1) <=TODAY() + 10)")
        fcYellow.Interior.Color = RGB(255, 255, 102)
        Set fcYellow = Nothing
        
        '// 30日以内に到来
        Dim fcGreen As FormatCondition
        Set fcGreen = .Add(Type:=xlExpression, Formula1:="=AND(DATEVALUE(P1) > TODAY() + 10, DATEVALUE(P1) <=TODAY() + 30)")
        fcGreen.Interior.Color = RGB(50, 205, 50)
        Set fcGreen = Nothing
    End With
End Sub


'/**
 '* ヘッダーの罫線を太線に設定
 '**/
Private Sub setHeaderLine(targetSheet As Worksheet, sc As sheetController, lineWeight As Long)

    Dim startRow As Long: startRow = sc.getRow("start", 1, targetSheet)

    With targetSheet
        With .Range(.Cells(startRow, 1), .Cells(startRow, sc.getColumn("last", startRow, targetSheet)))
            .Borders.Weight = lineWeight
        End With
    End With

End Sub

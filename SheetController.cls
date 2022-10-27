VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sheetController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// シートやセルを操作するクラス
Option Explicit

'// 先頭(最終)行を取得
Public Function getRow(rowType As String, standardColumn As Long, targetSheet As Worksheet) As Long

    With targetSheet
    
        If rowType = "start" Then
            getRow = .Cells(1, standardColumn).End(xlDown).Row
        ElseIf rowType = "last" Then
            getRow = .Cells(Rows.Count, standardColumn).End(xlUp).Row
        End If
        
    End With
    
End Function

'// 先頭(最終)列を取得
Public Function getColumn(columnType As String, standardRow As Long, targetSheet As Worksheet) As Long
    
    With targetSheet
    
        If columnType = "start" Then
            getColumn = .Cells(standardRow, 1).End(xlToRight).Column
        ElseIf columnType = "last" Then
            getColumn = .Cells(standardRow, Columns.Count).End(xlToLeft).Column
        End If

    End With
    
End Function

'// シートが存在しなければ作成(あればクリア)
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

'// シートがあるかの判定
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
 '* 文字列を数値に変換
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
 '* データを並べ替え
 '* @params keyColumn 並べ替えのキー
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
 '* 日付の書式を和暦に変更
 '* @params startRow 開始行
 '* @params lastRow 最終行
 '* @params targetColumn 対象列
 '* @params dateType 月、日のどこまで表示するか指定(month,dayのどちらか)
 '**/
Public Sub changeDateFormat(startRow As Long, lastRow As Long, targetColumn As Long, dateType As String)

    Columns(targetColumn).Insert (xlToLeft)
    
    Dim columnIndexName As String: columnIndexName = Split(Columns(targetColumn + 1).Address, "$")(2)
    
    If dateType = "day" Then
        Cells(startRow, targetColumn).Formula = "=TEXT(DATE(LEFT(" & columnIndexName & startRow & ",4),MID(" & columnIndexName & startRow & " ,6,2),RIGHT(" & columnIndexName & startRow & ",2)),""ggge年mm月dd日"")"
    ElseIf dateType = "month" Then
        Cells(startRow, targetColumn).Formula = "=TEXT(DATE(LEFT(" & columnIndexName & startRow & ",4),MID(" & columnIndexName & startRow & " ,6,2),RIGHT(" & columnIndexName & startRow & ",2)),""ggge年mm月"")"
    End If
    
    Cells(startRow, targetColumn).AutoFill Range(Cells(startRow, targetColumn), Cells(lastRow, targetColumn)), xlFillValues
    
    Columns(targetColumn).Copy
    Columns(targetColumn).PasteSpecial xlPasteValues
    Columns(targetColumn + 1).Delete xlToLeft
    
End Sub
'/**
'* 指定の列の値を半角に変換
'* @params targetColumn 対象列
'**/
Public Sub convertIntoLower(startRow As Long, lastRow As Long, targetColumn As Long)

    Columns(targetColumn).Insert xlToRight
    
    Dim columnIndexName As String: columnIndexName = Split(Columns(targetColumn + 1).Address, "$")(2)
    
    Cells(startRow, targetColumn).Formula = "=SUBSTITUTE(ASC(" & columnIndexName & startRow & "),""増ﾄﾝ"",""増t"")"
    Cells(startRow, targetColumn).AutoFill Range(Cells(startRow, targetColumn), Cells(lastRow, targetColumn)), xlFillValues
    
    Columns(targetColumn).Copy
    Columns(targetColumn).PasteSpecial xlPasteValues
    Columns(targetColumn + 1).Delete xlToLeft

End Sub

'/**
' * フィルターをかけた後、抽出したものを削除する(フィルターがかかった状態で使用する)
' * @params standardCell: 基準となるセル
' * @params offsetRow: 下げる行数
'**/
Public Sub deleteAfterFilter(standardCell As Range)

    Dim targetRange As Range: Set targetRange = standardCell.CurrentRegion
    
    targetRange.Offset(1).Resize(targetRange.Rows.Count - 1).Delete (xlUp)
    standardCell.AutoFilter
    
    Set targetRange = Nothing
    
End Sub

'/**
 '* フィルターをかけた後、抽出したものを別シートに貼り付け
 '* @params standardCell: 基準となるセル
 '* @params targetSheet: 貼り付け先のシート
 '**/
 Private Sub pasteAfterFilter(standardCell As Range, standardSheet As Worksheet, targetSheet As Worksheet)
  
    Dim targetRange As Range: Set targetRange = standardCell.CurrentRegion
    targetRange.Offset(1).Resize(targetRange.Rows.Count - 1).Copy
    
    targetSheet.Cells(Me.getRow("last", 1, targetSheet) + 1, 1).PasteSpecial
    
    Set targetRange = Nothing
    
 End Sub

'// 車両の分類ごとに分けるサブルーチン
Public Sub divideTruck(trucktype() As Variant, tmpSheet As Worksheet, targetSheet As Worksheet)

    tmpSheet.Activate
    
    Dim startRow As Long: startRow = Me.getRow("start", 1, tmpSheet)
    Dim lastRow As Long: lastRow = Me.getRow("last", 1, tmpSheet)
    Dim lastColumn As Long: lastColumn = Me.getColumn("last", startRow, tmpSheet)
    Dim targetRange As Range: Set targetRange = Range(Cells(startRow, 1), Cells(lastRow, lastColumn))
    
    '// 指定の分類でフィルターをかける(truckTypeがその他の場合はフィルターかけない)
    If trucktype(1) = "その他" Then
        targetRange.AutoFilter 1, "<>"
    Else
        filterByTruckTypes trucktype, targetSheet, targetRange
    End If
    
    Me.sortValues 1, targetRange                           '// 車種名で並び替え
    
    '// フィルターをかけて何も引っかからなっかたら抜ける
    If Cells(Rows.Count, 1).End(xlUp).Row = startRow Then
        Cells(startRow, 1).AutoFilter
        Exit Sub
    End If
    
    '//車種の見出しセル作成
    setTruckTypeHeader trucktype(), targetSheet
        
    '// 抽出した分をtargetSheetに貼り付け & tmpSheetから削除
    Dim standardCell As Range: Set standardCell = Cells(Me.getRow("start", 1, tmpSheet), 1)
    Call pasteAfterFilter(standardCell, tmpSheet, targetSheet)
    Call deleteAfterFilter(standardCell)
    
    '// 車種ごとの1番下の行に2重下線
    Dim tmpLastRow As Long: tmpLastRow = Me.getRow("last", 1, targetSheet)
    
    With targetSheet
        .Activate
        .Range(.Cells(tmpLastRow, 1), .Cells(tmpLastRow, lastColumn)).Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    
    '// 通し番号付与
    Call countTrucks(Cells(startRow, 2), targetSheet)
    
End Sub

'/**
 '* 車種で絞る際、複数車種を一つのカテゴリーとする場合(3tと4tなど)
 '* 車種ごとにフィルターをかけて引っかかったものの最終列の1つ右の列の値を targeted にして、最後にtargetedのものでフィルターをかける
 '*/
Private Sub filterByTruckTypes(truckTypes() As Variant, targetSheet As Worksheet, targetRange As Range)

    Dim i As Long
    Dim tmpRange As Range
    
    For i = 0 To UBound(truckTypes)
        '// 車種が空白だったら次のループへ
        If truckTypes(i) = "" Then
            GoTo Continue
        End If
        
        targetRange.AutoFilter 1, "*" & StrConv(truckTypes(i), vbNarrow) & "*", xlOr, "*" & StrConv(truckTypes(i), vbWide) & "*"
        
        '// フィルターをかけて何も引っかからなければ次のループへ
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

'// 車種の先頭の見出し設定
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

'// 車種の見出し名作成
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
            truckTypeHeader = truckTypeHeader & trucktype(i) & "・"
        End If

Continue:
    Next
    
    createTruckTypeHeader = truckTypeHeader
    
End Function

'// 車種ごとに通し番号を付与
Private Sub countTrucks(standardCell As Range, targetSheet As Worksheet)

    Dim startRow As Long: startRow = Me.getRow("last", 2, targetSheet) + 2
    Dim lastRow As Long: lastRow = Me.getRow("last", 1, targetSheet)
    
    Dim i As Long
    Dim counter As Long: counter = 1

    For i = startRow To lastRow
        targetSheet.Cells(i, 2).Value = counter
        counter = counter + 1
    Next
    
    '// 車種の台数の合計をヘッダーの横に入力(下の表で使用するため)
    targetSheet.Cells(startRow - 1, 2).Value = counter - 1
    
End Sub

'// 車両一覧の下に各車種ごとの台数を表示した表を作成
Public Sub createNumberOfTrucksChart(targetSheet As Worksheet, configSheet As Worksheet)

    targetSheet.Activate

    Cells(Me.getRow("last", 1, targetSheet) + 2, 1).Value = "車種ごとの台数"

    Dim i As Long
    
    Dim truckTypes() As Variant
    Dim truckTypeHeader As String
    
    Dim targetRow As Long
    
    Dim numberOfTrucks As Long
    Dim totalAmount As Long
    
    For i = 2 To Me.getRow("last", 1, configSheet)
        truckTypes = splitToArray(Split(configSheet.Cells(i, 1), ","))
        truckTypeHeader = createTruckTypeHeader(truckTypes)                     '// 車種名
        
        If truckTypeHeader = "" Then
            GoTo Continue
        End If
        
        '// 対象の見出しが書かれたセルの行番号(見つからない場合はエラーが発生するのでこの部分はエラー無視)
        On Error Resume Next
        targetRow = 0
        
        targetRow = WorksheetFunction.Match(truckTypeHeader, targetSheet.Columns(1), 0)
        
        On Error GoTo 0
        
        If targetRow = 0 Then
            GoTo Continue
        End If
        
        numberOfTrucks = Cells(targetRow, 2).Value                                         '// 車種ごとの合計台数
        
        With Cells(Me.getRow("last", 1, targetSheet) + 1, 1)
            .Value = truckTypeHeader
            .Offset(, 1).Value = numberOfTrucks & "台"
        End With
        
        Cells(targetRow, 2).Value = ""
        
        totalAmount = totalAmount + numberOfTrucks

Continue:
    Next
        
    With Cells(Me.getRow("last", 1, targetSheet) + 1, 1)
        .Interior.Color = RGB(144, 238, 144)
        .Value = "合計"
        
        With .Offset(, 1)
            .Value = totalAmount & "台"
            .Interior.Color = RGB(144, 238, 144)
        End With
    End With
            
End Sub

'// split関数で分けた値を配列に変換
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


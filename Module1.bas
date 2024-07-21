Attribute VB_Name = "Module1"
Public Const SETTING_START_COL = 4       ' 設定シートの開始行番号
Public Const SETTING_HEADER_ROW = 2      ' 設定シートのタスク一覧の見出し行の列番号
Public Const SETTING_TASK_START_ROW = 3  ' 設定シートのタスク一覧の開始行の列番号
Public Const SETTING_TASK_START_COL = 4  ' 設定シートのタスク一覧の開始列の列番号
Public Const SETTING_TASK_END_COL = 5    ' 設定シートのタスク一覧の終了列の列番号
Public Const SETTING_TASK_PRIOR_COL = 6  ' 設定シートのタスク一覧の並び替え優先項目列の列番号


' タスク一覧をソートする
Sub タスクのソート()
    Dim headerRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As String
    Dim endCol As String
    Dim initSelection As Range
    Dim priorColNames() As String
    Dim targetCol As String
    Dim i As Long
    
    ' ソート実行前の準備
    Application.ScreenUpdating = False                                                    ' スクリーンの更新を止める
    Set initSelection = Selection                                                         ' 現在のセル選択位置を退避する
    headerRow = Worksheets("設定").Cells(SETTING_START_COL, SETTING_HEADER_ROW).value     ' 設定シートの見出し行番号を取得
    startRow = Worksheets("設定").Cells(SETTING_START_COL, SETTING_TASK_START_ROW).value  ' 設定シートの開始行番号を取得
    startCol = Worksheets("設定").Cells(SETTING_START_COL, SETTING_TASK_START_COL).value  ' 設定シートの開始列の取得を取得
    endCol = Worksheets("設定").Cells(SETTING_START_COL, SETTING_TASK_END_COL).value      ' 設定シートの終了列の取得を取得
    Range(startCol & startRow).Select                                                     ' 開始セルの選択
    endRow = Selection.End(xlDown).Row                                                    ' 最終行の取得
    
    ' データの並び替え優先度を設定する
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Clear  ' ソートの設定を初期化
    priorColNames = Split(GetPriorColNames(), ",")
    For i = LBound(priorColNames) To UBound(priorColNames)
        targetCol = FindColumn(headerRow, priorColNames(i))
        ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Add _
            Key:=Range(targetCol & startRow & ":" & targetCol & endRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
    Next i
    
    ' データの並び替え
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SetRange Range(GetNextColumn(startCol) & startRow & ":" & endCol & endRow)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' No.の番号をオートフィルで設定しなおす
    Range(startCol & startRow & ":" & startCol & startRow + 1).Select
    Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown)), Type:=xlFillValues
    
    ' ソート実行後の後処理
    initSelection.Select               ' 初期のセル選択位置に戻す
    Application.ScreenUpdating = True  ' スクリーンの更新を再開する
End Sub


' ソートを優先する項目名の取得
Function GetPriorColNames() As String
    Dim startRow As Long
    Dim i As Integer
    Dim value As String
    
    i = SETTING_START_COL
    Do While Worksheets("設定").Cells(i, SETTING_TASK_PRIOR_COL).value <> ""   ' 設定シートの並び替え優先項目の開始行を取得
        GetPriorColNames = GetPriorColNames & "," & Worksheets("設定").Cells(i, SETTING_TASK_PRIOR_COL).value
        i = i + 1 ' 次の行に移動
    Loop
    
    ' 先頭の","を消す
    GetPriorColNames = Mid(GetPriorColNames, 2)
End Function


' 指定文字列を引数の行番号を始点に検索し、見つかった行番号を返す
Function FindColumn(ByVal rowNum As Long, ByVal searchStr As String) As String
    Dim i As Long
    Dim lastColumn As Long
    
    ' 最終列の取得
    lastColumn = Cells(rowNum, Columns.Count).End(xlToLeft).Column
    
    ' 指定文字列を探す
    For i = 1 To lastColumn
        If Cells(rowNum, i).value = searchStr Then
            FindColumn = Split(Cells(rowNum, i).Address, "$")(1)  ' 列番号の取得
            Exit Function                                         ' 対象の列が見つかった場合は関数を終了
        End If
    Next i
    
    ' 対象の列が見つからなかった場合はエラー終了
    Err.Raise 9999, , "指定の文字列が見つかりませんでした"
End Function


' 次の列名を取得する
Function GetNextColumn(ByVal columnStr As String) As String
    Dim currentColumun As Range
    Dim nextColumun As Range
    
    Set currentColumn = Range(columnStr & "1")         ' 指定した列の1行目のセルを取得
    Set nextColumn = currentColumn.Offset(0, 1)        ' 指定した列の次の列のセル名を取得
    GetNextColumn = Split(nextColumn.Address, "$")(1)  ' 指定した次の列の列名を取得する
End Function


' 新しい行を一番下に追加する
Sub 行の追加()
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As String
    Dim endCol As String
    Dim networkdaysCol As String
    Dim initSelection As Range
    
    ' 行追加前の準備
    Application.ScreenUpdating = False                                                    ' スクリーンの更新を止める
    Set initSelection = Selection                                                         ' 現在のセル選択位置を退避する
    headerRow = Worksheets("設定").Cells(SETTING_START_COL, SETTING_HEADER_ROW).value     ' 設定シートの見出し行番号を取得
    startRow = Worksheets("設定").Cells(SETTING_START_COL, SETTING_TASK_START_ROW).value  ' 設定シートの開始行番号を取得
    startCol = Worksheets("設定").Cells(SETTING_START_COL, SETTING_TASK_START_COL).value  ' 設定シートの開始列の取得を取得
    Range(startCol & startRow).Select                                                     ' 開始セルの選択
    endRow = Selection.End(xlDown).Row                                                    ' 最終行の取得
    ActiveSheet.Outline.ShowLevels columnlevels:=2                                        ' グループ化された列を必ず表示させる
    
    '新しい行の挿入
    Set ws = ActiveSheet                                    ' アクティブなシートを取得
    ws.Rows(endRow + 1).Insert Shift:=xlDown                ' 新しい行を挿入
    ws.Rows(endRow).EntireRow.Copy                          ' 前の行をコピー
    ws.Rows(endRow + 1).PasteSpecial Paste:=xlPasteFormats  ' 前の行の書式を新しい行にコピー
    
    ' 「No.」の列は前の行に＋1する
    Range(startCol & endRow + 1).value = Range(startCol & endRow).value + 1
    
    ' 「日数」の列は前の行の数式をコピーする
    networkdaysCol = FindColumn(headerRow, "日数")                          ' 日数列の取得
    Range(networkdaysCol & endRow).Copy                                     ' 追加前の最終行をコピー
    Range(networkdaysCol & endRow + 1).PasteSpecial Paste:=xlPasteFormulas  ' 追加行に前行の数式をコピーする
    
    ' 行追加後の後処理
    Application.CutCopyMode = False    ' クリップボードをクリアする
    initSelection.Select               ' 初期のセル選択位置に戻す
    Application.ScreenUpdating = True  ' スクリーンの更新を再開する
End Sub

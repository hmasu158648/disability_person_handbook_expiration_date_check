Attribute VB_Name = "Module1"

' 有効期限が指定月数以内または過ぎた利用者を抽出するマクロ
Sub 有効期限抽出()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim i As Long
    Dim expiryDate As Date
    Dim today As Date
    Dim targetDate As Date
    Dim expiryValue As Variant
    Dim statusValue As String
    Dim extractCount As Long
    Dim monthsValue As Variant
    Dim months As Long
    Dim targetSheetName As String
    Dim ws As Worksheet
    Dim calcMode As XlCalculation
    
    ' 現在の計算モードを保存
    calcMode = Application.Calculation
    
    ' 画面更新を停止
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' エラーハンドリング
    On Error GoTo ErrorHandler
    
    ' シートを設定
    Set wsSource = ThisWorkbook.Worksheets("Sheet1")
    
    ' J7セルから月数を取得
    monthsValue = wsSource.Range("K7").Value
    
    ' 月数の妥当性チェック
    If IsNumeric(monthsValue) And monthsValue > 0 Then
        months = CLng(monthsValue)
    Else
        ' 無効な値の場合は警告して処理を中止
        MsgBox "セルJ7に有効な月数（1以上の数値）を入力してください。", vbExclamation, "入力エラー"
        GoTo CleanupExit
    End If
    
    ' 今日の日付を取得
    today = Date
    ' 指定月数後の日付を計算
    targetDate = DateAdd("m", months, today)
    
    ' ターゲットシート名を動的に生成
    targetSheetName = "有効期限" & months & "か月以内一覧"
    
    ' 既存の類似シートを削除（他の月数のシート）
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "有効期限*か月以内一覧" And ws.Name <> targetSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    On Error GoTo 0
    
    ' ターゲットシートが存在するか確認、なければ作成
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets(targetSheetName)
    On Error GoTo 0
    
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Worksheets.Add(After:=wsSource)
        wsTarget.Name = targetSheetName
    Else
        ' 既存のシートをクリア（ヘッダー行以外、罫線や書式も含む）
        If wsTarget.Cells(1, 1).Value <> "" Then
            wsTarget.Rows("2:" & wsTarget.Rows.Count).Clear
        End If
    End If
    
    ' ヘッダー行をコピー（値と書式をコピー、条件付き書式の色は静的な色として適用）
    wsSource.Rows(1).Copy
    wsTarget.Rows(1).PasteSpecial xlPasteValuesAndNumberFormats
    wsTarget.Rows(1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
        
    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' ターゲットシートの開始行
    targetRow = 2
    
    ' データを1行ずつチェック
    For i = 2 To lastRow
        ' 状況列（I列）をチェック - 「退職」は除外
        statusValue = Trim(wsSource.Cells(i, "I").Value)
        If statusValue = "退職" Or statusValue = "確認済み" Then
            GoTo NextIteration
        End If
        
        ' 有効期限列（E列）の値を取得
        expiryValue = wsSource.Cells(i, "E").Value
        
        ' 「無」は除外
        If expiryValue = "無" Then
            GoTo NextIteration
        End If
        
        ' 日付形式かチェック
        If IsDate(expiryValue) Then
            expiryDate = CDate(expiryValue)
            
            ' 有効期限が今日から指定月数以内、または過ぎている場合
            If expiryDate <= targetDate Then
                ' A列からI列までをコピー（値と書式をコピー、条件付き書式の色は静的な色として適用）
                wsSource.Range("A" & i & ":I" & i).Copy
                wsTarget.Range("A" & targetRow).PasteSpecial xlPasteValuesAndNumberFormats
                wsTarget.Range("A" & targetRow).PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
                
                ' E列の条件付き書式による表示色を通常の背景色として適用
                wsTarget.Cells(targetRow, 5).Interior.Color = _
                    wsSource.Cells(i, 5).DisplayFormat.Interior.Color
                
                ' F列の条件付き書式によるフォント色を通常のフォント色として適用
                wsTarget.Cells(targetRow, 6).Font.Color = _
                    wsSource.Cells(i, 6).DisplayFormat.Font.Color
                
                targetRow = targetRow + 1
            End If
        End If
        
NextIteration:
    Next i
    
    ' 抽出件数を計算（ヘッダー行と1行進んだ分を引く）
    extractCount = targetRow - 2
    
    ' 抽出件数に応じて処理を分岐
    If extractCount > 0 Then
        ' 列幅を自動調整
        wsTarget.Columns.AutoFit
        
        ' オートフィルターを設定（設定されていない場合のみ）
        If Not wsTarget.AutoFilterMode Then
            wsTarget.Range("A1").AutoFilter
        End If
        
        ' 所属列（C列）で昇順にソート
        wsTarget.Sort.SortFields.Clear
        wsTarget.Sort.SortFields.Add Key:=wsTarget.Range("C1"), SortOn:=xlSortOnValues, Order:=xlAscending
        With wsTarget.Sort
            .SetRange wsTarget.Range("A1:I" & targetRow - 1)
            .Header = xlYes
            .Apply
        End With

        ' ターゲットシートをアクティブにする
        wsTarget.Activate
        
        ' 画面更新を再開してシートを表示
        Application.ScreenUpdating = True

        ' 完了メッセージ
        MsgBox "抽出が完了しました。" & vbCrLf & _
               "設定月数: " & months & "か月以内" & vbCrLf & _
               "抽出件数: " & extractCount & "件", vbInformation, "完了"

        
    Else
        ' 0件の場合はソースシートをアクティブにする
        wsSource.Activate
        
        ' 画面更新を再開してシートを表示
        Application.ScreenUpdating = True
        
        ' 該当者なしのメッセージ
        MsgBox "有効期限が" & months & "か月以内または過ぎている利用者は見つかりませんでした。", vbInformation, "抽出結果"
    End If
    
CleanupExit:
    ' 設定を元に戻す
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' エラー時も設定を元に戻す
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
    
End Sub

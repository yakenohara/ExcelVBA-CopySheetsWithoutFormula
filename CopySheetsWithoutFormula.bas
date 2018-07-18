Attribute VB_Name = "CopySheetsWithoutFormula"
'
'選択しているシート(複数選択可)の
'値コピーで作った新しいブックを作る
'
Sub CopySheetsWithoutFormula()
    
    '変数宣言
    Dim isFirstCopy As Boolean

    Dim activeWnw As Window

    Dim sourceBk As Workbook
    Dim sourceSht As Worksheet

    Dim newBk As Workbook
    
    Dim toActivateShtName As String
    
    Dim shtNameArr() As String
    Dim indexOfShtNameArr As Long
    Dim toRestoreShtNameArr() As String

    Dim beforeTmpSecond As Date
    Dim nowTmpSecond As Date

    beforeTmpSecond = Time

    '処理
    
    'アプリ設定
    Application.ScreenUpdating = False
    
    Set sourceBk = ActiveWorkbook
    
    'シート選択状態復元時の選択シートの保存
    toActivateShtName = sourceBk.ActiveSheet.Name
        
    '選択シートリストを作成
    Set activeWnw = Application.ActiveWindow
    indexOfShtNameArr = 0
    For Each s In activeWnw.SelectedSheets
        ReDim Preserve shtNameArr(indexOfShtNameArr)
        shtNameArr(indexOfShtNameArr) = s.Name
        indexOfShtNameArr = indexOfShtNameArr + 1
    Next s

    Set newBk = Workbooks.Add
    
    'シートコピー
    isFirstCopy = True
    cnt = 1
    cntMx = UBound(shtNameArr) + 1
    For Each shtName In shtNameArr
        
        '進捗状態表示
        caption = "Processing... Sheet[" & cnt & "(" & shtName & ")/" & cntMx & "]"
        Application.StatusBar = caption
        cnt = cnt + 1
        
        Set sourceSht = sourceBk.Worksheets(shtName)
        
        If (isFirstCopy) Then
            
            sourceSht.Copy Before:=newBk.Sheets(1)
            Call pasteValues(sourceBk, sourceSht, newBk, newBk.Sheets(1), caption)
            
            '新ブック生成時のデフォルトシートの削除
            numOfShts = newBk.Sheets.count
            For i = 2 To numOfShts
                
                Application.DisplayAlerts = False
                newBk.Sheets(i).Delete
                Application.DisplayAlerts = True
                
            Next
            
            isFirstCopy = False
            
        Else '2シート目以降のコピーの時

            sourceSht.Copy After:=newBk.Sheets(newBk.Sheets.count)
            Call pasteValues(sourceBk, sourceSht, newBk, newBk.Sheets(newBk.Sheets.count), caption)
        
        End If

        '応答無し回避
        nowTmpSecond = Time
        If (nowTmpSecond <> beforeTmpSecond) Then
            DoEvents
        End If
        beforeTmpSecond = nowTmpSecond
        
    Next shtName

    '選択シート状態の復元用配列の生成
    ReDim toRestoreShtNameArr(UBound(shtNameArr))
    toRestoreShtNameArr(0) = toActivateShtName
    indexOfToRestoreShtNameArr = 1
    For Each shtName In shtNameArr
        If (shtName <> toActivateShtName) Then
            toRestoreShtNameArr(indexOfToRestoreShtNameArr) = shtName
            indexOfToRestoreShtNameArr = indexOfToRestoreShtNameArr + 1
        End If
        
    Next shtName
    
    '選択シート状態の復元
    sourceBk.Activate
    sourceBk.Sheets(toRestoreShtNameArr).Select
    newBk.Activate
    newBk.Sheets(toRestoreShtNameArr).Select
    
    'アプリ設定の復元
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Done!"
    
End Sub

'
'シート間で値コピーする
'
Private Sub pasteValues(ByVal sourceBk As Workbook, ByVal sourceSht As Worksheet, ByVal destBk As Workbook, ByVal destSht As Worksheet, ByVal caption As String)
    
    Dim sourceSlection As String
    Dim sourceRng As String
    
    Application.StatusBar = caption

    'コピー元シートの選択
    sourceBk.Activate
    sourceSht.Select

    'コピー元シートのセル選択状態の保存
    sourceSlection = Selection.Address 'セル選択状態の保存
    sourceRng = sourceSht.UsedRange.Address

    sourceSht.Range(sourceRng).Copy

    'コピー先シートの選択
    destBk.Activate
    destSht.Activate
    destSht.Range(sourceRng).Select
    
    '貼り付け
On Error GoTo ERROR_ '結合セルが存在する場合は、Errorが発生するので、 ERROR_:でキャッチする
    Selection.PasteSpecial Paste:=xlPasteValues, _
                           Operation:=xlNone, _
                           SkipBlanks:=False, _
                           Transpose:=False
    
    Application.CutCopyMode = False '範囲コピー状態を解除
    
    'セル選択状態の復元
    sourceBk.Activate
    sourceSht.Activate
    sourceSht.Range(sourceSlection).Select
    destSht.Activate
    destSht.Activate
    destSht.Range(sourceSlection).Select
    
    Application.StatusBar = False

    Exit Sub

ERROR_:  '結合セルが存在する場合
    Err.Clear
    On Error GoTo 0 'On Error GoTo~ の解除
    Call breakFormula(sourceSht, destSht, caption)
    Resume Next

End Sub

'
'コピー先シートの数式を値に変更する
'
Private Sub breakFormula(ByVal sourceSht As Worksheet, ByVal destSht As Worksheet, ByVal caption As String)
    
    Dim processedCount As Long
    Dim MxOfProcessedCount As Long
    Dim toProcessAddress As String
    
    Dim beforeTmpSecond As Date
    Dim nowTmpSecond As Date

    beforeTmpSecond = Time
    
    toProcessAddress = sourceSht.UsedRange.Address
    MxOfProcessedCount = sourceSht.UsedRange.count
    processedCount = 0
    
    For Each Rng In sourceSht.UsedRange
        
        processedCount = processedCount + 1
        Application.StatusBar = caption & ", Cell[" & processedCount & "/" & MxOfProcessedCount & "(" & toProcessAddress & ")]"
        
        If (Rng.HasFormula = True) Then '数式の場合
            destSht.Range(Rng.Address).Value = Rng.Value '値に変更する
            
        End If
        
        '応答無し回避
        nowTmpSecond = Time
        If (nowTmpSecond <> beforeTmpSecond) Then
            DoEvents
        End If
        beforeTmpSecond = nowTmpSecond
        
    Next Rng
    
    Application.StatusBar = False
    
End Sub


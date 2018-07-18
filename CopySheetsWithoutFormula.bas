Attribute VB_Name = "値コピーでシートコピー"
Sub 値コピーでシートコピー()
    
    '変数宣言
    Dim activeWnw As Window
    Dim newBk As Workbook
    Dim isFirstCopy As Boolean
    Dim formulaRange As Range
    
    Application.ScreenUpdating = False
    
    Set activeWnw = Application.ActiveWindow
    Set newBk = Workbooks.Add
    
    'シートコピー
    isFirstCopy = True
    For Each s In activeWnw.SelectedSheets
        
        If (isFirstCopy) Then
            
            s.Copy Before:=newBk.Sheets(1)
            Call breakFormula(newBk.Sheets(1), s)
            
            'デフォルトシートの削除
            numOfShts = newBk.Sheets.count
            For i = 2 To numOfShts
                
                Application.DisplayAlerts = False
                newBk.Sheets(i).Delete
                Application.DisplayAlerts = True
                
            Next
            
            isFirstCopy = False
            
        Else
        
            s.Copy After:=newBk.Sheets(newBk.Sheets.count)
            Call breakFormula(newBk.Sheets(newBk.Sheets.count), s)
        
        End If
        
    
    Next s
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Done!"
    
End Sub

'
'コピー先シートの数式を値に変更する
'
Private Function breakFormula(ByRef destSht As Worksheet, ByVal sourceSht As Worksheet)
    
    Dim count As Long: count = 0
    Dim countMax As Long: countMax = 3000
    
    For Each Rng In sourceSht.UsedRange
            
        Application.StatusBar = "Checking formula Sheet[" & sourceSht.Name & "], Cell[" & Rng.Address & "]"
        
        If (Rng.HasFormula = True) Then
            
            destSht.Range(Rng.Address).Value = Rng.Value
            
        End If
        
        count = count + 1
        
        If count > countMax Then
            
            DoEvents '応答無し回避
            count = 0
            
        End If
        
    Next Rng
    
End Function

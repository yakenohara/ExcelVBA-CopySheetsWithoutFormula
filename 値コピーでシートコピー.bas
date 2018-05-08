Attribute VB_Name = "�l�R�s�[�ŃV�[�g�R�s�["
Sub �l�R�s�[�ŃV�[�g�R�s�[()
    
    '�ϐ��錾
    Dim activeWnw As Window
    Dim newBk As Workbook
    Dim isFirstCopy As Boolean
    Dim formulaRange As Range
    
    Application.ScreenUpdating = False
    
    Set activeWnw = Application.ActiveWindow
    Set newBk = Workbooks.Add
    
    '�V�[�g�R�s�[
    isFirstCopy = True
    For Each s In activeWnw.SelectedSheets
        
        If (isFirstCopy) Then
            
            s.Copy Before:=newBk.Sheets(1)
            Call breakFormula(newBk.Sheets(1), s)
            
            '�f�t�H���g�V�[�g�̍폜
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
'�R�s�[��V�[�g�̐�����l�ɕύX����
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
            
            DoEvents '�����������
            count = 0
            
        End If
        
    Next Rng
    
End Function

Attribute VB_Name = "CopySheetsWithoutFormula"
'<License>------------------------------------------------------------
'
' Copyright (c) 2018 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
'�I�����Ă���V�[�g(�����I����)��
'�l�R�s�[�ō�����V�����u�b�N�����
'
Sub CopySheetsWithoutFormula()
    
    '�ϐ��錾
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

    '����
    
    '�A�v���ݒ�
    Application.ScreenUpdating = False
    
    Set sourceBk = ActiveWorkbook
    
    '�V�[�g�I����ԕ������̑I���V�[�g�̕ۑ�
    toActivateShtName = sourceBk.ActiveSheet.Name
        
    '�I���V�[�g���X�g���쐬
    Set activeWnw = Application.ActiveWindow
    indexOfShtNameArr = 0
    For Each s In activeWnw.SelectedSheets
        ReDim Preserve shtNameArr(indexOfShtNameArr)
        shtNameArr(indexOfShtNameArr) = s.Name
        indexOfShtNameArr = indexOfShtNameArr + 1
    Next s

    Set newBk = Workbooks.Add
    
    '�V�[�g�R�s�[
    isFirstCopy = True
    cnt = 1
    cntMx = UBound(shtNameArr) + 1
    For Each shtName In shtNameArr
        
        '�i����ԕ\��
        caption = "Processing... Sheet[" & cnt & "(" & shtName & ")/" & cntMx & "]"
        Application.StatusBar = caption
        cnt = cnt + 1
        
        Set sourceSht = sourceBk.Worksheets(shtName)
        
        If (isFirstCopy) Then
            
            sourceSht.Copy Before:=newBk.Sheets(1)
            Call pasteValues(sourceBk, sourceSht, newBk, newBk.Sheets(1), caption)
            
            '�V�u�b�N�������̃f�t�H���g�V�[�g�̍폜
            numOfShts = newBk.Sheets.count
            For i = 2 To numOfShts
                
                Application.DisplayAlerts = False
                newBk.Sheets(i).Delete
                Application.DisplayAlerts = True
                
            Next
            
            isFirstCopy = False
            
        Else '2�V�[�g�ڈȍ~�̃R�s�[�̎�

            sourceSht.Copy After:=newBk.Sheets(newBk.Sheets.count)
            Call pasteValues(sourceBk, sourceSht, newBk, newBk.Sheets(newBk.Sheets.count), caption)
        
        End If

        '�����������
        nowTmpSecond = Time
        If (nowTmpSecond <> beforeTmpSecond) Then
            DoEvents
        End If
        beforeTmpSecond = nowTmpSecond
        
    Next shtName

    '�I���V�[�g��Ԃ̕����p�z��̐���
    ReDim toRestoreShtNameArr(UBound(shtNameArr))
    toRestoreShtNameArr(0) = toActivateShtName
    indexOfToRestoreShtNameArr = 1
    For Each shtName In shtNameArr
        If (shtName <> toActivateShtName) Then
            toRestoreShtNameArr(indexOfToRestoreShtNameArr) = shtName
            indexOfToRestoreShtNameArr = indexOfToRestoreShtNameArr + 1
        End If
        
    Next shtName
    
    '�I���V�[�g��Ԃ̕���
    sourceBk.Activate
    sourceBk.Sheets(toRestoreShtNameArr).Select
    newBk.Activate
    newBk.Sheets(toRestoreShtNameArr).Select
    
    '�A�v���ݒ�̕���
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Done!"
    
End Sub

'
'�V�[�g�ԂŒl�R�s�[����
'
Private Sub pasteValues(ByVal sourceBk As Workbook, ByVal sourceSht As Worksheet, ByVal destBk As Workbook, ByVal destSht As Worksheet, ByVal caption As String)
    
    Dim sourceSlection As String
    Dim sourceRng As String
    
    Application.StatusBar = caption

    '�R�s�[���V�[�g�̑I��
    sourceBk.Activate
    sourceSht.Select

    '�R�s�[���V�[�g�̃Z���I����Ԃ̕ۑ�
    sourceSlection = Selection.Address '�Z���I����Ԃ̕ۑ�
    sourceRng = sourceSht.UsedRange.Address

    sourceSht.Range(sourceRng).Copy

    '�R�s�[��V�[�g�̑I��
    destBk.Activate
    destSht.Activate
    destSht.Range(sourceRng).Select
    
    '�\��t��
On Error GoTo ERROR_ '�����Z�������݂���ꍇ�́AError����������̂ŁA ERROR_:�ŃL���b�`����
    Selection.PasteSpecial Paste:=xlPasteValues, _
                           Operation:=xlNone, _
                           SkipBlanks:=False, _
                           Transpose:=False
    
    Application.CutCopyMode = False '�͈̓R�s�[��Ԃ�����
    
    '�Z���I����Ԃ̕���
    sourceBk.Activate
    sourceSht.Activate
    sourceSht.Range(sourceSlection).Select
    destSht.Activate
    destSht.Activate
    destSht.Range(sourceSlection).Select
    
    Application.StatusBar = False

    Exit Sub

ERROR_:  '�����Z�������݂���ꍇ
    Err.Clear
    On Error GoTo 0 'On Error GoTo~ �̉���
    Call breakFormula(sourceSht, destSht, caption)
    Resume Next

End Sub

'
'�R�s�[��V�[�g�̐�����l�ɕύX����
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
        
        If (Rng.HasFormula = True) Then '�����̏ꍇ
            destSht.Range(Rng.Address).Value = Rng.Value '�l�ɕύX����
            
        End If
        
        '�����������
        nowTmpSecond = Time
        If (nowTmpSecond <> beforeTmpSecond) Then
            DoEvents
        End If
        beforeTmpSecond = nowTmpSecond
        
    Next Rng
    
    Application.StatusBar = False
    
End Sub


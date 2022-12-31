
 
Public Sub 正方形長方形1_Click()
On Error GoTo 正方形長方形1_Click_Err

    Dim strCopiedFolderName As String

    '画面の更新を止める
    Application.ScreenUpdating = False
    
    strCopiedFolderName = ThisWorkbook.Path & "\コピー"
    
    'フォルダコピー
    If FolderCopy(ThisWorkbook.Path, strCopiedFolderName) = False Then
        Call MsgBox(FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If

    'CSVファイル出力
    If CSVFilesToExcelFiles(ThisWorkbook.Path, strCopiedFolderName) = False Then
        Call MsgBox(FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
    'Excelファイル統合
    If DeleteAllFiles(strCopiedFolderName, "accdb") = False Then
        Call MsgBox(FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
    'Excelファイル統合
    If CombineAllExcelFile(ThisWorkbook.Path, strCopiedFolderName) = False Then
        Call MsgBox(FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
        
正方形長方形1_Click_Err:

正方形長方形1_Click_Exit:
    '画面の更新を再開する
    Application.ScreenUpdating = False
End Sub

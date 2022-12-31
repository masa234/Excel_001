Public Const DATA_SHEET_NAME = "データ"
Public Const FAILED = "失敗しました。"
Public Const CONFIRM = "確認"

'【概要】シートが存在するか
Public Function IsExistsSheet(ByVal objWb As Excel.Workbook, _
                                ByVal strSheetName As String) As Boolean
On Error GoTo IsExistsSheet_Err

    IsExistsSheet = False
    
    Dim objWs As Excel.Worksheet
    
    '引数のシート名でシートオブジェクトを参照する
    'シートが存在しない場合、エラーが発生する
    Set objWs = objWb.Worksheets(strSheetName)
    
    IsExistsSheet = True
    
IsExistsSheet_Err:

IsExistsSheet_Exit:
    Set objWs = Nothing
End Function


'【概要】シート削除
Public Function SheetDelete(ByVal objWb As Excel.Workbook, _
                                ByVal strSheetName As String) As Boolean
On Error GoTo SheetDelete_Err

    SheetDelete = False
    
    '警告メッセージをFalseにする
    Application.DisplayAlerts = False
    
    'シート削除
    objWb.Worksheets(strSheetName).Delete
    
    SheetDelete = True
    
SheetDelete_Err:

SheetDelete_Exit:
    '警告メッセージをTrueにする
    Application.DisplayAlerts = True
End Function


'【概要】ファイル削除
Public Function FileDelete(ByVal strDeleteFilePath) As Boolean
On Error GoTo FileDelete_Err

    FileDelete = False
    
    'ファイル削除
    Kill strDeleteFilePath
    
    FileDelete = True
    
FileDelete_Err:

FileDelete_Exit:

End Function


'【概要】ファイル名変更
Public Function RenameFile(ByVal strFilePath As String, _
                        ByVal strRenameFileName As String) As Boolean
On Error GoTo RenameFile_Err

    RenameFile = False
    
    Dim objFso As FileSystemObject
    
    'Fsoを呼びだす
    Set objFso = New FileSystemObject
    
    '名前変更
    objFso.GetFile(strFilePath).Name = strRenameFileName
    
    RenameFile = True
    
RenameFile_Err:

RenameFile_Exit:
    Set objFso = Nothing
End Function


'【概要】特定のディレクトリの特定の拡張子のファイル群を取得する
Public Function GetAllFilePaths(ByVal strDirectoryPath As String, _
                                ByVal strExtensionName As String) As Variant
On Error GoTo GetAllFilePaths_Err
    
    Dim lngArrIdx As Long
    Dim arrRet() As Variant
    Dim objFso As FileSystemObject
    Dim objFile As File
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    With objFso
        'ファイルの数だけ繰り返す
        For Each objFile In .GetFolder(strDirectoryPath).Files
            'ファイルの拡張子が指定の者だった場合
            If .GetExtensionName(objFile.Path) = strExtensionName Then
                '配列再宣言
                ReDim Preserve arrRet(lngArrIdx)
                '配列格納
                arrRet(lngArrIdx) = objFile.Path
                '配列の要素番号を1つ進める
                lngArrIdx = lngArrIdx + 1
            End If
        Next objFile
    End With
    
    GetAllFilePaths = arrRet
    
GetAllFilePaths_Err:

GetAllFilePaths_Exit:
    Set objFso = Nothing
    Set objFile = Nothing
End Function


'【概要】特定のディレクトリの特定の拡張子のファイルを削除する
Public Function DeleteAllFiles(ByVal strDirectoryPath As String, _
                            ByVal strExtensionName As String) As String
On Error GoTo DeleteAllFiles_Err

    DeleteAllFiles = False
    
    Dim objFso As FileSystemObject
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    With objFso
        'ファイルの数だけ繰り返す
        For Each objFile In .GetFolder(strDirectoryPath).Files
            'ファイルの拡張子が指定のものだった場合、
            If .GetExtensionName(objFile.Name) = strExtensionName Then
                'ファイルを削除する
                If FileDelete(objFile.Path) = False Then
                    GoTo DeleteAllFiles_Exit
                End If
            End If
        Next objFile
    End With
    
    DeleteAllFiles = True
    
DeleteAllFiles_Err:

DeleteAllFiles_Exit:
    Set objFso = Nothing
End Function


'【概要】配列の内容をExcelファイルとして出力する
Public Function ArrToExcelFile(ByVal arrOutput As Variant, _
                            ByVal strOutputExcelFilePath As String) As Boolean
On Error GoTo ArrToExcelFile_Err

    ArrToExcelFile = False
    
    Dim lngArrIdx As Long
    Dim lngCurrentRow As Long
    Dim strSaveBookName As String
    Dim objWb As Excel.Workbook
    
    'Excelファイル作成
    Set objWb = Workbooks.Add
    
    'シート名称をDATAにする
    ActiveSheet.Name = DATA_SHEET_NAME

    '列初期化
    lngCurrentRow = 1
    
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrOutput)
        '出力
        objWb.Worksheets(DATA_SHEET_NAME).Cells(lngCurrentRow, 1).Value = arrOutput(lngArrIdx)
        '行をカウントアップ
        lngCurrentRow = lngCurrentRow + 1
    Next lngArrIdx
    
    '保存するファイル名
    strSaveBookName = strOutputExcelFilePath & "\" & GetFileNameWithDate(strOutputExcelFilePath, DATA_SHEET_NAME) & ".xlsx"
    
    'ブックを閉じる
    If CloseBook(objWb, strSaveBookName, True) = False Then
        GoTo ArrToExcelFile_Exit
    End If
        
    ArrToExcelFile = True
    
ArrToExcelFile_Err:

ArrToExcelFile_Exit:
    Set objWb = Nothing
End Function


'【概要】ブックを閉じる
Public Function CloseBook(ByVal objWb As Excel.Workbook, _
                        ByVal strSaveBookName As String, _
                        Optional ByVal boolSave As Boolean = False) As Boolean
On Error GoTo CloseBook_Err

    CloseBook = False
    
    '保存が有効な場合
    If boolSave Then
        objWb.SaveAs
    End If
    
    'ブックを閉じる
    objWb.Close
    
    CloseBook = True
    
CloseBook_Err:

CloseBook_Exit:

End Function


'【概要】フォルダコピー
Public Function FolderCopy(ByVal strCopyFolderPath As String, _
                        ByVal strCopiedFilePath As String) As Boolean
On Error GoTo FolderCopy_Err

    FolderCopy = False
    
    Dim objFso As FileSystemObject

    'Fsoを呼びだす
    Set objFso = New FileSystemObject
    
    'Fsoのコピーフォルダを使う
    objFso.CopyFolder strCopyFolderPath, strCopiedFilePath
    
    FolderCopy = True
    
FolderCopy_Err:

FolderCopy_Exit:
    Set objFso = Nothing
End Function


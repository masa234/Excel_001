

'【概要】次の連番シート名取得
Public Function GetSheetNameWithSeqNumber(ByVal objWb As Excel.Workbook, _
                            ByVal strBaseSheetName As String) As String
On Error GoTo GetSheetNameWithSeqNumber_Err

    Dim lngCount As Long
    Dim strSheetName As String

    '100回繰り返す
    For lngCount = 1 To 100
        'シート名設定
        strSheetName = strBaseSheetName & "_" & CStr(lngCount)
        'シートが存在する場合、処理を終了する
        If IsExistsSheet(objWb, strSheetName) = False Then
            GetSheetNameWithSeqNumber = strSheetName
            GoTo GetSheetNameWithSeqNumber_Exit
        End If
    Next lngCount
    
GetSheetNameWithSeqNumber_Err:

GetSheetNameWithSeqNumber_Exit:

End Function


'【概要】Excelファイルへの転記
Public Function CombineAllExcelFile(ByVal strDirectoryPath As String, _
                                    ByVal strExcelFilePath As String) As Boolean
On Error GoTo CombineAllExcelFile_Err

    CombineAllExcelFile = False
    
    Dim arrExcelFilePaths() As Variant
    
    'Excelファイル群を取得
    arrExcelFilePaths = GetAllFilePaths(strDirectoryPath, "xlsm")
    
    'Excelファイルを出力する
    If ExcelFilesToExcelFile(arrExcelFilePaths, strExcelFilePath, DATA_SHEET_NAME) = False Then
        GoTo CombineAllExcelFile_Exit
    End If
    
    CombineAllExcelFile = True
    
CombineAllExcelFile_Err:

CombineAllExcelFile_Exit:

End Function


'【概要】Excelファイル（複数）からExcelファイルへの転記
Public Function ExcelFilesToExcelFile(ByVal arrExcelFilePaths As Variant, _
                                    ByVal strDirectoryPath As String, _
                                    ByVal strBaseSheetName As String) As Boolean
On Error GoTo ExcelFilesToExcelFile_Err

    ExcelFilesToExcelFile = False
    
    Dim lngArrIdx As Long
    Dim lngLastRow As Long
    Dim lngPasteRow As Long
    Dim lngPastedRow As Long
    Dim strFilePath As String
    Dim strRenameFilePath As String
    Dim strPastedSheetName As String
    Dim objPasteWb As Excel.Workbook
    Dim objPastedWb As Excel.Workbook
    Dim objPasteWs As Excel.Worksheet
    
    'Excelファイルのパスの数だけ繰り返す
    For lngArrIdx = 0 To UBound(arrExcelFilePaths)
        'Excelファイルを開く
        Workbooks.Open arrExcelFilePaths(lngArrIdx)
        '貼り付け元ブック
        Set objPasteWb = ActiveWorkbook
        '新規でExcelファイルを作成する
        Set objPastedWb = Workbooks.Add
        'シート名変更
        ActiveSheet.Name = DATA_SHEET_NAME
        '貼り付け先シート名設定
        strPastedSheetName = DATA_SHEET_NAME
        '貼り付け元のシートの数だけ繰り返す
        For Each objPasteWs In objPasteWb.Worksheets
            With objPasteWb.Worksheets(objPasteWs.Name)
                '値が設定されていない場合
                If .Cells(1, 1) <> vbNullString Then
                    'シートを削除する
                    If SheetDelete(objPasteWb, objPasteWb.Name) = False Then
                        GoTo ExcelFilesToExcelFile_Exit
                    End If
                    '次のシートへ
                    GoTo nextSheet
                End If
                '貼り付け元最終行取得
                lngLastRow = .Cells(1, 1).End(xlDown).Row
                '列を初期化
                lngPastedRow = 1
                '最終行まで繰り返す
                For lngPasteRow = 1 To lngLastRow
                    '貼り付け元→貼り付け先
                    objPastedWb.Worksheets(strPastedSheetName).Cells(lngPastedRow, 1).Value = .Cells(lngPasteRow, 1).Value
                    '貼り付け先行が100の超える場合
                    If lngPastedRow = 100 Then
                        '次のシート名取得
                        strPastedSheetName = GetSheetNameWithSeqNumber(objPastedWb, strBaseSheetName)
                        'シート追加
                        objPastedWb.Worksheets.Add
                        'シート名変更
                        ActiveSheet.Name = strPastedSheetName
                        '初期化
                        lngPastedRow = 1
                    End If
                    '貼り付け先行をカウントアップ
                    lngPastedRow = lngPastedRow + 1
                Next lngPasteRow
            End With
nextSheet:
        Next objPasteWs
    Next lngArrIdx
    
    ExcelFilesToExcelFile = True
    
ExcelFilesToExcelFile_Err:

ExcelFilesToExcelFile_Exit:
    Set objPasteWb = Nothing
    Set objPastedWb = Nothing
    Set objPasteWs = Nothing
End Function


'【概要】yyyymmdd付きの連番ファイル名を取得する
Public Function GetFileNameWithDate(ByVal strDirectoryPath As String, _
                                    ByVal strFileName As String) As String
On Error GoTo GetFileNameWithDate_Err

    GetFileNameWithDate = False
    
    Dim lngCount As Long
    Dim strChkFileName As String
    Dim objFso As FileSystemObject
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    '100回繰り返す
    For lngCount = 1 To 100
        strChkFileName = strFileName & Format(Date, "yyyy_mm_dd") & "_" & CStr(lngCount)
        'ファイルが存在する限り繰り返す
        If objFso.FileExists(strDirectoryPath & "\" & strChkFileName) = False Then
            'ファイル名を設定
            GetFileNameWithDate = strChkFileName
            GoTo GetFileNameWithDate_Exit
        End If
    Next lngCount
    
    GetFileNameWithDate = True
    
GetFileNameWithDate_Err:

GetFileNameWithDate_Exit:
    Set objFso = Nothing
End Function


'【概要】CSVのデータを配列で取得する
Public Function GetCSVData(ByVal strCSVFilePath As String) As Variant
On Error GoTo GetCSVData_Err
    
    Dim lngFreeFile As Long
    Dim lngArrIdx As Long
    Dim strLine As String
    Dim arrRet() As Variant
    
    'フリーファイルを取得
    lngFreeFile = FreeFile
    
    'CSVファイルを開く
    Open strCSVFilePath For Input As #lngFreeFile
    
    '終端まで繰り返す
    Do Until EOF(lngFreeFile)
        '1行読み込み
        Line Input #lngFreeFile, strLine
        '配列再宣言
        ReDim Preserve arrRet(lngArrIdx)
        '配列に格納
        arrRet(lngArrIdx) = strLine
        '配列の要素番号を1つ進める
        lngArrIdx = lngArrIdx + 1
    Loop
        
    GetCSVData = arrRet
    
GetCSVData_Err:

GetCSVData_Exit:
    'CSVファイルを閉じる
    Close #lngFreeFile
End Function


'【概要】カレントディレクトリのCSVファイルをExcelファイルとして出力する
Public Function CSVFilesToExcelFiles(ByVal strDirectoryPath As String, _
                                    ByVal strExcelFilePath As String) As Boolean
On Error GoTo CSVFilesToExcelFiles_Err

    CSVFilesToExcelFiles = False
    
    Dim lngArrIdx As Long
    Dim arrCSVFilePaths() As Variant
    Dim arrCSVData() As Variant
    
    'CSVファイル群を取得
    arrCSVFilePaths = GetAllFilePaths(strDirectoryPath, "csv")
    
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrCSVFilePaths)
        'CSVファイルを配列に格納
        arrCSVData = GetCSVData(arrCSVFilePaths(lngArrIdx))
        '配列をExcelファイルとして出力
        If ArrToExcelFile(arrCSVData, strExcelFilePath) = False Then
            GoTo CSVFilesToExcelFiles_Exit
        End If
    Next lngArrIdx
    
    CSVFilesToExcelFiles = True
    
CSVFilesToExcelFiles_Err:

CSVFilesToExcelFiles_Exit:

End Function

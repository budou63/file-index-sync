Attribute VB_Name = "modDummyExport"
Option Explicit

Private Const DUMMY_PROCESS_NAME As String = "ダミーブック出力"
Private Const DUMMY_SUFFIX_PREFIX As String = "_ダミー_"
Private Const DUMMY_LOG_SHEET As String = "ログ"
Private Const MAX_HEADER_SCAN_ROWS As Long = 30
Private Const MAX_DUMMY_PREFIX_LENGTH As Long = 20

Public Sub ダミーブック出力()
    Dim sourceBook As Workbook
    Dim dummyBook As Workbook
    Dim folderPath As String
    Dim outputPath As String
    Dim baseName As String
    Dim extensionName As String
    Dim replacedCellCount As Long
    Dim processedSheetCount As Long
    Dim skippedSheetCount As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim answer As VbMsgBoxResult
    Dim currentStage As String
    Dim openedDummy As Boolean

    On Error GoTo ErrorHandler

    Set sourceBook = ThisWorkbook
    If Len(sourceBook.Path) = 0 Then
        MsgBox "元ブックが未保存のため、ダミーブックを出力できません。" & vbCrLf & _
               "先にこのブックを保存してから再実行してください。", vbExclamation, DUMMY_PROCESS_NAME
        Exit Sub
    End If

    answer = MsgBox("現在のブックを複製し、ダミーデータ版を別ファイルとして出力します。" & vbCrLf & _
                    "元のブックは変更されません。" & vbCrLf & _
                    "実行しますか？", vbQuestion Or vbYesNo Or vbDefaultButton2, DUMMY_PROCESS_NAME)
    If answer <> vbYes Then Exit Sub

    folderPath = SelectOutputFolder()
    If Len(folderPath) = 0 Then
        MsgBox "保存先フォルダの選択がキャンセルされました。", vbInformation, DUMMY_PROCESS_NAME
        Exit Sub
    End If

    baseName = GetWorkbookBaseName(sourceBook.Name)
    extensionName = GetWorkbookExtension(sourceBook.Name)
    If Len(extensionName) = 0 Then extensionName = ".xlsm"
    outputPath = BuildPath(folderPath, baseName & DUMMY_SUFFIX_PREFIX & Format(Now, "yyyymmdd_hhnnss") & extensionName)

    If Len(Dir(outputPath)) > 0 Then
        MsgBox "出力先に同名ファイルが既に存在します。" & vbCrLf & outputPath, vbExclamation, DUMMY_PROCESS_NAME
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    currentStage = "ブック複製"
    sourceBook.SaveCopyAs outputPath

    currentStage = "複製ブックを開く"
    Set dummyBook = Application.Workbooks.Open(Filename:=outputPath, UpdateLinks:=False, ReadOnly:=False, AddToMru:=False)
    openedDummy = True

    currentStage = "ダミーデータ化"
    DummyizeWorkbook dummyBook, processedSheetCount, skippedSheetCount, replacedCellCount
    If processedSheetCount = 0 Then
        Err.Raise vbObjectError + 5201, , "見出しを判定できるシートが一つもありません。"
    End If

    currentStage = "ログ記録"
    WriteDummyExportLog dummyBook, outputPath, folderPath, processedSheetCount, skippedSheetCount, replacedCellCount

    currentStage = "保存"
    dummyBook.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    dummyBook.Close SaveChanges:=False
    openedDummy = False

    MsgBox "ダミーブックを出力しました。" & vbCrLf & vbCrLf & _
           "保存先：" & vbCrLf & outputPath & vbCrLf & vbCrLf & _
           "置換セル数：" & Format(replacedCellCount, "#,##0") & "件", vbInformation, DUMMY_PROCESS_NAME

Cleanup:
    On Error Resume Next
    If openedDummy Then dummyBook.Close SaveChanges:=False
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Dim errMessage As String
    errMessage = "ダミーブック出力中にエラーが発生しました。" & vbCrLf & vbCrLf & _
                 "処理段階：" & currentStage & vbCrLf & _
                 "内容：" & Err.Description
    MsgBox errMessage, vbExclamation, DUMMY_PROCESS_NAME
    Resume Cleanup
End Sub

Private Sub DummyizeWorkbook(ByVal targetBook As Workbook, ByRef processedSheetCount As Long, ByRef skippedSheetCount As Long, ByRef replacedCellCount As Long)
    Dim excludeSheets As Variant
    Dim ws As Worksheet
    Dim sheetReplacementCount As Long

    excludeSheets = Array("ログ", "設定", "マスタ")

    For Each ws In targetBook.Worksheets
        If IsExcludedSheet(ws.Name, excludeSheets) Then
            skippedSheetCount = skippedSheetCount + 1
        Else
            sheetReplacementCount = DummyizeWorksheet(ws)
            If sheetReplacementCount >= 0 Then
                processedSheetCount = processedSheetCount + 1
                replacedCellCount = replacedCellCount + sheetReplacementCount
            Else
                skippedSheetCount = skippedSheetCount + 1
            End If
        End If
    Next ws
End Sub

Private Function DummyizeWorksheet(ByVal ws As Worksheet) As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim rowIndex As Long
    Dim headerText As String
    Dim dummyPrefix As String
    Dim valueText As String
    Dim map As Object
    Dim mapKey As String
    Dim nextIndex As Long
    Dim replacedCount As Long
    Dim cell As Range

    headerRow = DetectHeaderRow(ws)
    If headerRow = 0 Then
        DummyizeWorksheet = -1
        Exit Function
    End If

    lastRow = GetLastUsedRowLocal(ws)
    lastCol = GetLastUsedColLocal(ws)
    If lastRow <= headerRow Or lastCol = 0 Then
        DummyizeWorksheet = -1
        Exit Function
    End If

    For col = 1 To lastCol
        headerText = CleanHeaderText(CStr(ws.Cells(headerRow, col).Value))
        If Len(headerText) > 0 Then
            dummyPrefix = BuildDummyPrefix(headerText)
            Set map = CreateObject("Scripting.Dictionary")
            nextIndex = 1

            For rowIndex = headerRow + 1 To lastRow
                Set cell = ws.Cells(rowIndex, col)
                If ShouldDummyizeCell(cell) Then
                    valueText = CStr(cell.Value)
                    mapKey = valueText
                    If Not map.Exists(mapKey) Then
                        map.Add mapKey, dummyPrefix & Format(nextIndex, "000")
                        nextIndex = nextIndex + 1
                    End If
                    cell.Value = map(mapKey)
                    replacedCount = replacedCount + 1
                End If
            Next rowIndex
        End If
    Next col

    DummyizeWorksheet = replacedCount
End Function

Private Function DetectHeaderRow(ByVal ws As Worksheet) As Long
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim bestRow As Long
    Dim bestScore As Long
    Dim score As Long
    Dim scanLastRow As Long

    On Error Resume Next
    If ws.ListObjects.Count > 0 Then
        For Each lo In ws.ListObjects
            If Not lo.HeaderRowRange Is Nothing Then
                DetectHeaderRow = lo.HeaderRowRange.Row
                Exit Function
            End If
        Next lo
    End If
    On Error GoTo 0

    If ws.AutoFilterMode Then
        On Error Resume Next
        If Not ws.AutoFilter.Range Is Nothing Then
            DetectHeaderRow = ws.AutoFilter.Range.Row
            Exit Function
        End If
        On Error GoTo 0
    End If

    scanLastRow = Application.WorksheetFunction.Min(GetLastUsedRowLocal(ws), MAX_HEADER_SCAN_ROWS)
    For rowIndex = 1 To scanLastRow
        score = HeaderRowScore(ws, rowIndex)
        If score > bestScore Then
            bestScore = score
            bestRow = rowIndex
        End If
    Next rowIndex

    If bestScore >= 2 Then DetectHeaderRow = bestRow
End Function

Private Function HeaderRowScore(ByVal ws As Worksheet, ByVal rowIndex As Long) As Long
    Dim lastCol As Long
    Dim col As Long
    Dim textCount As Long
    Dim knownCount As Long
    Dim valueText As String

    lastCol = GetLastUsedColLocal(ws)
    For col = 1 To lastCol
        valueText = CleanHeaderText(CStr(ws.Cells(rowIndex, col).Value))
        If Len(valueText) > 0 And Not IsNumericOnlyText(valueText) Then
            textCount = textCount + 1
            If IsKnownHeader(valueText) Then knownCount = knownCount + 1
        End If
    Next col

    HeaderRowScore = textCount
    If knownCount > 0 Then HeaderRowScore = HeaderRowScore + knownCount * 3
End Function

Private Function ShouldDummyizeCell(ByVal cell As Range) As Boolean
    Dim v As Variant
    Dim textValue As String

    If cell.HasFormula Then Exit Function
    v = cell.Value
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If IsDate(v) Then Exit Function
    If IsNumeric(v) Then Exit Function

    textValue = Trim$(CStr(v))
    If Len(textValue) = 0 Then Exit Function
    If UCase$(textValue) = "TRUE" Or UCase$(textValue) = "FALSE" Then Exit Function
    If IsNumericOnlyText(textValue) Then Exit Function
    If IsDateOnlyText(textValue) Then Exit Function
    If IsYearOnlyText(textValue) Then Exit Function
    If IsEraYearOnlyText(textValue) Then Exit Function

    ShouldDummyizeCell = True
End Function

Private Function IsKnownHeader(ByVal textValue As String) As Boolean
    Dim headers As Variant
    Dim i As Long

    headers = Array("通し番号", "キャビネット番号", "第1ガイド", "第2ガイド", "個別フォルダー", _
                    "内容・取扱い", "備考", "ファイルID", "ファイル番号", "サブタイトル", _
                    "背表紙用タイトル", "公開用タイトル", "作成者名", "登録日時", "更新日時")
    For i = LBound(headers) To UBound(headers)
        If InStr(1, textValue, CStr(headers(i)), vbTextCompare) > 0 Then
            IsKnownHeader = True
            Exit Function
        End If
    Next i
End Function

Private Function IsNumericOnlyText(ByVal textValue As String) As Boolean
    IsNumericOnlyText = RegexTest(textValue, "^[-+]?\d+(\.\d+)?$")
End Function

Private Function IsDateOnlyText(ByVal textValue As String) As Boolean
    If RegexTest(textValue, "^\d{4}[/-]\d{1,2}[/-]\d{1,2}$") Then IsDateOnlyText = True
End Function

Private Function IsYearOnlyText(ByVal textValue As String) As Boolean
    IsYearOnlyText = RegexTest(textValue, "^\d{4}(年度|年)?$")
End Function

Private Function IsEraYearOnlyText(ByVal textValue As String) As Boolean
    IsEraYearOnlyText = RegexTest(textValue, "^(令和|平成|昭和)\d{1,2}(年度|年)?$|^R\d{1,2}$|^R\d{2}$")
End Function

Private Function RegexTest(ByVal textValue As String, ByVal pattern As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    RegexTest = re.Test(Trim$(textValue))
End Function

Private Function BuildDummyPrefix(ByVal headerText As String) As String
    Dim s As String
    s = CleanHeaderText(headerText)
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    s = Replace(s, "（", "")
    s = Replace(s, "）", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    If Len(s) > MAX_DUMMY_PREFIX_LENGTH Then s = Left$(s, MAX_DUMMY_PREFIX_LENGTH)
    If Len(s) = 0 Then s = "ダミー"
    BuildDummyPrefix = s
End Function

Private Function CleanHeaderText(ByVal textValue As String) As String
    CleanHeaderText = Trim$(Replace(Replace(textValue, vbCr, ""), vbLf, ""))
End Function

Private Function SelectOutputFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ダミーブックの保存先フォルダを選択してください"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        SelectOutputFolder = .SelectedItems(1)
    End With
End Function

Private Function GetWorkbookBaseName(ByVal fileName As String) As String
    Dim dotPosition As Long
    dotPosition = InStrRev(fileName, ".")
    If dotPosition > 0 Then
        GetWorkbookBaseName = Left$(fileName, dotPosition - 1)
    Else
        GetWorkbookBaseName = fileName
    End If
End Function

Private Function GetWorkbookExtension(ByVal fileName As String) As String
    Dim dotPosition As Long
    dotPosition = InStrRev(fileName, ".")
    If dotPosition > 0 Then GetWorkbookExtension = Mid$(fileName, dotPosition)
End Function

Private Function BuildPath(ByVal folderPath As String, ByVal fileName As String) As String
    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        BuildPath = folderPath & fileName
    Else
        BuildPath = folderPath & Application.PathSeparator & fileName
    End If
End Function

Private Function IsExcludedSheet(ByVal sheetName As String, ByVal excludeSheets As Variant) As Boolean
    Dim i As Long
    For i = LBound(excludeSheets) To UBound(excludeSheets)
        If StrComp(sheetName, CStr(excludeSheets(i)), vbTextCompare) = 0 Then
            IsExcludedSheet = True
            Exit Function
        End If
    Next i
End Function

Private Function GetLastUsedRowLocal(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not foundCell Is Nothing Then GetLastUsedRowLocal = foundCell.Row
End Function

Private Function GetLastUsedColLocal(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not foundCell Is Nothing Then GetLastUsedColLocal = foundCell.Column
End Function

Private Sub WriteDummyExportLog(ByVal targetBook As Workbook, ByVal outputPath As String, ByVal folderPath As String, ByVal processedSheetCount As Long, ByVal skippedSheetCount As Long, ByVal replacedCellCount As Long)
    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim lastCell As Range

    On Error Resume Next
    Set wsLog = targetBook.Worksheets(DUMMY_LOG_SHEET)
    If wsLog Is Nothing Then
        Set wsLog = targetBook.Worksheets.Add(After:=targetBook.Worksheets(targetBook.Worksheets.Count))
        wsLog.Name = DUMMY_LOG_SHEET
    End If
    On Error GoTo 0
    If wsLog Is Nothing Then Exit Sub

    If Application.WorksheetFunction.CountA(wsLog.Range("A1:G1")) = 0 Then
        wsLog.Range("A1:G1").Value = Array("実行日時", "処理名", "出力ファイル名", "保存先", "処理対象シート数", "スキップシート数", "置換セル数")
    End If

    Set lastCell = wsLog.Range("A:G").Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        nextRow = 2
    Else
        nextRow = lastCell.Row + 1
        If nextRow < 2 Then nextRow = 2
    End If

    wsLog.Cells(nextRow, 1).Value = Now
    wsLog.Cells(nextRow, 2).Value = DUMMY_PROCESS_NAME
    wsLog.Cells(nextRow, 3).Value = Dir(outputPath)
    wsLog.Cells(nextRow, 4).Value = folderPath
    wsLog.Cells(nextRow, 5).Value = processedSheetCount
    wsLog.Cells(nextRow, 6).Value = skippedSheetCount
    wsLog.Cells(nextRow, 7).Value = replacedCellCount
End Sub

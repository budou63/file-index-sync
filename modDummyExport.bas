Attribute VB_Name = "modDummyExport"
Option Explicit

Private Const DUMMY_PROCESS_NAME As String = "ダミーブック出力"
Private Const DUMMY_SUFFIX As String = "_ダミー_"
Private Const MAX_HEADER_LENGTH As Long = 20

Public Sub ExportDummyWorkbook()
    Dim sourceBook As Workbook
    Dim dummyBook As Workbook
    Dim folderPath As String
    Dim outputPath As String
    Dim replacedCellCount As Long
    Dim processedSheetCount As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldStatusBar As Variant
    Dim copyOpened As Boolean

    On Error GoTo ErrorHandler

    Set sourceBook = ThisWorkbook
    If Len(sourceBook.Path) = 0 Then
        MsgBox "元ブックが未保存のため、ダミーブックを出力できません。" & vbCrLf & _
               "先にこのブックを保存してから再実行してください。", vbExclamation, DUMMY_PROCESS_NAME
        Exit Sub
    End If

    If MsgBox("現在のブックを複製し、文字列をダミー化した別ブックを出力します。" & vbCrLf & _
              "元のブックは変更されません。" & vbCrLf & _
              "実行しますか？", vbQuestion Or vbYesNo Or vbDefaultButton2, DUMMY_PROCESS_NAME) <> vbYes Then
        Exit Sub
    End If

    folderPath = SelectOutputFolder()
    If Len(folderPath) = 0 Then Exit Sub

    outputPath = BuildOutputPath(folderPath, sourceBook.Name)
    If Len(Dir(outputPath)) > 0 Then
        MsgBox "出力先に同名ファイルが既に存在します。" & vbCrLf & outputPath, vbExclamation, DUMMY_PROCESS_NAME
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "ダミーデータを作成しています..."

    sourceBook.SaveCopyAs outputPath
    Set dummyBook = Application.Workbooks.Open(Filename:=outputPath, UpdateLinks:=False, ReadOnly:=False, AddToMru:=False)
    copyOpened = True

    DummyizeWorkbook dummyBook, processedSheetCount, replacedCellCount

    dummyBook.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    dummyBook.Close SaveChanges:=False
    copyOpened = False

    MsgBox "ダミーブックを出力しました。" & vbCrLf & vbCrLf & _
           "保存先：" & vbCrLf & outputPath & vbCrLf & vbCrLf & _
           "置換セル数：" & Format(replacedCellCount, "#,##0") & "件" & vbCrLf & _
           "処理シート数：" & Format(processedSheetCount, "#,##0") & "枚", vbInformation, DUMMY_PROCESS_NAME

Cleanup:
    On Error Resume Next
    If copyOpened Then dummyBook.Close SaveChanges:=False
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Application.StatusBar = oldStatusBar
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    MsgBox "ダミーブック出力中にエラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号：" & Err.Number & vbCrLf & _
           "内容：" & Err.Description, vbExclamation, DUMMY_PROCESS_NAME
    Resume Cleanup
End Sub

Private Sub DummyizeWorkbook(ByVal targetBook As Workbook, ByRef processedSheetCount As Long, ByRef replacedCellCount As Long)
    Dim excludeSheets As Variant
    Dim ws As Worksheet
    Dim sheetReplaceCount As Long

    excludeSheets = Array("ログ", "設定", "マスタ")

    For Each ws In targetBook.Worksheets
        If Not IsExcludedSheet(ws.Name, excludeSheets) Then
            sheetReplaceCount = DummyizeWorksheetUsedRange(ws)
            If sheetReplaceCount >= 0 Then
                processedSheetCount = processedSheetCount + 1
                replacedCellCount = replacedCellCount + sheetReplaceCount
            End If
        End If
    Next ws
End Sub

Private Function DummyizeWorksheetUsedRange(ByVal ws As Worksheet) As Long
    Dim targetRange As Range
    Dim valueArray As Variant
    Dim formulaArray As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim r As Long
    Dim c As Long
    Dim headerText As String
    Dim dummyPrefix As String
    Dim dictionaries() As Object
    Dim nextNumbers() As Long
    Dim originalText As String
    Dim replacedCount As Long

    Set targetRange = ws.UsedRange
    If Application.WorksheetFunction.CountA(targetRange) = 0 Then
        DummyizeWorksheetUsedRange = -1
        Exit Function
    End If

    rowCount = targetRange.Rows.Count
    colCount = targetRange.Columns.Count
    If rowCount < 2 Or colCount < 1 Then
        DummyizeWorksheetUsedRange = -1
        Exit Function
    End If

    valueArray = targetRange.Value2
    formulaArray = targetRange.Formula
    ReDim dictionaries(1 To colCount)
    ReDim nextNumbers(1 To colCount)

    For c = 1 To colCount
        headerText = CStr(valueArray(1, c))
        dummyPrefix = BuildDummyPrefix(headerText)
        Set dictionaries(c) = CreateObject("Scripting.Dictionary")
        nextNumbers(c) = 1

        For r = 2 To rowCount
            If ShouldDummyizeValue(valueArray(r, c), formulaArray(r, c)) Then
                originalText = CStr(valueArray(r, c))
                If Not dictionaries(c).Exists(originalText) Then
                    dictionaries(c).Add originalText, dummyPrefix & Format(nextNumbers(c), "000")
                    nextNumbers(c) = nextNumbers(c) + 1
                End If
                formulaArray(r, c) = dictionaries(c)(originalText)
                replacedCount = replacedCount + 1
            End If
        Next r
    Next c

    targetRange.Formula = formulaArray
    DummyizeWorksheetUsedRange = replacedCount
End Function

Private Function ShouldDummyizeValue(ByVal cellValue As Variant, ByVal cellFormula As Variant) As Boolean
    Dim textValue As String

    If IsError(cellValue) Then Exit Function
    If Len(CStr(cellFormula)) > 0 Then
        If Left$(CStr(cellFormula), 1) = "=" Then Exit Function
    End If
    If IsEmpty(cellValue) Then Exit Function
    If VarType(cellValue) = vbBoolean Then Exit Function
    If IsNumeric(cellValue) Then Exit Function
    If IsDate(cellValue) Then Exit Function

    textValue = Trim$(CStr(cellValue))
    If Len(textValue) = 0 Then Exit Function
    If UCase$(textValue) = "TRUE" Or UCase$(textValue) = "FALSE" Then Exit Function
    If IsDateOnlyText(textValue) Then Exit Function
    If IsYearOnlyText(textValue) Then Exit Function

    ShouldDummyizeValue = True
End Function

Private Function IsDateOnlyText(ByVal textValue As String) As Boolean
    IsDateOnlyText = RegexTest(textValue, "^\d{4}[/-]\d{1,2}[/-]\d{1,2}$")
End Function

Private Function IsYearOnlyText(ByVal textValue As String) As Boolean
    IsYearOnlyText = RegexTest(textValue, "^\d{4}年度$|^(令和|平成|昭和)\d{1,2}(年度|年)$|^R\d{1,2}$")
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
    Dim prefix As String

    prefix = Trim$(Replace(Replace(CStr(headerText), vbCr, ""), vbLf, ""))
    prefix = Replace(prefix, " ", "")
    prefix = Replace(prefix, "　", "")
    If Len(prefix) = 0 Then prefix = "ダミー"
    If Len(prefix) > MAX_HEADER_LENGTH Then prefix = Left$(prefix, MAX_HEADER_LENGTH)
    BuildDummyPrefix = prefix
End Function

Private Function SelectOutputFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ダミーブックの保存先フォルダを選択してください"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        SelectOutputFolder = .SelectedItems(1)
    End With
End Function

Private Function BuildOutputPath(ByVal folderPath As String, ByVal originalFileName As String) As String
    Dim baseName As String
    Dim outputFileName As String

    baseName = GetBaseName(originalFileName)
    outputFileName = baseName & DUMMY_SUFFIX & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"

    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        BuildOutputPath = folderPath & outputFileName
    Else
        BuildOutputPath = folderPath & Application.PathSeparator & outputFileName
    End If
End Function

Private Function GetBaseName(ByVal fileName As String) As String
    Dim dotPosition As Long

    dotPosition = InStrRev(fileName, ".")
    If dotPosition > 0 Then
        GetBaseName = Left$(fileName, dotPosition - 1)
    Else
        GetBaseName = fileName
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

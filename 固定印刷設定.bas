Attribute VB_Name = "固定印刷設定"
Option Explicit

'========================================================
' 土木課C4476R 固定印刷プロファイル
' - Windows APIはこの標準モジュールに分離する。
' - 固定設定は各PC/各WindowsユーザーのLOCALAPPDATAへ保存する。
' - ExcelのActivePrinterポート（Ne03:等）は固定せず、Windows登録プリンターから解決する。
'========================================================

Private Const FIXED_PRINTER_NAME As String = "土木課C4476R"
Private Const FIXED_PRINTER_SERVER As String = "PRTSV03.sojanet.local"
Private Const SINGLE_SEAL_SHEET As String = "個別フォルダシール"
Private Const MULTI_SEAL_SHEET As String = "個別フォルダシール（複数）"
Private Const PROFILE_FORMAT_VERSION As String = "1"
Private Const PROFILE_BASE_FOLDER As String = "file-index-sync\固定印刷設定"
Private Const PROFILE_FILE_NAME As String = "土木課C4476R.dat"
Private Const SEAL_PER_PAGE_LOCAL As Long = 12
Private Const SEAL_PAGE_ROWS_LOCAL As Long = 20
Private Const PROFILE_LINE_SEPARATOR As String = vbLf
Private Const PROFILE_KEY_VALUE_SEPARATOR As String = "="
Private Const DM_OUT_BUFFER As Long = 2
Private Const DM_IN_PROMPT As Long = 4
Private Const DM_IN_BUFFER As Long = 8
Private Const IDOK As Long = 1
Private Const IDCANCEL As Long = 2
Private Const SEAL_PRINT_MODE_UNSET As Long = 0
Private Const SEAL_PRINT_MODE_SINGLE As Long = 1
Private Const SEAL_PRINT_MODE_MULTI As Long = 2

' 現在のExcelセッション中だけ、最後に正常完了したシール反映モードを保持する。
Private gSealPrintMode As Long

#If VBA7 Then
    Private Declare PtrSafe Function OpenPrinterW Lib "winspool.drv" (ByVal pPrinterName As LongPtr, ByRef phPrinter As LongPtr, ByVal pDefault As LongPtr) As Long
    Private Declare PtrSafe Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
    Private Declare PtrSafe Function GetPrinterW Lib "winspool.drv" (ByVal hPrinter As LongPtr, ByVal Level As Long, ByVal pPrinter As LongPtr, ByVal cbBuf As Long, ByRef pcbNeeded As Long) As Long
    Private Declare PtrSafe Function SetPrinterW Lib "winspool.drv" (ByVal hPrinter As LongPtr, ByVal Level As Long, ByVal pPrinter As LongPtr, ByVal Command As Long) As Long
    Private Declare PtrSafe Function DocumentPropertiesW Lib "winspool.drv" (ByVal hwnd As LongPtr, ByVal hPrinter As LongPtr, ByVal pDeviceName As LongPtr, ByVal pDevModeOutput As LongPtr, ByVal pDevModeInput As LongPtr, ByVal fMode As Long) As Long
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
#Else
    Private Declare Function OpenPrinterW Lib "winspool.drv" (ByVal pPrinterName As Long, ByRef phPrinter As Long, ByVal pDefault As Long) As Long
    Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
    Private Declare Function GetPrinterW Lib "winspool.drv" (ByVal hPrinter As Long, ByVal Level As Long, ByVal pPrinter As Long, ByVal cbBuf As Long, ByRef pcbNeeded As Long) As Long
    Private Declare Function SetPrinterW Lib "winspool.drv" (ByVal hPrinter As Long, ByVal Level As Long, ByVal pPrinter As Long, ByVal Command As Long) As Long
    Private Declare Function DocumentPropertiesW Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As Long, ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, ByVal fMode As Long) As Long
    Private Declare Function GetLastError Lib "kernel32" () As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
#End If

#If VBA7 Then
Private Type PrinterInfo9
    pDevMode As LongPtr
End Type
#Else
Private Type PrinterInfo9
    pDevMode As Long
End Type
#End If

Private Type FixedPrinterInfo
    WindowsName As String
    ExcelName As String
    PortName As String
    DriverName As String
    ServerName As String
End Type

Private Type PrinterDevModeSnapshot
    HasPerUserDevMode As Boolean
    DevModeBytes() As Byte
End Type

Private Type FixedPrintProfile
    FormatVersion As String
    PrinterName As String
    ServerName As String
    WindowsName As String
    DriverName As String
    WindowsUserName As String
    ComputerName As String
    RegisteredAt As String
    DevModeSize As Long
    Checksum As String
    DevModeBytes() As Byte
End Type

Public Sub 固定印刷モードを単件に設定する()
    gSealPrintMode = SEAL_PRINT_MODE_SINGLE
End Sub

Public Sub 固定印刷モードを複数に設定する()
    gSealPrintMode = SEAL_PRINT_MODE_MULTI
End Sub

Public Sub 固定印刷モードを未設定に戻す()
    gSealPrintMode = SEAL_PRINT_MODE_UNSET
End Sub

Public Sub 固定印刷設定を登録する()
    Dim currentStage As String
    Dim errNumber As Long
    Dim errDescription As String
    Dim info As FixedPrinterInfo
    Dim devMode() As Byte
    Dim profile As FixedPrintProfile
    Dim profilePath As String
    Dim answer As VbMsgBoxResult
    Dim loaded As FixedPrintProfile
    Dim wasCanceled As Boolean

    On Error GoTo ErrorHandler

    currentStage = "対象プリンター検索"
    info = ResolveFixedPrinter()
    profilePath = GetProfileFilePath()

    If ProfileFileExists() Then
        answer = MsgBox("このPCには既に固定印刷設定が登録されています。" & vbCrLf & _
                        "現在の設定で上書きしますか？", vbQuestion Or vbYesNo Or vbDefaultButton2, "固定印刷設定の上書き確認")
    Else
        answer = MsgBox("現在の『" & FIXED_PRINTER_NAME & "』の印刷設定を、" & vbCrLf & _
                        "このPCの固定印刷設定として登録します。" & vbCrLf & _
                        "よろしいですか？", vbQuestion Or vbYesNo Or vbDefaultButton2, "固定印刷設定の登録確認")
    End If
    If answer <> vbYes Then Exit Sub

    currentStage = "プリンター設定画面"
    MsgBox "固定印刷で使用する設定を指定してください。" & vbCrLf & _
           "設定後に［OK］を押すと、このPCの固定印刷設定として登録されます。", vbInformation
    devMode = PromptUserDevModeBytes(info.WindowsName, wasCanceled)
    If wasCanceled Then
        MsgBox "固定印刷設定の登録をキャンセルしました。既存の設定ファイルは変更していません。", vbInformation
        Exit Sub
    End If
    profile = BuildProfile(info, devMode)

    currentStage = "LOCALAPPDATAフォルダ作成"
    EnsureProfileFolderExists

    currentStage = "設定ファイル保存"
    SaveProfileToLocalAppData profile

    currentStage = "保存ファイル再読込み"
    loaded = LoadProfileFromLocalAppData(False)

    currentStage = "登録情報検証"
    ValidateProfileForCurrentPc loaded, info
    VerifyProfileData loaded

    WriteFixedPrintLog "固定印刷設定登録成功", "成功", "Printer=" & info.WindowsName & "; Computer=" & GetComputerNameText() & "; User=" & GetWindowsUserNameText()
    MsgBox "固定印刷設定を登録しました。" & vbCrLf & vbCrLf & _
           "プリンター名: " & info.WindowsName & vbCrLf & _
           "コンピューター名: " & GetComputerNameText() & vbCrLf & _
           "Windowsユーザー名: " & GetWindowsUserNameText() & vbCrLf & _
           "登録日時: " & profile.RegisteredAt & vbCrLf & _
           "保存先: " & profilePath, vbInformation
    Exit Sub

ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    WriteFixedPrintLog "固定印刷設定登録失敗", "失敗", BuildErrorLogNote(currentStage, errNumber, errDescription)
    MsgBox "固定印刷設定を登録できませんでした。" & vbCrLf & vbCrLf & _
           BuildErrorMessage(currentStage, errNumber, errDescription), vbExclamation
End Sub

Public Sub 固定印刷設定をテストする()
    Dim currentStage As String
    Dim errNumber As Long
    Dim errDescription As String
    Dim info As FixedPrinterInfo
    Dim profile As FixedPrintProfile
    Dim originalSnapshot As PrinterDevModeSnapshot
    Dim restoreDevMode As Boolean
    Dim restoreWarning As String
    Dim msg As String
    Dim appliedDevMode() As Byte

    On Error GoTo ErrorHandler

    currentStage = "対象プリンター検索"
    info = ResolveFixedPrinter()
    currentStage = "現在設定退避"
    originalSnapshot = GetPrinterDevModeSnapshot(info.WindowsName)
    restoreDevMode = True

    currentStage = "設定ファイル読込み"
    profile = LoadProfileFromLocalAppData(True)
    currentStage = "登録情報検証"
    ValidateProfileForCurrentPc profile, info
    currentStage = "固定DEVMODE適用"
    appliedDevMode = ApplyUserDevModeBytes(info.WindowsName, profile.DevModeBytes)
    currentStage = "適用結果確認"
    VerifyAppliedDevMode info.WindowsName, appliedDevMode
    currentStage = "印刷前設定復元"
    If Not TryRestorePrinterSnapshot(info.WindowsName, originalSnapshot, restoreWarning) Then Err.Raise vbObjectError + 5201, , restoreWarning
    restoreDevMode = False
    currentStage = "復元結果確認"
    VerifyRestoredSnapshot info.WindowsName, originalSnapshot

    WriteFixedPrintLog "固定設定テスト成功", "成功", info.WindowsName & "; Computer=" & GetComputerNameText() & "; User=" & GetWindowsUserNameText()
    MsgBox "固定印刷設定の適用と復元を確認しました。実際の印刷は行っていません。", vbInformation
    Exit Sub

ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    msg = BuildErrorMessage(currentStage, errNumber, errDescription)
    If restoreDevMode Then
        If Not TryRestorePrinterSnapshot(info.WindowsName, originalSnapshot, restoreWarning) Then msg = msg & vbCrLf & restoreWarning
    End If
    WriteFixedPrintLog "固定設定テスト失敗", "失敗", BuildErrorLogNote(currentStage, errNumber, errDescription)
    MsgBox "固定印刷設定テストに失敗しました。" & vbCrLf & vbCrLf & msg, vbExclamation
End Sub

Public Sub 固定印刷プリンター診断()
    Dim currentStage As String
    Dim errNumber As Long
    Dim errDescription As String
    Dim ws As Worksheet
    Dim svc As Object, printers As Object, p As Object
    Dim rowNo As Long

    On Error GoTo ErrorHandler

    currentStage = "診断シート作成"
    Set ws = GetOrCreatePrinterDiagnosticSheet()
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Name", "ServerName", "PortName", "DriverName")

    currentStage = "Win32_Printer検索"
    rowNo = 2
    Set svc = GetObject("winmgmts:\\.\root\cimv2")
    Set printers = svc.ExecQuery("SELECT Name, ServerName, PortName, DriverName FROM Win32_Printer")
    For Each p In printers
        If InStr(CStr(p.Name), FIXED_PRINTER_NAME) > 0 Then
            ws.Cells(rowNo, 1).Value = CStr(p.Name)
            ws.Cells(rowNo, 2).Value = CStr(Nz(p.ServerName))
            ws.Cells(rowNo, 3).Value = CStr(Nz(p.PortName))
            ws.Cells(rowNo, 4).Value = CStr(Nz(p.DriverName))
            rowNo = rowNo + 1
        End If
    Next
    ws.Columns("A:D").AutoFit

    If rowNo = 2 Then
        MsgBox "名前に『" & FIXED_PRINTER_NAME & "』を含むプリンターはWin32_Printerに見つかりませんでした。" & vbCrLf & _
               "診断シート: " & ws.Name, vbInformation
    Else
        MsgBox "プリンター診断が完了しました。" & vbCrLf & _
               "診断シート『" & ws.Name & "』に Name / ServerName / PortName / DriverName を出力しました。", vbInformation
    End If
    Exit Sub

ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    WriteFixedPrintLog "固定印刷プリンター診断失敗", "失敗", BuildErrorLogNote(currentStage, errNumber, errDescription)
    MsgBox "固定印刷プリンター診断に失敗しました。" & vbCrLf & vbCrLf & _
           BuildErrorMessage(currentStage, errNumber, errDescription), vbExclamation
End Sub

Public Sub 固定設定で印刷する()
    Dim currentStage As String
    Dim errNumber As Long
    Dim errDescription As String
    Dim oldActivePrinter As String
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean, oldDisplayAlerts As Boolean
    Dim oldPrintCommunication As Boolean, canUsePrintCommunication As Boolean
    Dim info As FixedPrinterInfo
    Dim profile As FixedPrintProfile
    Dim originalSnapshot As PrinterDevModeSnapshot
    Dim restoreDevMode As Boolean, restoreActivePrinter As Boolean
    Dim targetSheet As Worksheet
    Dim mainError As String, restoreWarning As String
    Dim appliedDevMode() As Byte

    On Error GoTo ErrorHandler

    currentStage = "登録ファイル確認"
    If Not ProfileFileExists() Then Err.Raise vbObjectError + 5301, , GetProfileNotRegisteredMessage()
    currentStage = "設定ファイル読込み"
    profile = LoadProfileFromLocalAppData(True)
    currentStage = "対象プリンター検索"
    info = ResolveFixedPrinter()
    currentStage = "登録情報検証"
    ValidateProfileForCurrentPc profile, info
    currentStage = "印刷対象シート判定"
    Set targetSheet = ResolveSealPrintTargetSheet()

    oldActivePrinter = Application.ActivePrinter
    restoreActivePrinter = (Len(oldActivePrinter) > 0)
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    On Error Resume Next
    oldPrintCommunication = Application.PrintCommunication
    canUsePrintCommunication = (Err.Number = 0)
    Err.Clear
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    If canUsePrintCommunication Then Application.PrintCommunication = False

    currentStage = "現在DEVMODE退避"
    originalSnapshot = GetPrinterDevModeSnapshot(info.WindowsName)
    restoreDevMode = True
    currentStage = "固定DEVMODE適用"
    appliedDevMode = ApplyUserDevModeBytes(info.WindowsName, profile.DevModeBytes)
    currentStage = "適用結果確認"
    VerifyAppliedDevMode info.WindowsName, appliedDevMode

    If canUsePrintCommunication Then Application.PrintCommunication = True
    currentStage = "ActivePrinter切替"
    Application.ActivePrinter = info.ExcelName
    currentStage = "PrintOut実行"
    targetSheet.PrintOut

    WriteFixedPrintLog "固定印刷成功", "成功", "Sheet=" & targetSheet.Name & "; Printer=" & info.WindowsName & "; Computer=" & GetComputerNameText() & "; User=" & GetWindowsUserNameText()

Cleanup:
    If restoreDevMode Then
        If Not TryRestorePrinterSnapshot(info.WindowsName, originalSnapshot, restoreWarning) Then
            WriteFixedPrintLog "プリンター設定復元失敗", "失敗", restoreWarning
        End If
    End If
    If restoreActivePrinter Then
        On Error Resume Next
        Application.ActivePrinter = oldActivePrinter
        If Err.Number <> 0 Then
            restoreWarning = AppendWarning(restoreWarning, "Application.ActivePrinterの復元に失敗しました: Err " & CStr(Err.Number) & " " & Err.Description)
            Err.Clear
        End If
        On Error GoTo 0
    End If
    On Error Resume Next
    If canUsePrintCommunication Then Application.PrintCommunication = oldPrintCommunication
    Application.CutCopyMode = False
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0

    If Len(restoreWarning) > 0 Then
        MsgBox "印刷前のプリンター設定へ完全に戻せなかった可能性があります。" & vbCrLf & _
               "通常印刷を行う前にプリンター設定を確認してください。" & vbCrLf & vbCrLf & restoreWarning, vbExclamation
    End If
    If Len(mainError) > 0 Then
        WriteFixedPrintLog "固定印刷失敗", "失敗", mainError
        MsgBox "固定設定で印刷できませんでした。誤った設定での印刷を防ぐため中止しました。" & vbCrLf & mainError, vbExclamation
    End If
    Exit Sub

ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    mainError = BuildErrorMessage(currentStage, errNumber, errDescription)
    Resume Cleanup
End Sub

Private Function ResolveSealPrintTargetSheet() As Worksheet
    Dim wsSingle As Worksheet, wsMulti As Worksheet

    Set wsSingle = ThisWorkbook.Worksheets(SINGLE_SEAL_SHEET)
    Set wsMulti = ThisWorkbook.Worksheets(MULTI_SEAL_SHEET)

    Select Case gSealPrintMode
        Case SEAL_PRINT_MODE_SINGLE
            If Not HasSingleSealOutputData(wsSingle) Then
                Err.Raise vbObjectError + 5302, , "直前に反映した単件用シートに印刷データがありません。" & vbCrLf & _
                          "先に単件のシール内容を反映してください。"
            End If
            Set ResolveSealPrintTargetSheet = wsSingle

        Case SEAL_PRINT_MODE_MULTI
            If Not HasMultiSealOutputData(wsMulti) Then
                Err.Raise vbObjectError + 5302, , "直前に反映した複数用シートに印刷データがありません。" & vbCrLf & _
                          "先に複数のシール内容を反映してください。"
            End If
            Set ResolveSealPrintTargetSheet = wsMulti

        Case Else
            Err.Raise vbObjectError + 5303, , "印刷対象を判定できません。" & vbCrLf & _
                      "先に単件または複数のシール内容を反映してください。"
    End Select
End Function

Private Function HasSingleSealOutputData(ByVal ws As Worksheet) As Boolean
    Dim slot As Long
    For slot = 1 To SEAL_PER_PAGE_LOCAL
        If HasSealSlotOutputData(ws, slot, 0) Then
            HasSingleSealOutputData = True
            Exit Function
        End If
    Next slot
End Function

Private Function HasMultiSealOutputData(ByVal ws As Worksheet) As Boolean
    Dim pageCount As Long, pageNo As Long, slot As Long
    pageCount = GetMultiSealPageCountFromPrintArea(ws)
    For pageNo = 1 To pageCount
        For slot = 1 To SEAL_PER_PAGE_LOCAL
            If HasSealSlotOutputData(ws, slot, (pageNo - 1) * SEAL_PAGE_ROWS_LOCAL) Then
                HasMultiSealOutputData = True
                Exit Function
            End If
        Next slot
    Next pageNo
End Function

Private Function HasSealSlotOutputData(ByVal ws As Worksheet, ByVal slot As Long, ByVal rowOffset As Long) As Boolean
    Dim baseCol As Long, baseRow As Long, offSlash As Long, offG1 As Long
    GroupBase slot, baseCol, baseRow, offSlash, offG1
    baseRow = baseRow + rowOffset

    HasSealSlotOutputData = _
        Len(GetMergeTopLeftText(ws, baseRow, baseCol)) > 0 Or _
        Len(GetMergeTopLeftText(ws, baseRow, baseCol + 1)) > 0 Or _
        Len(GetMergeTopLeftText(ws, baseRow, baseCol + offSlash)) > 0 Or _
        Len(GetMergeTopLeftText(ws, baseRow, baseCol + 7)) > 0 Or _
        Len(GetMergeTopLeftText(ws, baseRow + 1, baseCol)) > 0 Or _
        Len(GetMergeTopLeftText(ws, baseRow + 1, baseCol + 1)) > 0
End Function

Private Function GetMergeTopLeftText(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal colNo As Long) As String
    With ws.Cells(rowNo, colNo)
        If .MergeCells Then
            GetMergeTopLeftText = Trim$(CStr(.MergeArea.Cells(1, 1).Value))
        Else
            GetMergeTopLeftText = Trim$(CStr(.Value))
        End If
    End With
End Function

Private Function GetMultiSealPageCountFromPrintArea(ByVal ws As Worksheet) As Long
    Dim areaText As String, rng As Range, bottomRow As Long
    On Error GoTo Fallback
    areaText = Trim$(ws.PageSetup.PrintArea)
    If Len(areaText) > 0 Then
        Set rng = ws.Range(areaText)
        bottomRow = rng.Row + rng.Rows.Count - 1
        GetMultiSealPageCountFromPrintArea = (bottomRow + SEAL_PAGE_ROWS_LOCAL - 1) \ SEAL_PAGE_ROWS_LOCAL
        If GetMultiSealPageCountFromPrintArea < 1 Then GetMultiSealPageCountFromPrintArea = 1
        Exit Function
    End If
Fallback:
    GetMultiSealPageCountFromPrintArea = 1
End Function

Private Function GetOrCreatePrinterDiagnosticSheet() As Worksheet
    Const DIAGNOSTIC_SHEET_NAME As String = "固定印刷プリンター診断"

    On Error Resume Next
    Set GetOrCreatePrinterDiagnosticSheet = ThisWorkbook.Worksheets(DIAGNOSTIC_SHEET_NAME)
    On Error GoTo 0
    If GetOrCreatePrinterDiagnosticSheet Is Nothing Then
        Set GetOrCreatePrinterDiagnosticSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreatePrinterDiagnosticSheet.Name = DIAGNOSTIC_SHEET_NAME
    End If
End Function

Private Function ResolveFixedPrinter() As FixedPrinterInfo
    Dim svc As Object, printers As Object, p As Object
    Dim candidates() As FixedPrinterInfo
    Dim uniqueCandidates() As FixedPrinterInfo
    Dim candidateCount As Long
    Dim uniqueCount As Long
    Dim i As Long, existingIndex As Long
    Dim info As FixedPrinterInfo
    Dim selectedInfo As FixedPrinterInfo

    Set svc = GetObject("winmgmts:\\.\root\cimv2")
    Set printers = svc.ExecQuery("SELECT Name, ServerName, PortName, DriverName FROM Win32_Printer")

    For Each p In printers
        If IsTargetPrinter(CStr(p.Name), CStr(Nz(p.ServerName)), CStr(Nz(p.PortName))) Then
            info.WindowsName = CStr(p.Name)
            info.PortName = CStr(Nz(p.PortName))
            info.DriverName = CStr(Nz(p.DriverName))
            info.ServerName = CStr(Nz(p.ServerName))
            info.ExcelName = BuildExcelActivePrinterName(info.WindowsName, info.PortName)
            candidateCount = candidateCount + 1
            ReDim Preserve candidates(1 To candidateCount)
            candidates(candidateCount) = info
        End If
    Next

    If candidateCount = 0 Then Err.Raise vbObjectError + 5101, , "対象プリンター（土木課C4476R / PRTSV03.sojanet.local）がWindowsに登録されていません。"

    For i = 1 To candidateCount
        existingIndex = FindSamePhysicalPrinterIndex(uniqueCandidates, uniqueCount, candidates(i))
        If existingIndex = 0 Then
            uniqueCount = uniqueCount + 1
            ReDim Preserve uniqueCandidates(1 To uniqueCount)
            uniqueCandidates(uniqueCount) = candidates(i)
        Else
            uniqueCandidates(existingIndex) = ChoosePreferredPrinterCandidate(uniqueCandidates(existingIndex), candidates(i))
        End If
    Next i

    If uniqueCount = 1 Then
        selectedInfo = uniqueCandidates(1)
        selectedInfo.ExcelName = ResolveExcelActivePrinterName(selectedInfo)
        WriteFixedPrintLog "固定印刷プリンター選択", "成功", _
            "BeforeDedup=" & CStr(candidateCount) & "; AfterDedup=" & CStr(uniqueCount) & _
            "; Selected=" & selectedInfo.WindowsName & "; WindowsName=" & selectedInfo.WindowsName & _
            "; PortName=" & selectedInfo.PortName & "; DriverName=" & selectedInfo.DriverName & _
            "; ResolvedActivePrinter=" & selectedInfo.ExcelName
        ResolveFixedPrinter = selectedInfo
        Exit Function
    End If

    WriteFixedPrintLog "固定印刷プリンター選択", "失敗", _
        "BeforeDedup=" & CStr(candidateCount) & "; AfterDedup=" & CStr(uniqueCount) & "; Selected=なし"
    Err.Raise vbObjectError + 5102, , "対象プリンター候補が複数あり一意に決められないため印刷しません。" & BuildPrinterCandidateSummary(uniqueCandidates, uniqueCount)
End Function

Private Function FindSamePhysicalPrinterIndex(ByRef candidates() As FixedPrinterInfo, ByVal candidateCount As Long, ByRef target As FixedPrinterInfo) As Long
    Dim i As Long
    If candidateCount <= 0 Then Exit Function

    For i = 1 To candidateCount
        If IsSamePhysicalPrinter(candidates(i), target) Then
            FindSamePhysicalPrinterIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function IsSamePhysicalPrinter(ByRef leftInfo As FixedPrinterInfo, ByRef rightInfo As FixedPrinterInfo) As Boolean
    IsSamePhysicalPrinter = _
        (StrComp(GetPrinterQueueName(leftInfo.WindowsName), GetPrinterQueueName(rightInfo.WindowsName), vbTextCompare) = 0) And _
        (StrComp(leftInfo.PortName, rightInfo.PortName, vbTextCompare) = 0) And _
        (StrComp(leftInfo.DriverName, rightInfo.DriverName, vbTextCompare) = 0) And _
        (StrComp(GetPrinterIdentityServerName(leftInfo), GetPrinterIdentityServerName(rightInfo), vbTextCompare) = 0)
End Function

Private Function ChoosePreferredPrinterCandidate(ByRef currentInfo As FixedPrinterInfo, ByRef newInfo As FixedPrinterInfo) As FixedPrinterInfo
    Dim activePrinterText As String

    On Error Resume Next
    activePrinterText = Application.ActivePrinter
    On Error GoTo 0

    If Len(activePrinterText) > 0 Then
        If InStr(1, activePrinterText, currentInfo.WindowsName, vbTextCompare) > 0 Then
            ChoosePreferredPrinterCandidate = currentInfo
            Exit Function
        End If
        If InStr(1, activePrinterText, newInfo.WindowsName, vbTextCompare) > 0 Then
            ChoosePreferredPrinterCandidate = newInfo
            Exit Function
        End If
    End If

    If CanSetExcelActivePrinter(newInfo) And Not CanSetExcelActivePrinter(currentInfo) Then
        ChoosePreferredPrinterCandidate = newInfo
    Else
        ChoosePreferredPrinterCandidate = currentInfo
    End If
End Function

Private Function CanSetExcelActivePrinter(ByRef info As FixedPrinterInfo) As Boolean
    Dim oldActivePrinter As String

    On Error GoTo Failed
    oldActivePrinter = Application.ActivePrinter
    CanSetExcelActivePrinter = (Len(ResolveExcelActivePrinterName(info, False)) > 0)

Restore:
    On Error Resume Next
    If Len(oldActivePrinter) > 0 Then Application.ActivePrinter = oldActivePrinter
    Err.Clear
    On Error GoTo 0
    Exit Function

Failed:
    CanSetExcelActivePrinter = False
    Resume Restore
End Function

Private Function ResolveExcelActivePrinterName(ByRef info As FixedPrinterInfo, Optional ByVal writeLog As Boolean = True) As String
    Dim oldActivePrinter As String
    Dim candidates() As String
    Dim candidateCount As Long
    Dim i As Long, neNo As Long
    Dim candidate As String
    Dim resolvedName As String
    Dim errorText As String
    Dim attemptLog As String
    Dim currentPort As String
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim restoreError As String
    Dim excelPrinterName As String

    On Error GoTo ResolveFailed

    oldActivePrinter = Application.ActivePrinter
    If Len(Trim$(oldActivePrinter)) = 0 Then
        Err.Raise vbObjectError + 5107, , "探索前のApplication.ActivePrinterを取得できなかったため、候補探索を中止しました。"
    End If

    If IsExcelActivePrinterForInfo(oldActivePrinter, info) Then
        AddExcelActivePrinterCandidate candidates, candidateCount, oldActivePrinter
        currentPort = ExtractExcelActivePrinterPort(oldActivePrinter)
        AddExcelActivePrinterPortCandidates candidates, candidateCount, info.WindowsName, currentPort
    End If

    AddExcelActivePrinterPortCandidates candidates, candidateCount, info.WindowsName, info.PortName
    excelPrinterName = GetPrinterQueueName(info.WindowsName)

    ' Excelのネットワークポート表現はWMIのPortNameと一致しない場合があるため、
    ' Ne番号を固定せず、一定範囲の候補を実際にApplication.ActivePrinterへ設定して確認する。
    For neNo = 0 To 99
        AddExcelActivePrinterCandidate candidates, candidateCount, _
            BuildExcelActivePrinterName(info.WindowsName, "Ne" & Format$(neNo, "00") & ":")
        AddExcelActivePrinterCandidate candidates, candidateCount, _
            BuildExcelActivePrinterName(info.WindowsName, "Ne" & Format$(neNo, "00"))
        If StrComp(excelPrinterName, info.WindowsName, vbTextCompare) <> 0 Then
            AddExcelActivePrinterCandidate candidates, candidateCount, _
                BuildExcelActivePrinterName(excelPrinterName, "Ne" & Format$(neNo, "00") & ":")
            AddExcelActivePrinterCandidate candidates, candidateCount, _
                BuildExcelActivePrinterName(excelPrinterName, "Ne" & Format$(neNo, "00"))
        End If
    Next neNo

    For i = 1 To candidateCount
        candidate = candidates(i)
        resolvedName = ""
        errorText = ""
        If TrySetExcelActivePrinterCandidate(candidate, info, resolvedName, errorText) Then
            attemptLog = AppendActivePrinterAttempt(attemptLog, candidate, "成功: " & resolvedName)
            GoTo ResolveSucceeded
        End If
        attemptLog = AppendActivePrinterAttempt(attemptLog, candidate, errorText)
    Next i

    failureNumber = vbObjectError + 5108
    failureDescription = "対象プリンターはWindows上に見つかりましたが、Excelで使用できるプリンター名（ActivePrinter）を特定できませんでした。" & vbCrLf & _
                         "対象プリンター以外へ切り替えないため、印刷を中止します。" & vbCrLf & _
                         BuildActivePrinterResolutionLog(info, oldActivePrinter, "", attemptLog)
    GoTo ResolveCleanup

ResolveSucceeded:
    info.ExcelName = resolvedName
    ResolveExcelActivePrinterName = resolvedName
    If writeLog Then
        WriteFixedPrintLog "固定印刷ActivePrinter解決", "成功", _
            BuildActivePrinterResolutionLog(info, oldActivePrinter, resolvedName, attemptLog)
    End If
    GoTo ResolveCleanup

ResolveFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description

ResolveCleanup:
    On Error Resume Next
    If Len(oldActivePrinter) > 0 Then
        Application.ActivePrinter = oldActivePrinter
        If Err.Number <> 0 Then
            restoreError = "探索前のActivePrinterへ復元できませんでした: Err " & CStr(Err.Number) & " " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo 0

    If Len(restoreError) > 0 Then
        If Len(failureDescription) = 0 Then
            failureNumber = vbObjectError + 5109
            failureDescription = restoreError
        Else
            failureDescription = failureDescription & vbCrLf & restoreError
        End If
    End If

    If Len(failureDescription) > 0 Then
        If writeLog Then
            WriteFixedPrintLog "固定印刷ActivePrinter解決", "失敗", _
                BuildActivePrinterResolutionLog(info, oldActivePrinter, "", attemptLog) & "; Error=" & failureDescription
        End If
        Err.Raise failureNumber, , failureDescription
    End If
End Function

Private Sub AddExcelActivePrinterCandidate(ByRef candidates() As String, ByRef candidateCount As Long, ByVal candidateText As String)
    Dim i As Long

    candidateText = Trim$(candidateText)
    If Len(candidateText) = 0 Then Exit Sub

    For i = 1 To candidateCount
        If StrComp(candidates(i), candidateText, vbTextCompare) = 0 Then Exit Sub
    Next i

    candidateCount = candidateCount + 1
    If candidateCount = 1 Then
        ReDim candidates(1 To 1)
    Else
        ReDim Preserve candidates(1 To candidateCount)
    End If
    candidates(candidateCount) = candidateText
End Sub

Private Sub AddExcelActivePrinterPortCandidates(ByRef candidates() As String, ByRef candidateCount As Long, ByVal printerName As String, ByVal portName As String)
    Dim normalizedPort As String
    Dim queueName As String

    normalizedPort = Trim$(portName)
    If Len(normalizedPort) = 0 Then Exit Sub

    AddExcelActivePrinterCandidate candidates, candidateCount, BuildExcelActivePrinterName(printerName, normalizedPort)
    If Right$(normalizedPort, 1) = ":" Then
        AddExcelActivePrinterCandidate candidates, candidateCount, _
            BuildExcelActivePrinterName(printerName, Left$(normalizedPort, Len(normalizedPort) - 1))
    Else
        AddExcelActivePrinterCandidate candidates, candidateCount, _
            BuildExcelActivePrinterName(printerName, normalizedPort & ":")
    End If

    queueName = GetPrinterQueueName(printerName)
    If StrComp(queueName, printerName, vbTextCompare) <> 0 Then
        AddExcelActivePrinterCandidate candidates, candidateCount, BuildExcelActivePrinterName(queueName, normalizedPort)
        If Right$(normalizedPort, 1) = ":" Then
            AddExcelActivePrinterCandidate candidates, candidateCount, _
                BuildExcelActivePrinterName(queueName, Left$(normalizedPort, Len(normalizedPort) - 1))
        Else
            AddExcelActivePrinterCandidate candidates, candidateCount, _
                BuildExcelActivePrinterName(queueName, normalizedPort & ":")
        End If
    End If
End Sub

Private Function ExtractExcelActivePrinterPort(ByVal activePrinterText As String) As String
    Dim markerPos As Long

    markerPos = InStrRev(activePrinterText, " on ", -1, vbTextCompare)
    If markerPos > 0 Then ExtractExcelActivePrinterPort = Trim$(Mid$(activePrinterText, markerPos + 4))
End Function

Private Function TrySetExcelActivePrinterCandidate(ByVal candidate As String, ByRef info As FixedPrinterInfo, ByRef resolvedName As String, ByRef errorText As String) As Boolean
    Dim errNumber As Long
    Dim errDescription As String
    Dim activePrinterText As String

    On Error Resume Next
    Err.Clear
    Application.ActivePrinter = candidate
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear

    If errNumber = 0 Then
        activePrinterText = Application.ActivePrinter
        errNumber = Err.Number
        errDescription = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If errNumber <> 0 Then
        errorText = "Err " & CStr(errNumber) & " " & errDescription
        Exit Function
    End If

    If Not IsExcelActivePrinterForInfo(activePrinterText, info) Then
        errorText = "設定後のActivePrinterが対象プリンターではありません: " & activePrinterText
        Exit Function
    End If

    resolvedName = activePrinterText
    TrySetExcelActivePrinterCandidate = True
End Function

Private Function IsExcelActivePrinterForInfo(ByVal activePrinterText As String, ByRef info As FixedPrinterInfo) As Boolean
    Dim markerPos As Long
    Dim activeQueueName As String
    Dim targetQueueName As String

    markerPos = InStrRev(activePrinterText, " on ", -1, vbTextCompare)
    If markerPos > 0 Then
        activeQueueName = Left$(activePrinterText, markerPos - 1)
    Else
        activeQueueName = activePrinterText
    End If

    activeQueueName = GetPrinterQueueName(activeQueueName)
    targetQueueName = GetPrinterQueueName(info.WindowsName)
    IsExcelActivePrinterForInfo = _
        (Len(activeQueueName) > 0) And _
        (StrComp(activeQueueName, targetQueueName, vbTextCompare) = 0)
End Function

Private Function AppendActivePrinterAttempt(ByVal currentText As String, ByVal candidate As String, ByVal resultText As String) As String
    If Len(currentText) > 0 Then currentText = currentText & vbCrLf
    AppendActivePrinterAttempt = currentText & candidate & " => " & resultText
End Function

Private Function BuildActivePrinterResolutionLog(ByRef info As FixedPrinterInfo, ByVal oldActivePrinter As String, ByVal resolvedName As String, ByVal attemptLog As String) As String
    If Len(resolvedName) = 0 Then resolvedName = "（なし）"
    If Len(attemptLog) = 0 Then attemptLog = "（候補なし）"

    BuildActivePrinterResolutionLog = _
        "WindowsName=" & info.WindowsName & "; PortName=" & info.PortName & _
        "; DriverName=" & info.DriverName & "; OldActivePrinter=" & oldActivePrinter & _
        "; ResolvedActivePrinter=" & resolvedName & "; TriedActivePrinter=" & attemptLog
End Function

Private Function BuildPrinterCandidateSummary(ByRef candidates() As FixedPrinterInfo, ByVal candidateCount As Long) As String
    Dim i As Long
    For i = 1 To candidateCount
        BuildPrinterCandidateSummary = BuildPrinterCandidateSummary & vbCrLf & _
            "- " & candidates(i).WindowsName & " / Server=" & candidates(i).ServerName & _
            " / Port=" & candidates(i).PortName & " / Driver=" & candidates(i).DriverName
    Next i
End Function

Private Function IsTargetPrinter(ByVal printerName As String, ByVal serverName As String, ByVal portName As String) As Boolean
    If InStr(printerName, FIXED_PRINTER_NAME) = 0 Then Exit Function

    IsTargetPrinter = _
        IsSamePrinterServer(printerName, FIXED_PRINTER_SERVER) Or _
        IsSamePrinterServer(serverName, FIXED_PRINTER_SERVER) Or _
        IsSamePrinterServer(portName, FIXED_PRINTER_SERVER)
End Function

Private Function IsSamePrinterServer(ByVal candidateText As String, ByVal expectedServer As String) As Boolean
    Dim candidate As String
    Dim expectedShort As String
    Dim expectedFqdn As String

    candidate = NormalizePrinterServerText(candidateText)
    expectedFqdn = NormalizePrinterServerText(expectedServer)
    expectedShort = GetShortServerName(expectedFqdn)

    IsSamePrinterServer = (InStr(candidate, expectedFqdn) > 0 Or InStr(candidate, expectedShort) > 0)
End Function

Private Function NormalizePrinterServerText(ByVal value As String) As String
    Dim normalized As String
    normalized = LCase$(Trim$(CStr(value)))
    normalized = Replace(normalized, "\", " ")
    normalized = Replace(normalized, "/", " ")
    normalized = Replace(normalized, "(", " ")
    normalized = Replace(normalized, ")", " ")
    normalized = Replace(normalized, "　", " ")
    Do While InStr(normalized, "  ") > 0
        normalized = Replace(normalized, "  ", " ")
    Loop
    NormalizePrinterServerText = normalized
End Function

Private Function GetShortServerName(ByVal normalizedServer As String) As String
    Dim dotPos As Long
    normalizedServer = Trim$(normalizedServer)
    dotPos = InStr(1, normalizedServer, ".", vbTextCompare)
    If dotPos > 1 Then
        GetShortServerName = Left$(normalizedServer, dotPos - 1)
    Else
        GetShortServerName = normalizedServer
    End If
End Function

Private Function GetPrinterQueueName(ByVal printerName As String) As String
    Dim normalized As String
    Dim slashPos As Long
    normalized = Replace(Trim$(printerName), "/", "\")
    Do While Left$(normalized, 1) = "\"
        normalized = Mid$(normalized, 2)
    Loop
    slashPos = InStrRev(normalized, "\")
    If slashPos > 0 Then
        GetPrinterQueueName = Mid$(normalized, slashPos + 1)
    ElseIf InStr(1, normalized, FIXED_PRINTER_NAME, vbTextCompare) > 0 Then
        GetPrinterQueueName = FIXED_PRINTER_NAME
    Else
        GetPrinterQueueName = normalized
    End If
End Function

Private Function GetPrinterIdentityServerName(ByRef info As FixedPrinterInfo) As String
    If Len(Trim$(info.ServerName)) > 0 Then
        GetPrinterIdentityServerName = NormalizePrinterServerForIdentity(info.ServerName)
    Else
        GetPrinterIdentityServerName = NormalizePrinterServerForIdentity(ExtractServerNameFromPrinterName(info.WindowsName))
    End If
End Function

Private Function ExtractServerNameFromPrinterName(ByVal printerName As String) As String
    Dim normalized As String
    Dim slashPos As Long
    normalized = Replace(Trim$(printerName), "/", "\")
    Do While Left$(normalized, 1) = "\"
        normalized = Mid$(normalized, 2)
    Loop
    slashPos = InStr(1, normalized, "\", vbTextCompare)
    If slashPos > 1 Then
        ExtractServerNameFromPrinterName = Left$(normalized, slashPos - 1)
    Else
        ExtractServerNameFromPrinterName = normalized
    End If
End Function

Private Function NormalizePrinterServerForIdentity(ByVal serverName As String) As String
    Dim normalized As String
    normalized = LCase$(Trim$(CStr(serverName)))
    normalized = Replace(normalized, "/", "\")
    Do While Left$(normalized, 1) = "\"
        normalized = Mid$(normalized, 2)
    Loop
    Do While Right$(normalized, 1) = "\"
        normalized = Left$(normalized, Len(normalized) - 1)
    Loop
    If Right$(normalized, Len(".sojanet.local")) = ".sojanet.local" Then
        normalized = Left$(normalized, Len(normalized) - Len(".sojanet.local"))
    End If
    NormalizePrinterServerForIdentity = normalized
End Function

Private Function BuildExcelActivePrinterName(ByVal printerName As String, ByVal portName As String) As String
    BuildExcelActivePrinterName = Trim$(printerName) & " on " & Trim$(portName)
End Function

Private Function BuildProfile(ByRef info As FixedPrinterInfo, ByRef devModeBytes() As Byte) As FixedPrintProfile
    Dim profile As FixedPrintProfile
    profile.FormatVersion = PROFILE_FORMAT_VERSION
    profile.PrinterName = FIXED_PRINTER_NAME
    profile.ServerName = FIXED_PRINTER_SERVER
    profile.WindowsName = info.WindowsName
    profile.DriverName = info.DriverName
    profile.WindowsUserName = GetWindowsUserNameText()
    profile.ComputerName = GetComputerNameText()
    profile.RegisteredAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    profile.DevModeSize = ByteArrayLength(devModeBytes)
    profile.Checksum = CalculateByteChecksum(devModeBytes)
    profile.DevModeBytes = devModeBytes
    BuildProfile = profile
End Function

Private Sub SaveProfileToLocalAppData(ByRef profile As FixedPrintProfile)
    Dim fso As Object, ts As Object, profileText As String
    profileText = BuildProfileFileText(profile)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(GetProfileFilePath(), True, True)
    ts.Write profileText
    ts.Close
End Sub

Private Function LoadProfileFromLocalAppData(ByVal showNotRegisteredGuide As Boolean) As FixedPrintProfile
    Dim fso As Object, ts As Object, text As String
    Dim profile As FixedPrintProfile
    If Not ProfileFileExists() Then
        If showNotRegisteredGuide Then
            Err.Raise vbObjectError + 5107, , GetProfileNotRegisteredMessage()
        Else
            Err.Raise vbObjectError + 5107, , "固定印刷プロファイルが未登録です。"
        End If
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(GetProfileFilePath(), 1, False, -1)
    text = ts.ReadAll
    ts.Close
    profile = ParseProfileFileText(text)
    VerifyProfileData profile
    LoadProfileFromLocalAppData = profile
End Function

Private Function BuildProfileFileText(ByRef profile As FixedPrintProfile) As String
    BuildProfileFileText = _
        "FormatVersion=" & EscapeProfileValue(profile.FormatVersion) & PROFILE_LINE_SEPARATOR & _
        "PrinterName=" & EscapeProfileValue(profile.PrinterName) & PROFILE_LINE_SEPARATOR & _
        "ServerName=" & EscapeProfileValue(profile.ServerName) & PROFILE_LINE_SEPARATOR & _
        "WindowsName=" & EscapeProfileValue(profile.WindowsName) & PROFILE_LINE_SEPARATOR & _
        "DriverName=" & EscapeProfileValue(profile.DriverName) & PROFILE_LINE_SEPARATOR & _
        "WindowsUserName=" & EscapeProfileValue(profile.WindowsUserName) & PROFILE_LINE_SEPARATOR & _
        "ComputerName=" & EscapeProfileValue(profile.ComputerName) & PROFILE_LINE_SEPARATOR & _
        "RegisteredAt=" & EscapeProfileValue(profile.RegisteredAt) & PROFILE_LINE_SEPARATOR & _
        "DevModeSize=" & CStr(profile.DevModeSize) & PROFILE_LINE_SEPARATOR & _
        "Checksum=" & profile.Checksum & PROFILE_LINE_SEPARATOR & _
        "DevModeBase64=" & BytesToBase64(profile.DevModeBytes) & PROFILE_LINE_SEPARATOR
End Function

Private Function ParseProfileFileText(ByVal text As String) As FixedPrintProfile
    Dim profile As FixedPrintProfile
    Dim lines As Variant, i As Long, lineText As String, pos As Long, key As String, value As String
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = CStr(lines(i))
        If Len(lineText) > 0 Then
            pos = InStr(1, lineText, PROFILE_KEY_VALUE_SEPARATOR, vbBinaryCompare)
            If pos > 0 Then
                key = Left$(lineText, pos - 1)
                value = UnescapeProfileValue(Mid$(lineText, pos + 1))
                Select Case key
                    Case "FormatVersion": profile.FormatVersion = value
                    Case "PrinterName": profile.PrinterName = value
                    Case "ServerName": profile.ServerName = value
                    Case "WindowsName": profile.WindowsName = value
                    Case "DriverName": profile.DriverName = value
                    Case "WindowsUserName": profile.WindowsUserName = value
                    Case "ComputerName": profile.ComputerName = value
                    Case "RegisteredAt": profile.RegisteredAt = value
                    Case "DevModeSize": profile.DevModeSize = CLng(Val(value))
                    Case "Checksum": profile.Checksum = value
                    Case "DevModeBase64": profile.DevModeBytes = Base64ToBytes(value)
                End Select
            End If
        End If
    Next i
    ParseProfileFileText = profile
End Function

Private Sub ValidateProfileForCurrentPc(ByRef profile As FixedPrintProfile, ByRef info As FixedPrinterInfo)
    VerifyProfileData profile
    If profile.FormatVersion <> PROFILE_FORMAT_VERSION Then Err.Raise vbObjectError + 5110, , "固定印刷設定ファイルの形式バージョンが対応外です。再登録してください。"
    If profile.PrinterName <> FIXED_PRINTER_NAME Or profile.ServerName <> FIXED_PRINTER_SERVER Then Err.Raise vbObjectError + 5111, , "固定印刷設定ファイルのプリンター情報が対象プリンターと一致しません。再登録してください。"
    If profile.ComputerName <> GetComputerNameText() Then Err.Raise vbObjectError + 5112, , "固定印刷設定は別PCで登録されたものです。現在のPCで再登録してください。登録PC=" & profile.ComputerName & " / 現在PC=" & GetComputerNameText()
    If profile.WindowsUserName <> GetWindowsUserNameText() Then Err.Raise vbObjectError + 5113, , "固定印刷設定は別Windowsユーザーで登録されたものです。現在のユーザーで再登録してください。登録ユーザー=" & profile.WindowsUserName & " / 現在ユーザー=" & GetWindowsUserNameText()
    If profile.DriverName <> info.DriverName Then Err.Raise vbObjectError + 5114, , "登録時と現在のプリンタードライバーが異なります。登録時=" & profile.DriverName & " / 現在=" & info.DriverName & "。固定印刷設定を再登録してください。"
End Sub

Private Sub VerifyProfileData(ByRef profile As FixedPrintProfile)
    If IsEmptyByteArray(profile.DevModeBytes) Then Err.Raise vbObjectError + 5115, , "固定印刷設定ファイルにDEVMODEデータがありません。再登録してください。"
    If profile.DevModeSize <> ByteArrayLength(profile.DevModeBytes) Then Err.Raise vbObjectError + 5116, , "固定印刷設定ファイルのDEVMODEサイズが一致しません。ファイル破損の可能性があります。再登録してください。"
    If profile.Checksum <> CalculateByteChecksum(profile.DevModeBytes) Then Err.Raise vbObjectError + 5117, , "固定印刷設定ファイルのチェック値が一致しません。ファイル破損の可能性があります。再登録してください。"
End Sub

Private Function GetProfileFolderPath() As String
    GetProfileFolderPath = Environ$("LOCALAPPDATA") & "\" & PROFILE_BASE_FOLDER
End Function

Private Function GetProfileFilePath() As String
    GetProfileFilePath = GetProfileFolderPath() & "\" & PROFILE_FILE_NAME
End Function

Private Function ProfileFileExists() As Boolean
    ProfileFileExists = (Len(Dir$(GetProfileFilePath(), vbNormal)) > 0)
End Function

Private Sub EnsureProfileFolderExists()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Environ$("LOCALAPPDATA") & "\file-index-sync") Then fso.CreateFolder Environ$("LOCALAPPDATA") & "\file-index-sync"
    If Not fso.FolderExists(GetProfileFolderPath()) Then fso.CreateFolder GetProfileFolderPath()
End Sub

Private Function GetProfileNotRegisteredMessage() As String
    GetProfileNotRegisteredMessage = "このPCでは固定印刷設定が登録されていません。" & vbCrLf & vbCrLf & _
        "1. Windowsで『" & FIXED_PRINTER_NAME & "』を所定の設定にしてください。" & vbCrLf & _
        "2. 『このPCの印刷設定を登録』を押してください。" & vbCrLf & _
        "3. 登録後、もう一度『固定設定で印刷』を実行してください。"
End Function

Private Function EscapeProfileValue(ByVal value As String) As String
    EscapeProfileValue = Replace(Replace(value, "%", "%25"), vbLf, "%0A")
End Function

Private Function UnescapeProfileValue(ByVal value As String) As String
    UnescapeProfileValue = Replace(Replace(value, "%0A", vbLf), "%25", "%")
End Function

Private Function PromptUserDevModeBytes(ByVal printerName As String, ByRef wasCanceled As Boolean) As Byte()
#If VBA7 Then
    Dim hPrinter As LongPtr, pDevMode As LongPtr
#Else
    Dim hPrinter As Long, pDevMode As Long
#End If
    Dim size As Long, result As Long, data() As Byte

    wasCanceled = False
    If OpenPrinterW(StrPtr(printerName), hPrinter, 0) = 0 Then RaiseApiError "OpenPrinterW"
    On Error GoTo CleanFail

    size = DocumentPropertiesW(0, hPrinter, StrPtr(printerName), 0, 0, 0)
    If size <= 0 Then RaiseApiError "DocumentPropertiesW(size)"

    pDevMode = GlobalAlloc(0, size)
    If pDevMode = 0 Then Err.Raise vbObjectError + 5105, , "GlobalAlloc(DEVMODE)に失敗しました。"

    result = DocumentPropertiesW(0, hPrinter, StrPtr(printerName), pDevMode, 0, DM_OUT_BUFFER Or DM_IN_PROMPT)
    If result = IDCANCEL Then
        wasCanceled = True
        GoTo CleanExit
    End If
    If result <> IDOK Then
        If result < 0 Then RaiseApiError "DocumentPropertiesW(prompt)"
        Err.Raise vbObjectError + 5120, , "DocumentPropertiesW(prompt)が想定外の戻り値を返しました。Result=" & CStr(result)
    End If

    ReDim data(0 To size - 1) As Byte
    CopyMemory VarPtr(data(0)), pDevMode, size
    PromptUserDevModeBytes = data

CleanExit:
    On Error Resume Next
    If pDevMode <> 0 Then GlobalFree pDevMode
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Exit Function
CleanFail:
    Dim d As String, n As Long
    d = Err.Description: n = Err.Number
    On Error Resume Next
    If pDevMode <> 0 Then GlobalFree pDevMode
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Err.Raise n, , d
End Function

Private Function GetPrinterDevModeSnapshot(ByVal printerName As String) As PrinterDevModeSnapshot
#If VBA7 Then
    Dim hPrinter As LongPtr, pBuf As LongPtr, pDevMode As LongPtr
#Else
    Dim hPrinter As Long, pBuf As Long, pDevMode As Long
#End If
    Dim needed As Long, pi9 As PrinterInfo9, size As Long, data() As Byte
    Dim snapshot As PrinterDevModeSnapshot

    If OpenPrinterW(StrPtr(printerName), hPrinter, 0) = 0 Then RaiseApiError "OpenPrinterW"
    On Error GoTo CleanFail

    GetPrinterW hPrinter, 9, 0, 0, needed
    If needed > 0 Then
        pBuf = GlobalAlloc(0, needed)
        If pBuf = 0 Then Err.Raise vbObjectError + 5103, , "GlobalAlloc(GetPrinterW)に失敗しました。"
        If GetPrinterW(hPrinter, 9, pBuf, needed, needed) = 0 Then RaiseApiError "GetPrinterW(level 9)"
        CopyMemory VarPtr(pi9), pBuf, LenB(pi9)
        pDevMode = pi9.pDevMode
    End If

    If pDevMode <> 0 Then
        size = DocumentPropertiesW(0, hPrinter, StrPtr(printerName), 0, pDevMode, 0)
        If size <= 0 Then RaiseApiError "DocumentPropertiesW(size per-user)"
        ReDim data(0 To size - 1) As Byte
        CopyMemory VarPtr(data(0)), pDevMode, size
        snapshot.HasPerUserDevMode = True
        snapshot.DevModeBytes = data
    Else
        snapshot.HasPerUserDevMode = False
        snapshot.DevModeBytes = GetDriverDefaultDevModeBytesWithHandle(hPrinter, printerName)
    End If

    GetPrinterDevModeSnapshot = snapshot
CleanExit:
    On Error Resume Next
    If pBuf <> 0 Then GlobalFree pBuf
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Exit Function
CleanFail:
    Dim d As String, n As Long
    d = Err.Description: n = Err.Number
    On Error Resume Next
    If pBuf <> 0 Then GlobalFree pBuf
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Err.Raise n, , d
End Function

#If VBA7 Then
Private Function GetDriverDefaultDevModeBytesWithHandle(ByVal hPrinter As LongPtr, ByVal printerName As String) As Byte()
#Else
Private Function GetDriverDefaultDevModeBytesWithHandle(ByVal hPrinter As Long, ByVal printerName As String) As Byte()
#End If
#If VBA7 Then
    Dim pDevMode As LongPtr
#Else
    Dim pDevMode As Long
#End If
    Dim size As Long, result As Long, data() As Byte

    size = DocumentPropertiesW(0, hPrinter, StrPtr(printerName), 0, 0, 0)
    If size <= 0 Then RaiseApiError "DocumentPropertiesW(size default)"
    pDevMode = GlobalAlloc(0, size)
    If pDevMode = 0 Then Err.Raise vbObjectError + 5105, , "GlobalAlloc(DEVMODE既定値)に失敗しました。"
    On Error GoTo CleanFail
    result = DocumentPropertiesW(0, hPrinter, StrPtr(printerName), pDevMode, 0, DM_OUT_BUFFER)
    If result < 0 Then RaiseApiError "DocumentPropertiesW(default)"
    ReDim data(0 To size - 1) As Byte
    CopyMemory VarPtr(data(0)), pDevMode, size
    GetDriverDefaultDevModeBytesWithHandle = data
CleanExit:
    On Error Resume Next
    If pDevMode <> 0 Then GlobalFree pDevMode
    Exit Function
CleanFail:
    Dim d As String, n As Long
    d = Err.Description: n = Err.Number
    On Error Resume Next
    If pDevMode <> 0 Then GlobalFree pDevMode
    Err.Raise n, , d
End Function

Private Function GetUserDevModeBytes(ByVal printerName As String) As Byte()
    Dim snapshot As PrinterDevModeSnapshot
    snapshot = GetPrinterDevModeSnapshot(printerName)
    GetUserDevModeBytes = snapshot.DevModeBytes
End Function

Private Function NormalizeDevModeBytes(ByVal printerName As String, ByRef sourceBytes() As Byte) As Byte()
#If VBA7 Then
    Dim hPrinter As LongPtr, pDevMode As LongPtr
#Else
    Dim hPrinter As Long, pDevMode As Long
#End If
    Dim sourceSize As Long, requiredSize As Long, bufferSize As Long
    Dim result As Long, normalizedBytes() As Byte

    sourceSize = ByteArrayLength(sourceBytes)
    If OpenPrinterW(StrPtr(printerName), hPrinter, 0) = 0 Then RaiseApiError "OpenPrinterW"
    On Error GoTo CleanFail

    requiredSize = DocumentPropertiesW(0, hPrinter, StrPtr(printerName), 0, 0, 0)
    If requiredSize <= 0 Then RaiseApiError "DocumentPropertiesW(size normalize)"
    bufferSize = requiredSize
    If sourceSize > bufferSize Then bufferSize = sourceSize

    pDevMode = GlobalAlloc(0, bufferSize)
    If pDevMode = 0 Then Err.Raise vbObjectError + 5105, , "GlobalAlloc(DEVMODE正規化)に失敗しました。"
    CopyMemory pDevMode, VarPtr(sourceBytes(0)), sourceSize
    result = DocumentPropertiesW(0, hPrinter, StrPtr(printerName), pDevMode, pDevMode, DM_IN_BUFFER Or DM_OUT_BUFFER)
    If result < 0 Then RaiseApiError "DocumentPropertiesW(normalize)"

    ReDim normalizedBytes(0 To requiredSize - 1) As Byte
    CopyMemory VarPtr(normalizedBytes(0)), pDevMode, requiredSize
    NormalizeDevModeBytes = normalizedBytes

CleanExit:
    On Error Resume Next
    If pDevMode <> 0 Then GlobalFree pDevMode
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Exit Function
CleanFail:
    Dim d As String, n As Long
    d = Err.Description: n = Err.Number
    On Error Resume Next
    If pDevMode <> 0 Then GlobalFree pDevMode
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Err.Raise n, , d
End Function

Private Function ApplyUserDevModeBytes(ByVal printerName As String, ByRef devModeBytes() As Byte) As Byte()
#If VBA7 Then
    Dim hPrinter As LongPtr, pDevMode As LongPtr, pInfo As LongPtr
#Else
    Dim hPrinter As Long, pDevMode As Long, pInfo As Long
#End If
    Dim size As Long, pi9 As PrinterInfo9
    Dim normalizedBytes() As Byte

    normalizedBytes = NormalizeDevModeBytes(printerName, devModeBytes)
    size = ByteArrayLength(normalizedBytes)
    If OpenPrinterW(StrPtr(printerName), hPrinter, 0) = 0 Then RaiseApiError "OpenPrinterW"
    On Error GoTo CleanFail
    pDevMode = GlobalAlloc(0, size)
    pInfo = GlobalAlloc(0, LenB(pi9))
    If pDevMode = 0 Or pInfo = 0 Then Err.Raise vbObjectError + 5105, , "GlobalAlloc(DEVMODE)に失敗しました。"
    CopyMemory pDevMode, VarPtr(normalizedBytes(0)), size
    pi9.pDevMode = pDevMode
    CopyMemory pInfo, VarPtr(pi9), LenB(pi9)
    If SetPrinterW(hPrinter, 9, pInfo, 0) = 0 Then RaiseApiError "SetPrinterW(level 9)"
    ApplyUserDevModeBytes = normalizedBytes
CleanExit:
    On Error Resume Next
    If pInfo <> 0 Then GlobalFree pInfo
    If pDevMode <> 0 Then GlobalFree pDevMode
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Exit Function
CleanFail:
    Dim d As String, n As Long
    d = Err.Description: n = Err.Number
    On Error Resume Next
    If pInfo <> 0 Then GlobalFree pInfo
    If pDevMode <> 0 Then GlobalFree pDevMode
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Err.Raise n, , d
End Function

Private Sub VerifyAppliedDevMode(ByVal printerName As String, ByRef expected() As Byte)
    Dim actual() As Byte

    actual = GetUserDevModeBytes(printerName)
    If Not DevModesEquivalent(printerName, actual, expected) Then
        Err.Raise vbObjectError + 5106, , "固定印刷プロファイルの適用確認に失敗しました。ドライバーが設定を書き換えた、またはドライバー更新後の可能性があります。再登録してください。"
    End If
End Sub

Private Function DevModesEquivalent(ByVal printerName As String, ByRef leftBytes() As Byte, ByRef rightBytes() As Byte) As Boolean
    Dim normalizedLeft() As Byte
    Dim normalizedRight() As Byte

    normalizedLeft = NormalizeDevModeBytes(printerName, leftBytes)
    normalizedRight = NormalizeDevModeBytes(printerName, rightBytes)
    DevModesEquivalent = ByteArraysEqual(normalizedLeft, normalizedRight)
End Function

Private Sub ClearUserDevMode(ByVal printerName As String)
#If VBA7 Then
    Dim hPrinter As LongPtr, pInfo As LongPtr
#Else
    Dim hPrinter As Long, pInfo As Long
#End If
    Dim pi9 As PrinterInfo9

    If OpenPrinterW(StrPtr(printerName), hPrinter, 0) = 0 Then RaiseApiError "OpenPrinterW"
    On Error GoTo CleanFail
    pInfo = GlobalAlloc(0, LenB(pi9))
    If pInfo = 0 Then Err.Raise vbObjectError + 5105, , "GlobalAlloc(PRINTER_INFO_9 NULL)に失敗しました。"
    pi9.pDevMode = 0
    CopyMemory pInfo, VarPtr(pi9), LenB(pi9)
    If SetPrinterW(hPrinter, 9, pInfo, 0) = 0 Then RaiseApiError "SetPrinterW(level 9 NULL)"
CleanExit:
    On Error Resume Next
    If pInfo <> 0 Then GlobalFree pInfo
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Exit Sub
CleanFail:
    Dim d As String, n As Long
    d = Err.Description: n = Err.Number
    On Error Resume Next
    If pInfo <> 0 Then GlobalFree pInfo
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Err.Raise n, , d
End Sub

Private Function TryRestorePrinterSnapshot(ByVal printerName As String, ByRef snapshot As PrinterDevModeSnapshot, ByRef warningText As String) As Boolean
    Dim restoredDevMode() As Byte

    On Error GoTo RestoreFailed
    If Len(printerName) = 0 Or IsEmptyByteArray(snapshot.DevModeBytes) Then
        warningText = AppendWarning(warningText, "復元に必要なプリンター名またはDEVMODE退避データがありません。")
        Exit Function
    End If

    If snapshot.HasPerUserDevMode Then
        restoredDevMode = ApplyUserDevModeBytes(printerName, snapshot.DevModeBytes)
        VerifyRestoredSnapshot printerName, snapshot
        TryRestorePrinterSnapshot = True
        Exit Function
    End If

    On Error GoTo NullRestoreFailed
    ClearUserDevMode printerName
    VerifyRestoredSnapshot printerName, snapshot
    TryRestorePrinterSnapshot = True
    Exit Function

NullRestoreFailed:
    Dim nullRestoreMessage As String
    nullRestoreMessage = "ユーザー別DEVMODEなしの状態への復元に失敗したため、退避した有効DEVMODEを適用して見た目上の設定を復元しました: Err " & CStr(Err.Number) & " " & Err.Description
    On Error GoTo RestoreFailed
    restoredDevMode = ApplyUserDevModeBytes(printerName, snapshot.DevModeBytes)
    VerifyAppliedDevMode printerName, restoredDevMode
    warningText = AppendWarning(warningText, nullRestoreMessage)
    WriteFixedPrintLog "プリンター設定復元代替", "警告", nullRestoreMessage
    TryRestorePrinterSnapshot = True
    Exit Function

RestoreFailed:
    warningText = AppendWarning(warningText, "プリンター設定の復元に失敗しました: Err " & CStr(Err.Number) & " " & Err.Description)
End Function

Private Sub VerifyRestoredSnapshot(ByVal printerName As String, ByRef snapshot As PrinterDevModeSnapshot)
    Dim currentSnapshot As PrinterDevModeSnapshot
    currentSnapshot = GetPrinterDevModeSnapshot(printerName)
    If snapshot.HasPerUserDevMode Then
        If Not currentSnapshot.HasPerUserDevMode Or Not DevModesEquivalent(printerName, currentSnapshot.DevModeBytes, snapshot.DevModeBytes) Then
            Err.Raise vbObjectError + 5121, , "印刷前のユーザー別DEVMODEへ復元できたことを確認できませんでした。"
        End If
    Else
        If Not DevModesEquivalent(printerName, currentSnapshot.DevModeBytes, snapshot.DevModeBytes) Then
            Err.Raise vbObjectError + 5122, , "印刷前の有効DEVMODEへ復元できたことを確認できませんでした。"
        End If
    End If
End Sub

Private Function BytesToBase64(ByRef bytes() As Byte) As String
    Dim dom As Object, node As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    BytesToBase64 = Replace(node.Text, vbLf, "")
End Function

Private Function Base64ToBytes(ByVal text As String) As Byte()
    Dim dom As Object, node As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = text
    Base64ToBytes = node.nodeTypedValue
End Function

Private Function ByteArraysEqual(ByRef leftBytes() As Byte, ByRef rightBytes() As Byte) As Boolean
    Dim i As Long
    If ByteArrayLength(leftBytes) <> ByteArrayLength(rightBytes) Then Exit Function
    For i = LBound(leftBytes) To UBound(leftBytes)
        If leftBytes(i) <> rightBytes(i - LBound(leftBytes) + LBound(rightBytes)) Then Exit Function
    Next i
    ByteArraysEqual = True
End Function

Private Function CalculateByteChecksum(ByRef bytes() As Byte) As String
    Dim i As Long
    Dim checksum As Double
    For i = LBound(bytes) To UBound(bytes)
        checksum = checksum + (CDbl(bytes(i)) * CDbl((i - LBound(bytes) + 1)))
        checksum = checksum - Fix(checksum / 4294967291#) * 4294967291#
    Next i
    CalculateByteChecksum = Format$(checksum, "0")
End Function

Private Function ByteArrayLength(ByRef bytes() As Byte) As Long
    ByteArrayLength = UBound(bytes) - LBound(bytes) + 1
End Function

Private Function Nz(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then Nz = "" Else Nz = CStr(value)
End Function

Private Function IsEmptyByteArray(ByRef bytes() As Byte) As Boolean
    On Error GoTo EmptyArray
    Dim n As Long
    n = UBound(bytes)
    IsEmptyByteArray = False
    Exit Function
EmptyArray:
    IsEmptyByteArray = True
End Function

Private Function GetWindowsUserNameText() As String
    GetWindowsUserNameText = Environ$("USERNAME")
End Function

Private Function GetComputerNameText() As String
    GetComputerNameText = Environ$("COMPUTERNAME")
End Function

Private Function AppendWarning(ByVal currentText As String, ByVal addText As String) As String
    If Len(currentText) = 0 Then
        AppendWarning = addText
    Else
        AppendWarning = currentText & vbCrLf & addText
    End If
End Function

Private Sub RaiseApiError(ByVal apiName As String)
    Err.Raise vbObjectError + 5199, , apiName & " に失敗しました。GetLastError=" & CStr(GetLastError())
End Sub

Private Function BuildErrorMessage(ByVal stageName As String, ByVal errNumber As Long, ByVal errDescription As String) As String
    If Len(stageName) = 0 Then stageName = "不明"
    If Len(errDescription) = 0 Then errDescription = "（エラー内容が空です）"
    BuildErrorMessage = "失敗した処理段階: " & stageName & vbCrLf & _
                        "エラー番号: " & CStr(errNumber) & vbCrLf & _
                        "エラー内容: " & errDescription
End Function

Private Function BuildErrorLogNote(ByVal stageName As String, ByVal errNumber As Long, ByVal errDescription As String) As String
    If Len(stageName) = 0 Then stageName = "不明"
    If Len(errDescription) = 0 Then errDescription = "（エラー内容が空です）"
    BuildErrorLogNote = "Stage=" & stageName & "; Err=" & CStr(errNumber) & "; Description=" & errDescription
End Function

Private Sub WriteFixedPrintLog(ByVal macroName As String, ByVal resultText As String, ByVal note As String)
    On Error Resume Next
    WriteProcessLog macroName, "", "", 0, resultText, note
End Sub

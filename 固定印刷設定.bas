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
Private Const DM_IN_BUFFER As Long = 8

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

Public Sub 固定印刷設定を登録する()
    On Error GoTo ErrorHandler

    Dim info As FixedPrinterInfo
    Dim devMode() As Byte
    Dim profile As FixedPrintProfile
    Dim profilePath As String
    Dim answer As VbMsgBoxResult
    Dim loaded As FixedPrintProfile

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

    devMode = GetUserDevModeBytes(info.WindowsName)
    profile = BuildProfile(info, devMode)
    SaveProfileToLocalAppData profile
    loaded = LoadProfileFromLocalAppData(False)
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
    WriteFixedPrintLog "固定印刷設定登録失敗", "失敗", Err.Description
    MsgBox "固定印刷設定を登録できませんでした。" & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub 固定印刷設定をテストする()
    Dim info As FixedPrinterInfo
    Dim originalDevMode() As Byte
    Dim restoreDevMode As Boolean
    Dim restoreWarning As String

    On Error GoTo ErrorHandler

    Dim profile As FixedPrintProfile
    info = ResolveFixedPrinter()
    originalDevMode = GetUserDevModeBytes(info.WindowsName)
    restoreDevMode = True

    profile = LoadProfileFromLocalAppData(True)
    ValidateProfileForCurrentPc profile, info
    ApplyUserDevModeBytes info.WindowsName, profile.DevModeBytes
    VerifyAppliedDevMode info.WindowsName, profile.DevModeBytes
    If Not TryRestorePrinterDevMode(info.WindowsName, originalDevMode, restoreWarning) Then Err.Raise vbObjectError + 5201, , restoreWarning
    restoreDevMode = False
    VerifyAppliedDevMode info.WindowsName, originalDevMode

    WriteFixedPrintLog "固定設定テスト成功", "成功", info.WindowsName & "; Computer=" & GetComputerNameText() & "; User=" & GetWindowsUserNameText()
    MsgBox "固定印刷設定の適用と復元を確認しました。実際の印刷は行っていません。", vbInformation
    Exit Sub

ErrorHandler:
    Dim msg As String
    msg = Err.Description
    If restoreDevMode Then
        If Not TryRestorePrinterDevMode(info.WindowsName, originalDevMode, restoreWarning) Then msg = msg & vbCrLf & restoreWarning
    End If
    WriteFixedPrintLog "固定設定テスト失敗", "失敗", msg
    MsgBox "固定印刷設定テストに失敗しました。" & vbCrLf & msg, vbExclamation
End Sub

Public Sub 固定設定で印刷する()
    Dim oldActivePrinter As String
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean, oldDisplayAlerts As Boolean
    Dim oldPrintCommunication As Boolean, canUsePrintCommunication As Boolean
    Dim info As FixedPrinterInfo
    Dim profile As FixedPrintProfile
    Dim originalDevMode() As Byte
    Dim restoreDevMode As Boolean, restoreActivePrinter As Boolean
    Dim targetSheet As Worksheet
    Dim mainError As String, restoreWarning As String

    On Error GoTo ErrorHandler

    If Not ProfileFileExists() Then Err.Raise vbObjectError + 5301, , GetProfileNotRegisteredMessage()
    profile = LoadProfileFromLocalAppData(True)
    info = ResolveFixedPrinter()
    ValidateProfileForCurrentPc profile, info
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

    originalDevMode = GetUserDevModeBytes(info.WindowsName)
    restoreDevMode = True
    ApplyUserDevModeBytes info.WindowsName, profile.DevModeBytes
    VerifyAppliedDevMode info.WindowsName, profile.DevModeBytes

    If canUsePrintCommunication Then Application.PrintCommunication = True
    Application.ActivePrinter = info.ExcelName
    targetSheet.PrintOut

    WriteFixedPrintLog "固定印刷成功", "成功", "Sheet=" & targetSheet.Name & "; Printer=" & info.WindowsName & "; Computer=" & GetComputerNameText() & "; User=" & GetWindowsUserNameText()

Cleanup:
    If restoreDevMode Then
        If Not TryRestorePrinterDevMode(info.WindowsName, originalDevMode, restoreWarning) Then
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
    mainError = Err.Description
    Resume Cleanup
End Sub

Private Function ResolveSealPrintTargetSheet() As Worksheet
    Dim wsSingle As Worksheet, wsMulti As Worksheet
    Dim hasSingle As Boolean, hasMulti As Boolean

    Set wsSingle = ThisWorkbook.Worksheets(SINGLE_SEAL_SHEET)
    Set wsMulti = ThisWorkbook.Worksheets(MULTI_SEAL_SHEET)
    hasSingle = HasSingleSealOutputData(wsSingle)
    hasMulti = HasMultiSealOutputData(wsMulti)

    If hasSingle And Not hasMulti Then
        Set ResolveSealPrintTargetSheet = wsSingle
    ElseIf hasMulti And Not hasSingle Then
        Set ResolveSealPrintTargetSheet = wsMulti
    ElseIf Not hasSingle And Not hasMulti Then
        Err.Raise vbObjectError + 5302, , "印刷するシールデータがありません。" & vbCrLf & _
                  "先に単件または複数の内容を反映してください。"
    Else
        Err.Raise vbObjectError + 5303, , "単件用シートと複数用シートの両方に" & vbCrLf & _
                  "印刷データが確認されました。" & vbCrLf & _
                  "印刷対象を確認してください。"
    End If
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

Private Function ResolveFixedPrinter() As FixedPrinterInfo
    Dim svc As Object, printers As Object, p As Object
    Dim hitCount As Long, hitSummary As String
    Dim info As FixedPrinterInfo
    Set svc = GetObject("winmgmts:\\.\root\cimv2")
    Set printers = svc.ExecQuery("SELECT Name, ServerName, PortName, DriverName FROM Win32_Printer")
    For Each p In printers
        If IsTargetPrinter(CStr(p.Name), CStr(Nz(p.ServerName)), CStr(Nz(p.PortName))) Then
            info.WindowsName = CStr(p.Name)
            info.PortName = CStr(Nz(p.PortName))
            info.DriverName = CStr(Nz(p.DriverName))
            info.ServerName = CStr(Nz(p.ServerName))
            info.ExcelName = BuildExcelActivePrinterName(info.WindowsName, info.PortName)
            hitCount = hitCount + 1
            hitSummary = hitSummary & vbCrLf & "- " & info.WindowsName & " / Port=" & info.PortName & " / Driver=" & info.DriverName
        End If
    Next
    If hitCount = 0 Then Err.Raise vbObjectError + 5101, , "対象プリンター（土木課C4476R / PRTSV03.sojanet.local）がWindowsに登録されていません。"
    If hitCount > 1 Then Err.Raise vbObjectError + 5102, , "対象プリンター候補が複数あり一意に決められないため印刷しません。" & hitSummary
    ResolveFixedPrinter = info
End Function

Private Function IsTargetPrinter(ByVal printerName As String, ByVal serverName As String, ByVal portName As String) As Boolean
    Dim allText As String
    allText = LCase$(printerName & " " & serverName & " " & portName)
    IsTargetPrinter = (InStr(printerName, FIXED_PRINTER_NAME) > 0 And InStr(allText, LCase$(FIXED_PRINTER_SERVER)) > 0)
End Function

Private Function BuildExcelActivePrinterName(ByVal printerName As String, ByVal portName As String) As String
    BuildExcelActivePrinterName = printerName & " on " & portName
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
    EnsureProfileFolderExists
    profileText = BuildProfileFileText(profile)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(GetProfileFilePath(), True, True)
    ts.Write profileText
    ts.Close
End Sub

Private Function LoadProfileFromLocalAppData(ByVal showNotRegisteredGuide As Boolean) As FixedPrintProfile
    Dim fso As Object, ts As Object, text As String
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
    Dim profile As FixedPrintProfile
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

Private Function GetUserDevModeBytes(ByVal printerName As String) As Byte()
#If VBA7 Then
    Dim hPrinter As LongPtr, pBuf As LongPtr, pDevMode As LongPtr
#Else
    Dim hPrinter As Long, pBuf As Long, pDevMode As Long
#End If
    Dim needed As Long, pi9 As PrinterInfo9, size As Long, data() As Byte
    If OpenPrinterW(StrPtr(printerName), hPrinter, 0) = 0 Then RaiseApiError "OpenPrinterW"
    On Error GoTo CleanFail
    GetPrinterW hPrinter, 9, 0, 0, needed
    If needed <= 0 Then RaiseApiError "GetPrinterW(level 9 size)"
    pBuf = GlobalAlloc(0, needed)
    If pBuf = 0 Then Err.Raise vbObjectError + 5103, , "GlobalAlloc(GetPrinterW)に失敗しました。"
    If GetPrinterW(hPrinter, 9, pBuf, needed, needed) = 0 Then RaiseApiError "GetPrinterW(level 9)"
    CopyMemory VarPtr(pi9), pBuf, LenB(pi9)
    pDevMode = pi9.pDevMode
    If pDevMode = 0 Then Err.Raise vbObjectError + 5104, , "PRINTER_INFO_9のpDevModeが空です。プリンターのユーザー別設定を開いてから再実行してください。"
    size = DocumentPropertiesW(0, hPrinter, StrPtr(printerName), 0, pDevMode, 0)
    If size <= 0 Then RaiseApiError "DocumentPropertiesW(size)"
    ReDim data(0 To size - 1) As Byte
    CopyMemory VarPtr(data(0)), pDevMode, size
    GetUserDevModeBytes = data
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

Private Sub ApplyUserDevModeBytes(ByVal printerName As String, ByRef devModeBytes() As Byte)
#If VBA7 Then
    Dim hPrinter As LongPtr, pDevMode As LongPtr, pInfo As LongPtr
#Else
    Dim hPrinter As Long, pDevMode As Long, pInfo As Long
#End If
    Dim size As Long, result As Long, pi9 As PrinterInfo9
    size = ByteArrayLength(devModeBytes)
    If OpenPrinterW(StrPtr(printerName), hPrinter, 0) = 0 Then RaiseApiError "OpenPrinterW"
    On Error GoTo CleanFail
    pDevMode = GlobalAlloc(0, size)
    pInfo = GlobalAlloc(0, LenB(pi9))
    If pDevMode = 0 Or pInfo = 0 Then Err.Raise vbObjectError + 5105, , "GlobalAlloc(DEVMODE)に失敗しました。"
    CopyMemory pDevMode, VarPtr(devModeBytes(0)), size
    result = DocumentPropertiesW(0, hPrinter, StrPtr(printerName), pDevMode, pDevMode, DM_IN_BUFFER Or DM_OUT_BUFFER)
    If result < 0 Then RaiseApiError "DocumentPropertiesW(validate)"
    pi9.pDevMode = pDevMode
    CopyMemory pInfo, VarPtr(pi9), LenB(pi9)
    If SetPrinterW(hPrinter, 9, pInfo, 0) = 0 Then RaiseApiError "SetPrinterW(level 9)"
CleanExit:
    On Error Resume Next
    If pInfo <> 0 Then GlobalFree pInfo
    If pDevMode <> 0 Then GlobalFree pDevMode
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Exit Sub
CleanFail:
    Dim d As String, n As Long
    d = Err.Description: n = Err.Number
    On Error Resume Next
    If pInfo <> 0 Then GlobalFree pInfo
    If pDevMode <> 0 Then GlobalFree pDevMode
    If hPrinter <> 0 Then ClosePrinter hPrinter
    Err.Raise n, , d
End Sub

Private Sub VerifyAppliedDevMode(ByVal printerName As String, ByRef expected() As Byte)
    Dim actual() As Byte
    actual = GetUserDevModeBytes(printerName)
    If Not ByteArraysEqual(actual, expected) Then
        Err.Raise vbObjectError + 5106, , "固定印刷プロファイルの適用確認に失敗しました。ドライバーが設定を書き換えた、またはドライバー更新後の可能性があります。再登録してください。"
    End If
End Sub

Private Function TryRestorePrinterDevMode(ByVal printerName As String, ByRef devModeBytes() As Byte, ByRef warningText As String) As Boolean
    On Error GoTo RestoreFailed
    If Len(printerName) = 0 Or IsEmptyByteArray(devModeBytes) Then
        warningText = AppendWarning(warningText, "復元に必要なプリンター名またはDEVMODE退避データがありません。")
        Exit Function
    End If
    ApplyUserDevModeBytes printerName, devModeBytes
    VerifyAppliedDevMode printerName, devModeBytes
    TryRestorePrinterDevMode = True
    Exit Function
RestoreFailed:
    warningText = AppendWarning(warningText, "プリンター設定の復元に失敗しました: Err " & CStr(Err.Number) & " " & Err.Description)
End Function

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

Private Sub WriteFixedPrintLog(ByVal macroName As String, ByVal resultText As String, ByVal note As String)
    On Error Resume Next
    WriteProcessLog macroName, "", "", 0, resultText, note
End Sub

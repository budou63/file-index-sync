Attribute VB_Name = "固定印刷設定"
Option Explicit

'========================================================
' 土木課C4476R 固定印刷プロファイル
' - PRINTER_INFO_9（現在ユーザー別DEVMODE）をブック内VeryHiddenシートへ保存/復元する。
' - ExcelのActivePrinterポート（Ne03:等）は固定せず、Windows登録プリンターから解決する。
'========================================================

Private Const FIXED_PRINTER_NAME As String = "土木課C4476R"
Private Const FIXED_PRINTER_SERVER As String = "PRTSV03.sojanet.local"
Private Const FIXED_PRINT_SHEET As String = "個別フォルダシール（複数）"
Private Const PROFILE_SHEET_NAME As String = "_固定印刷設定"
Private Const PROFILE_CHUNK_LEN As Long = 30000
Private Const CCHDEVICENAME As Long = 32
Private Const CCHFORMNAME As Long = 32
Private Const DM_OUT_BUFFER As Long = 2
Private Const DM_IN_BUFFER As Long = 8
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
Private Const xlSheetVeryHiddenConst As Long = 2

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

Public Sub 固定印刷設定を登録する()
    On Error GoTo ErrorHandler
    Dim info As FixedPrinterInfo
    Dim devMode() As Byte

    info = ResolveFixedPrinter()
    devMode = GetUserDevModeBytes(info.WindowsName)
    SaveFixedPrintProfile info, devMode
    WriteFixedPrintLog "固定印刷設定登録", "成功", info.WindowsName & " / Driver=" & info.DriverName
    MsgBox "固定印刷設定を登録しました。" & vbCrLf & _
           "プリンター: " & info.WindowsName & vbCrLf & _
           "ドライバー: " & info.DriverName, vbInformation
    Exit Sub

ErrorHandler:
    WriteFixedPrintLog "固定印刷設定登録", "失敗", Err.Description
    MsgBox "固定印刷設定を登録できませんでした。" & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub 固定印刷設定をテストする()
    On Error GoTo ErrorHandler
    Dim info As FixedPrinterInfo
    Dim originalDevMode() As Byte
    Dim profileDevMode() As Byte

    info = ResolveFixedPrinter()
    originalDevMode = GetUserDevModeBytes(info.WindowsName)
    profileDevMode = LoadFixedPrintProfile(info)
    ApplyUserDevModeBytes info.WindowsName, profileDevMode
    VerifyAppliedDevMode info.WindowsName, profileDevMode
    ApplyUserDevModeBytes info.WindowsName, originalDevMode
    VerifyAppliedDevMode info.WindowsName, originalDevMode
    WriteFixedPrintLog "固定印刷設定テスト", "成功", info.WindowsName
    MsgBox "固定印刷設定の適用と復元を確認しました。実際の印刷は行っていません。", vbInformation
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Len(info.WindowsName) > 0 And Not IsEmptyByteArray(originalDevMode) Then ApplyUserDevModeBytes info.WindowsName, originalDevMode
    WriteFixedPrintLog "固定印刷設定テスト", "失敗", Err.Description
    MsgBox "固定印刷設定テストに失敗しました。" & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub 固定設定で印刷する()
    Dim oldActivePrinter As String
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean, oldDisplayAlerts As Boolean
    Dim oldPrintCommunication As Boolean, canUsePrintCommunication As Boolean
    Dim info As FixedPrinterInfo
    Dim originalDevMode() As Byte, profileDevMode() As Byte
    Dim restoreDevMode As Boolean, restoreActivePrinter As Boolean
    Dim ws As Worksheet

    On Error GoTo ErrorHandler
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

    info = ResolveFixedPrinter()
    originalDevMode = GetUserDevModeBytes(info.WindowsName)
    restoreDevMode = True
    profileDevMode = LoadFixedPrintProfile(info)
    ApplyUserDevModeBytes info.WindowsName, profileDevMode
    VerifyAppliedDevMode info.WindowsName, profileDevMode

    If canUsePrintCommunication Then Application.PrintCommunication = True
    Application.ActivePrinter = info.ExcelName

    Set ws = ThisWorkbook.Worksheets(FIXED_PRINT_SHEET)
    ws.PrintOut
    WriteFixedPrintLog "固定設定で印刷", "成功", info.WindowsName

Cleanup:
    On Error Resume Next
    If restoreDevMode Then ApplyUserDevModeBytes info.WindowsName, originalDevMode
    If restoreActivePrinter Then Application.ActivePrinter = oldActivePrinter
    If canUsePrintCommunication Then Application.PrintCommunication = oldPrintCommunication
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

ErrorHandler:
    Dim msg As String
    msg = Err.Description
    WriteFixedPrintLog "固定設定で印刷", "失敗", msg
    MsgBox "固定設定で印刷できませんでした。誤った設定での印刷を防ぐため中止しました。" & vbCrLf & msg, vbExclamation
    Resume Cleanup
End Sub

Private Function ResolveFixedPrinter() As FixedPrinterInfo
    Dim svc As Object, printers As Object, p As Object
    Dim hitCount As Long
    Dim hitSummary As String
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
    size = UBound(devModeBytes) - LBound(devModeBytes) + 1
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
    If BytesToBase64(actual) <> BytesToBase64(expected) Then
        Err.Raise vbObjectError + 5106, , "固定印刷プロファイルの適用確認に失敗しました。ドライバーが設定を書き換えた、またはドライバー更新後の可能性があります。再登録してください。"
    End If
End Sub

Private Sub SaveFixedPrintProfile(ByRef info As FixedPrinterInfo, ByRef devModeBytes() As Byte)
    Dim ws As Worksheet, b64 As String, chunks As Long, i As Long
    Set ws = GetProfileSheet(True)
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("項目", "値")
    b64 = BytesToBase64(devModeBytes)
    ws.Range("A2").Value = "PrinterName": ws.Range("B2").Value = FIXED_PRINTER_NAME
    ws.Range("A3").Value = "ServerName": ws.Range("B3").Value = FIXED_PRINTER_SERVER
    ws.Range("A4").Value = "WindowsName": ws.Range("B4").Value = info.WindowsName
    ws.Range("A5").Value = "DriverName": ws.Range("B5").Value = info.DriverName
    ws.Range("A6").Value = "RegisteredAt": ws.Range("B6").Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("A7").Value = "DevModeBase64Length": ws.Range("B7").Value = Len(b64)
    chunks = (Len(b64) + PROFILE_CHUNK_LEN - 1) \ PROFILE_CHUNK_LEN
    ws.Cells(8, 1).Value = "ChunkCount": ws.Cells(8, 2).Value = chunks
    For i = 1 To chunks
        ws.Cells(8 + i, 1).Value = "DevModeBase64_" & Format$(i, "0000")
        ws.Cells(8 + i, 2).Value = Mid$(b64, (i - 1) * PROFILE_CHUNK_LEN + 1, PROFILE_CHUNK_LEN)
    Next
    ws.Visible = xlSheetVeryHiddenConst
End Sub

Private Function LoadFixedPrintProfile(ByRef info As FixedPrinterInfo) As Byte()
    Dim ws As Worksheet, driverName As String, b64 As String, chunks As Long, i As Long
    Set ws = GetProfileSheet(False)
    If ws Is Nothing Then Err.Raise vbObjectError + 5107, , "固定印刷プロファイルが未登録です。先に「固定印刷設定を登録する」を実行してください。"
    driverName = CStr(ws.Range("B5").Value)
    If Len(driverName) > 0 And driverName <> info.DriverName Then
        Err.Raise vbObjectError + 5108, , "登録時と現在のプリンタードライバーが異なります。登録時=" & driverName & " / 現在=" & info.DriverName & "。固定印刷設定を再登録してください。"
    End If
    chunks = CLng(Val(ws.Range("B8").Value))
    If chunks <= 0 Then Err.Raise vbObjectError + 5109, , "固定印刷プロファイルのBase64データがありません。再登録してください。"
    For i = 1 To chunks
        b64 = b64 & CStr(ws.Cells(8 + i, 2).Value)
    Next
    LoadFixedPrintProfile = Base64ToBytes(b64)
End Function

Private Function GetProfileSheet(ByVal createIfMissing As Boolean) As Worksheet
    On Error Resume Next
    Set GetProfileSheet = ThisWorkbook.Worksheets(PROFILE_SHEET_NAME)
    On Error GoTo 0
    If GetProfileSheet Is Nothing And createIfMissing Then
        Set GetProfileSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetProfileSheet.Name = PROFILE_SHEET_NAME
    End If
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

Private Function Nz(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then Nz = "" Else Nz = CStr(value)
End Function

Private Function IsEmptyByteArray(ByRef bytes() As Byte) As Boolean
    On Error GoTo EmptyArray
    Dim n As Long: n = UBound(bytes)
    IsEmptyByteArray = False
    Exit Function
EmptyArray:
    IsEmptyByteArray = True
End Function

Private Sub RaiseApiError(ByVal apiName As String)
    Err.Raise vbObjectError + 5199, , apiName & " に失敗しました。GetLastError=" & CStr(GetLastError())
End Sub

Private Sub WriteFixedPrintLog(ByVal macroName As String, ByVal resultText As String, ByVal note As String)
    On Error Resume Next
    WriteProcessLog macroName, "", "", 0, resultText, note
End Sub

Option Explicit

'========================================================
' 貼り付け先:
'   VBAProject
'   └ Microsoft Excel Objects
'      └ 新ファイル基準表
'
' ※標準モジュールではなく、必ず「新ファイル基準表」の
'   シートコードモジュールへ配置してください。
'========================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ExitHandler
    If Target Is Nothing Then Exit Sub
    If Intersect(Target, Me.Rows(1)) Is Nothing Then
        ' OK
    ElseIf Target.Rows.Count = 1 And Target.Row = 1 Then
        Exit Sub
    End If

    Application.EnableEvents = False

    ' 標準モジュール(Module1)の補完処理を呼び出す
    ' ※この中で「年度（和暦）→年度（西暦）」および
    '   「廃棄年月日自動計算」まで連続実行する
    新ファイル基準表_表示名からID自動補完 Target
    ' 追記: 色名が変更された行は 2色名 へ同値同期
    新ファイル基準表_色名から2色名同期 Target, Me

    ' 呼び出し先でイベント状態が変わる可能性に備えて再度OFF
    Application.EnableEvents = False

    ' タイトル/主要編集列の変更時は、通し番号を自動で再採番
    If ShouldReindexSerial(Target, Me) Then
        ReindexSerial_NewFileStandard Me
    End If

ExitHandler:
    Application.EnableEvents = True

End Sub

Private Function ShouldReindexSerial(ByVal Target As Range, ByVal ws As Worksheet) As Boolean

    Dim monitorHeaders As Variant
    Dim monitorRange As Range
    Dim workRange As Range
    Dim colNo As Long
    Dim i As Long

    ShouldReindexSerial = False
    If Target Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function

    monitorHeaders = Array("タイトル", "分類名２", "分類名３", "年度（和暦）", "保存期間")

    For i = LBound(monitorHeaders) To UBound(monitorHeaders)
        colNo = GetHeaderColumnByCandidates(ws, GetReindexHeaderCandidates(CStr(monitorHeaders(i))))
        If colNo > 0 Then
            If monitorRange Is Nothing Then
                Set monitorRange = ws.Range(ws.Cells(2, colNo), ws.Cells(ws.Rows.Count, colNo))
            Else
                Set monitorRange = Union(monitorRange, ws.Range(ws.Cells(2, colNo), ws.Cells(ws.Rows.Count, colNo)))
            End If
        End If
    Next i

    If monitorRange Is Nothing Then Exit Function

    Set workRange = Intersect(Target, monitorRange)
    ShouldReindexSerial = Not (workRange Is Nothing)

End Function

Private Function GetReindexHeaderCandidates(ByVal headerName As String) As Variant

    Select Case headerName
        Case "分類名２"
            GetReindexHeaderCandidates = Array("分類名２", "分類名2")
        Case "分類名３"
            GetReindexHeaderCandidates = Array("分類名３", "分類名3")
        Case "年度（和暦）"
            GetReindexHeaderCandidates = Array("年度（和暦）", "年度(和暦)")
        Case Else
            GetReindexHeaderCandidates = Array(headerName)
    End Select

End Function

Private Function GetHeaderColumnByCandidates(ByVal ws As Worksheet, ByVal candidates As Variant) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim i As Long
    Dim headerValue As String

    GetHeaderColumnByCandidates = 0
    If ws Is Nothing Then Exit Function

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol <= 0 Then Exit Function

    For c = 1 To lastCol
        headerValue = Trim(CStr(ws.Cells(1, c).Value))
        For i = LBound(candidates) To UBound(candidates)
            If StrComp(headerValue, CStr(candidates(i)), vbTextCompare) = 0 Then
                GetHeaderColumnByCandidates = c
                Exit Function
            End If
        Next i
    Next c

End Function

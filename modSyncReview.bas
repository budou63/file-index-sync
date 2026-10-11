Attribute VB_Name = "modSyncReview"
Option Explicit

Private Const REVIEW_SHEET As String = "同期候補確認履歴"
Private Const CSV_SHEET As String = "目録CSV"
Private Const DEST_SHEET As String = "新ファイル基準表"
Private reviewHeaderCache As Object
Private reviewLastCache As Object

Public Sub 確認候補を開く()
    Dim rawYear As String, yearValue As Long, cases As Collection
    On Error GoTo Failed
    rawYear = Trim$(InputBox("確認する対象年度を入力してください（例：R8）。CSV整形後に実行します。", "同期候補の確認", "R8"))
    If Len(rawYear) = 0 Then Exit Sub
    yearValue = ReviewYear(rawYear)
    If yearValue < 1900 Or yearValue > 2200 Then Err.Raise vbObjectError + 6101, , "対象年度を判定できません。"
    Set cases = SyncReviewCandidates(yearValue)
    frmSyncReview.LoadCases cases, yearValue
    frmSyncReview.Show vbModeless
    Exit Sub
Failed:
    MsgBox "確認候補を開けません: " & Err.Description, vbExclamation
End Sub

Private Function ReviewCol(ByVal ws As Worksheet, ParamArray aliases() As Variant) As Long
    Dim col As Long, a As Variant, h As String, found As Long, chosen As Long
    Dim cacheKey As String, rowNo As Long, selectedText As String, otherText As String
    If reviewHeaderCache Is Nothing Then Set reviewHeaderCache = CreateObject("Scripting.Dictionary")
    cacheKey = ws.Name
    For Each a In aliases: cacheKey = cacheKey & Chr$(30) & CStr(a): Next a
    If reviewHeaderCache.Exists(cacheKey) Then ReviewCol = CLng(reviewHeaderCache(cacheKey)): Exit Function
    For Each a In aliases
        found = 0
        For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            h = Trim$(CStr(ws.Cells(1, col).Value2))
            If StrComp(h, CStr(a), vbTextCompare) = 0 Then
                If found > 0 Then Err.Raise vbObjectError + 6102, , "同一見出しが重複しています: " & h
                found = col
            End If
        Next col
        If found > 0 Then
            If chosen = 0 Then
                chosen = found
            Else
                For rowNo = 2 To ReviewLast(ws)
                    If IsError(ws.Cells(rowNo, chosen).Value2) Or IsError(ws.Cells(rowNo, found).Value2) Then _
                        Err.Raise vbObjectError + 6103, , "別名見出しの値にエラーがあります。"
                    selectedText = Trim$(CStr(ws.Cells(rowNo, chosen).Value2))
                    otherText = Trim$(CStr(ws.Cells(rowNo, found).Value2))
                    If Len(otherText) > 0 And (Len(selectedText) = 0 Or _
                        StrComp(selectedText, otherText, vbBinaryCompare) <> 0) Then _
                        Err.Raise vbObjectError + 6102, , "別名見出しの内容が一致しません: " & ws.Name & " 行" & rowNo
                Next rowNo
            End If
        End If
    Next a
    reviewHeaderCache.Add cacheKey, chosen
    ReviewCol = chosen
End Function

Private Function ReviewValue(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long) As String
    If r < 2 Or col < 1 Then Exit Function
    If IsError(ws.Cells(r, col).Value2) Then Err.Raise vbObjectError + 6103, , "セルにエラー値があります: " & ws.Name & " 行" & r
    ReviewValue = Trim$(CStr(ws.Cells(r, col).Value2))
End Function

Private Function ReviewYear(ByVal raw As String) As Long
    Dim n As Long, text As String
    text = UCase$(Trim$(raw))
    text = Replace(text, "令和", "R")
    text = Replace(text, "年度", "")
    text = Replace(text, "年", "")
    If Len(text) < 1 Then Exit Function
    If Left$(text, 1) = "R" Then
        If Not IsNumeric(Mid$(text, 2)) Then Exit Function
        n = CLng(Mid$(text, 2))
        If n >= 1 And n <= 100 Then ReviewYear = n + 2018
    ElseIf IsNumeric(text) Then
        n = CLng(text)
        If n >= 1900 And n <= 2200 Then ReviewYear = n
    End If
End Function

Private Function ReviewFiscal(ByVal ws As Worksheet, ByVal r As Long) As Long
    ReviewFiscal = ReviewYear(ReviewValue(ws, r, ReviewCol(ws, "年度（和暦）", "年度(和暦)")))
    If ReviewFiscal = 0 Then ReviewFiscal = ReviewYear(ReviewValue(ws, r, ReviewCol(ws, "年度")))
End Function

Private Function ReviewTextKey(ByVal s As String) As String
    s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " "): s = Replace(s, vbTab, " ")
    s = Replace(s, "　", " "): s = Replace(s, "（", "("): s = Replace(s, "）", ")")
    s = Replace(s, "／", "/"): s = Replace(s, "・", "/"): s = Replace(s, "･", "/")
    On Error Resume Next
    s = StrConv(s, vbNarrow)
    On Error GoTo 0
    ReviewTextKey = LCase$(Replace(Trim$(s), " ", ""))
End Function

Private Function ReviewName(ByVal ws As Worksheet, ByVal r As Long, ByVal isCsv As Boolean, ByVal includeFolder As Boolean) As String
    Dim created As Long, g1 As String, g2 As String, folder As String
    created = ReviewYear(ReviewValue(ws, r, ReviewCol(ws, "作成年度（和暦）", "作成年度(和暦)")))
    If created = 0 Then created = ReviewYear(ReviewValue(ws, r, ReviewCol(ws, "作成年度")))
    If isCsv Then
        g1 = ReviewValue(ws, r, ReviewCol(ws, "分類名２", "分類名2", "第1ガイド", "第１ガイド"))
        g2 = ReviewValue(ws, r, ReviewCol(ws, "分類名３", "分類名3", "第2ガイド", "第２ガイド"))
        folder = ReviewValue(ws, r, ReviewCol(ws, "タイトル", "個別フォルダー", "個別フォルダ"))
    Else
        g1 = ReviewValue(ws, r, ReviewCol(ws, "第1ガイド", "第１ガイド", "分類名２"))
        g2 = ReviewValue(ws, r, ReviewCol(ws, "第2ガイド", "第２ガイド", "分類名３"))
        folder = ReviewValue(ws, r, ReviewCol(ws, "個別フォルダー", "個別フォルダ", "タイトル"))
    End If
    If created = 0 Or Len(g1) = 0 Or Len(g2) = 0 Then Exit Function
    If includeFolder And Len(folder) = 0 Then Exit Function
    ReviewName = CStr(created) & "|" & ReviewTextKey(g1) & "|" & ReviewTextKey(g2)
    If includeFolder Then ReviewName = ReviewName & "|" & ReviewTextKey(folder)
End Function

Private Function ReviewSimilar(ByVal a As String, ByVal b As String) As Boolean
    a = ReviewTextKey(Replace(a, " ", ""))
    b = ReviewTextKey(Replace(b, " ", ""))
    If Len(a) < 5 Or Len(b) < 5 Then Exit Function
    ReviewSimilar = (Left$(a, 5) = Left$(b, 5) Or InStr(1, a, b, vbTextCompare) > 0 Or _
        InStr(1, b, a, vbTextCompare) > 0)
End Function

Private Function ReviewPadding(ByVal s As String) As String
    Dim i As Long, c As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then ReviewPadding = s: Exit Function
    Next i
    i = 1
    Do While i < Len(s) And Mid$(s, i, 1) = "0": i = i + 1: Loop
    ReviewPadding = Mid$(s, i)
End Function

Private Function ReviewField(ByVal ws As Worksheet, ByVal r As Long, ByVal fieldName As String) As String
    ReviewField = ReviewValue(ws, r, ReviewCol(ws, fieldName))
End Function

Private Function ReviewLast(ByVal ws As Worksheet) As Long
    Dim cell As Range
    If ws.Name <> REVIEW_SHEET Then
        If Not reviewLastCache Is Nothing Then
            If reviewLastCache.Exists(ws.Name) Then ReviewLast = CLng(reviewLastCache(ws.Name)): Exit Function
        End If
    End If
    Set cell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If cell Is Nothing Then ReviewLast = 1 Else ReviewLast = cell.Row
    If ws.Name <> REVIEW_SHEET Then
        If reviewLastCache Is Nothing Then Set reviewLastCache = CreateObject("Scripting.Dictionary")
        reviewLastCache(ws.Name) = ReviewLast
    End If
End Function

Private Function ReviewSignature(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim col As Long, valueText As String, lastCol As Long
    If r < 2 Or r > ReviewLast(ws) Then Exit Function
    lastCol = Application.Max(ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column, _
        ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column)
    ReviewSignature = ws.Name & "|" & CStr(r) & "|" & CStr(lastCol) & "|"
    For col = 1 To lastCol
        If IsError(ws.Cells(r, col).Value2) Then Err.Raise vbObjectError + 6135, , "行にエラー値があります。"
        valueText = TypeName(ws.Cells(r, col).Value2) & ":" & CStr(ws.Cells(r, col).Value2)
        If ws.Cells(r, col).HasFormula Then valueText = valueText & "#" & CStr(ws.Cells(r, col).Formula)
        ReviewSignature = ReviewSignature & CStr(Len(valueText)) & ":" & valueText
    Next col
    ReviewSignature = ReviewSignature & "|id-format=" & _
        ws.Cells(r, ReviewCol(ws, "ファイルID")).NumberFormat & "|no-format=" & _
        ws.Cells(r, ReviewCol(ws, "ファイル番号")).NumberFormat & "|bar-format=" & _
        ws.Cells(r, ReviewCol(ws, "バーコード番号")).NumberFormat
End Function

Private Function ReviewDescription(ByVal ws As Worksheet, ByVal r As Long, ByVal isCsv As Boolean) As String
    Dim nameCol As Long, g1Col As Long, g2Col As Long
    If isCsv Then
        nameCol = ReviewCol(ws, "タイトル", "個別フォルダー")
        g1Col = ReviewCol(ws, "分類名２", "第1ガイド")
        g2Col = ReviewCol(ws, "分類名３", "第2ガイド")
    Else
        nameCol = ReviewCol(ws, "個別フォルダー", "タイトル")
        g1Col = ReviewCol(ws, "第1ガイド", "分類名２")
        g2Col = ReviewCol(ws, "第2ガイド", "分類名３")
    End If
    ReviewDescription = ws.Name & " 行" & r & " / ID=" & ReviewField(ws, r, "ファイルID") & _
        " / 番号=" & ReviewField(ws, r, "ファイル番号") & " / バーコード=" & ReviewField(ws, r, "バーコード番号") & _
        " / 年度=" & CStr(ReviewFiscal(ws, r)) & " / 作成年度=" & ReviewValue(ws, r, ReviewCol(ws, "作成年度（和暦）", "作成年度(和暦)")) & _
        " / 第1=" & ReviewValue(ws, r, g1Col) & " / 第2=" & ReviewValue(ws, r, g2Col) & _
        " / フォルダー=" & ReviewValue(ws, r, nameCol)
    If Not isCsv Then ReviewDescription = ReviewDescription & " / 同期状態=" & ReviewField(ws, r, "同期状態")
End Function

Private Function NewCase(ByVal category As String, ByVal yearValue As Long, ByVal src As Worksheet, _
    ByVal srcRow As Long, ByVal dst As Worksheet, ByVal rows As Collection, ByVal reason As String) As Object
    Dim item As Object, v As Variant, rowText As String, rowSignatures As String, details As String
    Set item = CreateObject("Scripting.Dictionary")
    For Each v In rows
        rowText = rowText & IIf(Len(rowText) > 0, ";", "") & CStr(v)
        rowSignatures = rowSignatures & ReviewSignature(dst, CLng(v))
        details = details & vbCrLf & ReviewDescription(dst, CLng(v), False)
    Next v
    item.Add "category", category
    item.Add "year", yearValue
    item.Add "sourceRow", srcRow
    item.Add "rows", rowText
    item.Add "reason", reason
    item.Add "sourceSignature", IIf(srcRow > 0, ReviewSignature(src, srcRow), "")
    item.Add "destinationSignature", rowSignatures
    item.Add "details", IIf(srcRow > 0, ReviewDescription(src, srcRow, True), "CSV対象なし") & vbCrLf & reason & vbCrLf & details
    item.Add "key", category & "|" & CStr(yearValue) & "|" & CStr(srcRow) & "|" & rowText
    Set NewCase = item
End Function

Private Sub ReviewMapAdd(ByVal map As Object, ByVal key As String, ByVal rowNo As Long)
    Dim rows As Collection
    If Len(key) = 0 Then Exit Sub
    If Not map.Exists(key) Then
        Set rows = New Collection
        map.Add key, rows
    End If
    Set rows = map(key)
    rows.Add rowNo
End Sub

Private Function ReviewRows(ByVal map As Object, ByVal key As String) As Collection
    Dim rows As New Collection
    If Len(key) > 0 Then
        If map.Exists(key) Then Set rows = map(key)
    End If
    Set ReviewRows = rows
End Function

Private Sub ReviewAppendUnique(ByVal rows As Collection, ByVal rowNo As Long)
    Dim v As Variant
    For Each v In rows
        If CLng(v) = rowNo Then Exit Sub
    Next v
    rows.Add rowNo
End Sub

Private Function ReviewCandidateRows(ByVal src As Worksheet, ByVal sRow As Long, ByVal dst As Worksheet, _
    ByVal idIndex As Object, ByVal barIndex As Object, ByVal noIndex As Object, ByVal nameIndex As Object, _
    ByVal guideIndex As Object) As Collection
    Dim rows As New Collection, v As Variant, id As String, bar As String, no As String, name As String, guide As String, sourceFolder As String
    id = ReviewField(src, sRow, "ファイルID")
    bar = ReviewField(src, sRow, "バーコード番号")
    no = ReviewPadding(ReviewField(src, sRow, "ファイル番号"))
    name = ReviewName(src, sRow, True, True)
    guide = ReviewName(src, sRow, True, False)
    sourceFolder = ReviewValue(src, sRow, ReviewCol(src, "タイトル", "個別フォルダー"))
    For Each v In ReviewRows(idIndex, id): ReviewAppendUnique rows, CLng(v): Next v
    For Each v In ReviewRows(barIndex, bar): ReviewAppendUnique rows, CLng(v): Next v
    For Each v In ReviewRows(noIndex, no): ReviewAppendUnique rows, CLng(v): Next v
    For Each v In ReviewRows(nameIndex, name)
        If ReviewField(dst, CLng(v), "同期状態") = "登録予定" Then ReviewAppendUnique rows, CLng(v)
    Next v
    If Len(guide) > 0 Then
        For Each v In ReviewRows(guideIndex, guide)
            If ReviewField(dst, CLng(v), "同期状態") = "登録予定" And _
                Len(ReviewField(dst, CLng(v), "ファイルID")) = 0 And _
                ReviewSimilar(sourceFolder, ReviewValue(dst, CLng(v), ReviewCol(dst, "個別フォルダー", "タイトル"))) Then _
                    ReviewAppendUnique rows, CLng(v)
        Next v
    End If
    Set ReviewCandidateRows = rows
End Function

Public Function SyncReviewCandidates(ByVal yearValue As Long) As Collection
    Dim src As Worksheet, dst As Worksheet, cases As New Collection
    Dim ids As Object, bars As Object, numbers As Object, names As Object, guides As Object, csvIds As Object
    Dim csvBars As Object, csvNumbers As Object, csvNames As Object, csvFolders As Object, targetCount As Long, dRow As Long, sRow As Long
    Dim srcLast As Long, dstLast As Long, destRows As Collection, idRows As Collection, v As Variant
    Dim fid As String, bar As String, no As String, name As String, guide As String
    Dim reason As String, collision As Boolean, planned As Boolean, chosen As Collection
    If yearValue < 1900 Or yearValue > 2200 Then Err.Raise vbObjectError + 6104, , "対象年度が不正です。"
    Set reviewHeaderCache = CreateObject("Scripting.Dictionary")
    Set reviewLastCache = CreateObject("Scripting.Dictionary")
    Set src = ThisWorkbook.Worksheets(CSV_SHEET)
    Set dst = ThisWorkbook.Worksheets(DEST_SHEET)
    If ReviewCol(src, "ファイルID") = 0 Or ReviewCol(src, "ファイル番号") = 0 Or _
        ReviewCol(src, "バーコード番号") = 0 Or ReviewCol(dst, "同期状態") = 0 Or _
        ReviewCol(dst, "ファイルID") = 0 Or ReviewCol(dst, "ファイル番号") = 0 Or _
        ReviewCol(dst, "バーコード番号") = 0 Then Err.Raise vbObjectError + 6105, , "必要な見出しが不足しています。"
    Set ids = CreateObject("Scripting.Dictionary")
    Set bars = CreateObject("Scripting.Dictionary")
    Set numbers = CreateObject("Scripting.Dictionary")
    Set names = CreateObject("Scripting.Dictionary")
    Set guides = CreateObject("Scripting.Dictionary")
    Set csvIds = CreateObject("Scripting.Dictionary")
    Set csvBars = CreateObject("Scripting.Dictionary")
    Set csvNumbers = CreateObject("Scripting.Dictionary")
    Set csvNames = CreateObject("Scripting.Dictionary")
    Set csvFolders = CreateObject("Scripting.Dictionary")
    srcLast = ReviewLast(src)
    dstLast = ReviewLast(dst)
    For dRow = 2 To dstLast
        ReviewMapAdd ids, ReviewField(dst, dRow, "ファイルID"), dRow
        ReviewMapAdd bars, ReviewField(dst, dRow, "バーコード番号"), dRow
        ReviewMapAdd numbers, ReviewPadding(ReviewField(dst, dRow, "ファイル番号")), dRow
        ReviewMapAdd names, ReviewName(dst, dRow, False, True), dRow
        ReviewMapAdd guides, ReviewName(dst, dRow, False, False), dRow
    Next dRow
    For sRow = 2 To srcLast
        ReviewMapAdd csvIds, ReviewField(src, sRow, "ファイルID"), sRow
        ReviewMapAdd csvBars, ReviewField(src, sRow, "バーコード番号"), sRow
        ReviewMapAdd csvNumbers, ReviewPadding(ReviewField(src, sRow, "ファイル番号")), sRow
        ReviewMapAdd csvNames, ReviewName(src, sRow, True, True), sRow
        ReviewMapAdd csvFolders, ReviewTextKey(ReviewValue(src, sRow, ReviewCol(src, "タイトル", "個別フォルダー"))), sRow
        If ReviewFiscal(src, sRow) = yearValue Then targetCount = targetCount + 1
    Next sRow
    If targetCount = 0 Then Err.Raise vbObjectError + 6136, , "対象年度の目録CSVが0行です。最新CSVを取り込んでください。"
    For sRow = 2 To srcLast
        If ReviewFiscal(src, sRow) <> yearValue Then GoTo NextSource
        fid = ReviewField(src, sRow, "ファイルID")
        bar = ReviewField(src, sRow, "バーコード番号")
        no = ReviewPadding(ReviewField(src, sRow, "ファイル番号"))
        name = ReviewName(src, sRow, True, True)
        guide = ReviewName(src, sRow, True, False)
        Set idRows = ReviewRows(ids, fid)
        Set destRows = ReviewCandidateRows(src, sRow, dst, ids, bars, numbers, names, guides)
        collision = False: planned = False: reason = ""
        If Len(fid) = 0 Then
            collision = True: reason = "CSVに正式ファイルIDがありません。"
        ElseIf ReviewRows(csvIds, fid).Count > 1 Or idRows.Count > 1 Then
            collision = True: reason = "正式IDがCSVまたはExcelに複数あります。自動統合禁止。"
        ElseIf Len(bar) > 0 And ReviewRows(bars, bar).Count > 0 Then
            For Each v In ReviewRows(bars, bar)
                If ReviewField(dst, CLng(v), "ファイルID") <> fid Then collision = True
            Next v
            If collision Then reason = "別のExcel行とバーコードが競合します。"
        End If
        If Not collision And Len(no) > 0 Then
            For Each v In ReviewRows(numbers, no)
                If ReviewField(dst, CLng(v), "ファイルID") <> fid Then collision = True
            Next v
            If collision Then reason = "別のExcel行とファイル番号が競合します。"
        End If
        If Not collision Then
            For Each v In destRows
                If ReviewField(dst, CLng(v), "同期状態") = "登録予定" And _
                    Len(ReviewField(dst, CLng(v), "ファイルID")) = 0 Then planned = True
            Next v
            If planned Then
                reason = "同名または同一作成年度・第1/第2ガイドの類似候補です。名前だけで統合しません。"
                cases.Add NewCase("B", yearValue, src, sRow, dst, destRows, reason)
            End If
        Else
            cases.Add NewCase("A", yearValue, src, sRow, dst, destRows, reason)
        End If
NextSource:
    Next sRow
    For dRow = 2 To dstLast
        If ReviewFiscal(dst, dRow) <> yearValue Then GoTo NextDestination
        If ReviewField(dst, dRow, "同期状態") <> "削除予定" Then GoTo NextDestination
        fid = ReviewField(dst, dRow, "ファイルID")
        bar = ReviewField(dst, dRow, "バーコード番号")
        no = ReviewPadding(ReviewField(dst, dRow, "ファイル番号"))
        name = ReviewName(dst, dRow, False, True)
        If Len(name) = 0 Or _
           ReviewRows(csvIds, fid).Count > 0 Or ReviewRows(csvBars, bar).Count > 0 Or _
           ReviewRows(csvNumbers, no).Count > 0 Or ReviewRows(csvNames, name).Count > 0 Or _
           ReviewRows(csvFolders, ReviewTextKey(ReviewValue(dst, dRow, ReviewCol(dst, "個別フォルダー", "タイトル")))).Count > 0 Then GoTo NextDestination
        If Len(fid) > 0 Then
            If ReviewRows(ids, fid).Count > 1 Then GoTo NextDestination
        End If
        Set chosen = New Collection
        chosen.Add dRow
        cases.Add NewCase("C", yearValue, src, 0, dst, chosen, "CSVに同一ID・バーコード・番号・名称の候補がありません。削除前に再検証します。")
NextDestination:
    Next dRow
    Set SyncReviewCandidates = cases
End Function

Private Function ReviewDigest(ByVal text As String) As String
    Dim h1 As Double, h2 As Double, i As Long, n As Long
    h1 = 7#: h2 = 11#
    For i = 1 To Len(text)
        n = AscW(Mid$(text, i, 1))
        If n < 0 Then n = n + 65536
        h1 = h1 * 131# + n
        h1 = h1 - Fix(h1 / 2147483629#) * 2147483629#
        h2 = h2 * 137# + n
        h2 = h2 - Fix(h2 / 2147483587#) * 2147483587#
    Next i
    ReviewDigest = CStr(Len(text)) & ":" & CStr(CLng(h1)) & ":" & CStr(CLng(h2))
End Function

Private Function ReviewAudit(ByVal createIfMissing As Boolean) As Worksheet
    Dim ws As Worksheet, columns As Variant, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(REVIEW_SHEET)
    On Error GoTo 0
    If ws Is Nothing And createIfMissing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = REVIEW_SHEET
        columns = Array("時刻", "対象年度", "分類", "案件キー", "判断", "CSV行", "Excel行", "理由", "CSV署名", "Excel署名", "変更内容", "事前バックアップ")
        For i = 0 To UBound(columns): ws.Cells(1, i + 1).Value2 = columns(i): Next i
        ws.Rows(1).Font.Bold = True
    End If
    Set ReviewAudit = ws
End Function

Private Function ReviewAuditHashMatches(ByVal stored As Variant, ByVal expected As String) As Boolean
    If CStr(stored) = expected Then ReviewAuditHashMatches = True: Exit Function
    If expected = ReviewDigest("") Then
        If IsEmpty(stored) Then ReviewAuditHashMatches = True: Exit Function
        If Len(Trim$(CStr(stored))) = 0 Then ReviewAuditHashMatches = True: Exit Function
        If IsNumeric(stored) Then
            ReviewAuditHashMatches = (CDbl(stored) = 0# Or _
                Abs(CDbl(stored) - CDbl(TimeValue(expected))) < 0.000000001#)
        End If
    End If
End Function

Public Function SyncReviewDecision(ByVal item As Object) As String
    Dim ws As Worksheet, r As Long, sourceHash As String, destHash As String
    Set ws = ReviewAudit(False)
    If ws Is Nothing Then Exit Function
    sourceHash = ReviewDigest(CStr(item("sourceSignature")))
    destHash = ReviewDigest(CStr(item("destinationSignature")))
    For r = ReviewLast(ws) To 2 Step -1
        If CStr(ws.Cells(r, 4).Value2) = CStr(item("key")) And _
           ReviewAuditHashMatches(ws.Cells(r, 9).Value2, sourceHash) And _
            ReviewAuditHashMatches(ws.Cells(r, 10).Value2, destHash) Then
            Select Case CStr(ws.Cells(r, 5).Value2)
                Case "別物", "保留", "同一として処理", "削除完了"
                    SyncReviewDecision = CStr(ws.Cells(r, 5).Value2)
                    Exit Function
            End Select
        End If
    Next r
End Function

Private Sub ReviewWriteAudit(ByVal item As Object, ByVal decision As String, ByVal detail As String, ByVal backup As String)
    Dim ws As Worksheet, r As Long
    Set ws = ReviewAudit(True)
    r = ReviewLast(ws) + 1
    ws.Cells(r, 1).Value2 = Now
    ws.Cells(r, 2).Value2 = CLng(item("year"))
    ws.Cells(r, 3).Value2 = CStr(item("category"))
    ws.Cells(r, 4).NumberFormat = "@"
    ws.Cells(r, 9).NumberFormat = "@": ws.Cells(r, 10).NumberFormat = "@"
    ws.Cells(r, 4).Value2 = CStr(item("key"))
    ws.Cells(r, 5).Value2 = decision
    ws.Cells(r, 6).Value2 = CLng(item("sourceRow"))
    ws.Cells(r, 7).Value2 = CStr(item("rows"))
    ws.Cells(r, 8).Value2 = CStr(item("reason"))
    ws.Cells(r, 9).Value2 = ReviewDigest(CStr(item("sourceSignature")))
    ws.Cells(r, 10).Value2 = ReviewDigest(CStr(item("destinationSignature")))
    ws.Cells(r, 11).Value2 = detail
    ws.Cells(r, 12).Value2 = backup
End Sub

Private Function ReviewCurrent(ByVal item As Object, Optional ByVal current As Collection) As Boolean
    Dim candidate As Object, i As Long
    If current Is Nothing Then Set current = SyncReviewCandidates(CLng(item("year")))
    For i = 1 To current.Count
        Set candidate = current(i)
        If CStr(candidate("key")) = CStr(item("key")) And _
           CStr(candidate("sourceSignature")) = CStr(item("sourceSignature")) And _
           CStr(candidate("destinationSignature")) = CStr(item("destinationSignature")) Then
            ReviewCurrent = True
            Exit Function
        End If
    Next i
End Function

Public Sub SyncReviewNavigate(ByVal item As Object, ByVal selectedRow As Long)
    Dim dst As Worksheet
    Set dst = ThisWorkbook.Worksheets(DEST_SHEET)
    If selectedRow < 2 Or selectedRow > ReviewLast(dst) Or Not ReviewCurrent(item) Then _
        Err.Raise vbObjectError + 6106, , "対象行が移動・変更されました。候補を再読み込みしてください。"
    If InStr(1, ";" & CStr(item("rows")) & ";", ";" & CStr(selectedRow) & ";", vbBinaryCompare) = 0 Then _
        Err.Raise vbObjectError + 6107, , "候補外の行へは移動できません。"
    Application.Goto dst.Cells(selectedRow, 1), True
End Sub

Public Function SyncReviewDecide(ByVal item As Object, ByVal decision As String) As Boolean
    On Error GoTo Failed
    If decision <> "別物" And decision <> "保留" Then Err.Raise vbObjectError + 6108, , "記録できない判断です。"
    If Not ReviewCurrent(item) Then Err.Raise vbObjectError + 6109, , "CSVまたはExcelの行・内容・候補が変わりました。再読み込みしてください。"
    ReviewWriteAudit item, decision, "データ変更なし。次回の変更時には再判定", ""
    SyncReviewDecide = True
    Exit Function
Failed:
    MsgBox "判断を記録できません: " & Err.Description, vbExclamation
End Function

Public Function SyncReviewApprove(ByVal item As Object, ByVal selectedRow As Long) As Boolean
    Dim src As Worksheet, dst As Worksheet, sRow As Long, id As String, no As String, bar As String
    Dim oldText As String, newText As String, backup As String, failureText As String
    Dim i As Long, prevEvents As Boolean, changing As Boolean
    Dim cols(1 To 4) As Long, oldValues(1 To 4) As Variant, oldFormats(1 To 4) As String, j As Long
    Dim oldFormulas(1 To 4) As String, wasFormula(1 To 4) As Boolean
    On Error GoTo Failed
    If CStr(item("category")) <> "B" Then Err.Raise vbObjectError + 6110, , "ID競合や複数行の統合は自動承認できません。手動で確認してください。"
    If Not ReviewCurrent(item) Then Err.Raise vbObjectError + 6109, , "候補・行・データが変わりました。再読み込みしてください。"
    Set src = ThisWorkbook.Worksheets(CSV_SHEET): Set dst = ThisWorkbook.Worksheets(DEST_SHEET)
    sRow = CLng(item("sourceRow"))
    If CStr(item("rows")) <> CStr(selectedRow) Then Err.Raise vbObjectError + 6111, , "候補が複数あります。既存行の統合は行いません。"
    If ReviewFiscal(src, sRow) <> CLng(item("year")) Or ReviewFiscal(dst, selectedRow) <> CLng(item("year")) Then _
        Err.Raise vbObjectError + 6112, , "年度が一致しません。"
    If ReviewField(dst, selectedRow, "同期状態") <> "登録予定" Or Len(ReviewField(dst, selectedRow, "ファイルID")) > 0 Then _
        Err.Raise vbObjectError + 6113, , "登録予定かつID空欄の一意な行に限ります。"
    If Len(ReviewName(src, sRow, True, False)) = 0 Or ReviewName(src, sRow, True, False) <> ReviewName(dst, selectedRow, False, False) Then _
        Err.Raise vbObjectError + 6114, , "作成年度・ガイドを一意に照合できません。"
    id = ReviewField(src, sRow, "ファイルID")
    no = ReviewField(src, sRow, "ファイル番号")
    bar = ReviewField(src, sRow, "バーコード番号")
    If Len(id) = 0 Or Len(no) = 0 Or Len(bar) = 0 Then Err.Raise vbObjectError + 6115, , "3つの識別番号が揃っていません。"
    For i = 2 To ReviewLast(src)
        If i <> sRow Then
            If ReviewField(src, i, "ファイルID") = id Or ReviewField(src, i, "バーコード番号") = bar Or _
               ReviewPadding(ReviewField(src, i, "ファイル番号")) = ReviewPadding(no) Then _
               Err.Raise vbObjectError + 6116, , "CSV側の識別番号が一意ではありません。"
        End If
    Next i
    For i = 2 To ReviewLast(dst)
        If i <> selectedRow Then
            If ReviewField(dst, i, "ファイルID") = id Or ReviewField(dst, i, "バーコード番号") = bar Or _
               ReviewPadding(ReviewField(dst, i, "ファイル番号")) = ReviewPadding(no) Then _
               Err.Raise vbObjectError + 6117, , "他のExcel行と識別番号が重なります。統合せず保留してください。"
        End If
    Next i
    If Len(ReviewField(dst, selectedRow, "ファイル番号")) > 0 And _
       ReviewPadding(ReviewField(dst, selectedRow, "ファイル番号")) <> ReviewPadding(no) Then _
       Err.Raise vbObjectError + 6118, , "登録予定行のファイル番号が異なります。"
    If Len(ReviewField(dst, selectedRow, "バーコード番号")) > 0 And _
       ReviewField(dst, selectedRow, "バーコード番号") <> bar Then _
       Err.Raise vbObjectError + 6119, , "登録予定行のバーコードが異なります。"
    oldText = ReviewDescription(dst, selectedRow, False)
    newText = "Excel行" & selectedRow & " の正式ID [空欄]→" & id & _
        "、ファイル番号 [" & ReviewField(dst, selectedRow, "ファイル番号") & "]→" & no & _
        "、バーコード [" & ReviewField(dst, selectedRow, "バーコード番号") & "]→" & bar & _
        "、同期状態 [登録予定]→[同期済]。ガイド・フォルダー・Excel独自列は変更しません。"
    If MsgBox(oldText & vbCrLf & ReviewDescription(src, sRow, True) & vbCrLf & vbCrLf & newText & vbCrLf & _
        "実行しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "最終確認") <> vbYes Then Exit Function
    If Not ReviewCurrent(item) Then Err.Raise vbObjectError + 6120, , "確認中にCSVまたはExcelの内容が変わりました。"
    If ThisWorkbook.ReadOnly Or dst.ProtectContents Then Err.Raise vbObjectError + 6121, , "ブックが読み取り専用またはシートが保護されています。"
    cols(1) = ReviewCol(dst, "ファイルID"): cols(2) = ReviewCol(dst, "ファイル番号")
    cols(3) = ReviewCol(dst, "バーコード番号"): cols(4) = ReviewCol(dst, "同期状態")
    For j = 1 To 4
        oldValues(j) = dst.Cells(selectedRow, cols(j)).Value2
        wasFormula(j) = dst.Cells(selectedRow, cols(j)).HasFormula
        If wasFormula(j) Then oldFormulas(j) = dst.Cells(selectedRow, cols(j)).Formula
        oldFormats(j) = dst.Cells(selectedRow, cols(j)).NumberFormat
    Next j
    backup = ReviewBackup()
    If Not ReviewCurrent(item) Then Err.Raise vbObjectError + 6137, , "バックアップ後に候補が変わりました。"
    prevEvents = Application.EnableEvents: Application.EnableEvents = False: changing = True
    For j = 1 To 3: dst.Cells(selectedRow, cols(j)).NumberFormat = "@": Next j
    dst.Cells(selectedRow, cols(1)).Value2 = id
    dst.Cells(selectedRow, cols(2)).Value2 = no
    dst.Cells(selectedRow, cols(3)).Value2 = bar
    dst.Cells(selectedRow, cols(4)).Value2 = "同期済"
    ReviewWriteAudit item, "同一として処理", newText, backup
    changing = False: Application.EnableEvents = prevEvents
    SyncReviewApprove = True
    Exit Function
Failed:
    failureText = Err.Description
    If changing Then
        On Error Resume Next
        For j = 1 To 4
            dst.Cells(selectedRow, cols(j)).NumberFormat = oldFormats(j)
            If wasFormula(j) Then
                dst.Cells(selectedRow, cols(j)).Formula = oldFormulas(j)
            Else
                dst.Cells(selectedRow, cols(j)).Value2 = oldValues(j)
            End If
        Next j
        Application.EnableEvents = prevEvents
        On Error GoTo 0
    End If
    MsgBox "承認を実行できません: " & failureText & vbCrLf & "バックアップ: " & backup & vbCrLf & _
        "編集済みなら保存前に対象行と履歴を確認してください。", vbCritical
End Function

Private Function ReviewBackup() As String
    Dim original As String, base As String, ext As String, target As String, suffix As Long, dotPos As Long
    original = ThisWorkbook.FullName
    If Len(ThisWorkbook.Path) = 0 Or ThisWorkbook.ReadOnly Then _
        Err.Raise vbObjectError + 6122, , "保存済みの編集可能なブックで実行してください。"
    dotPos = InStrRev(original, ".")
    If dotPos = 0 Then Err.Raise vbObjectError + 6123, , "バックアップの拡張子を判断できません。"
    base = Left$(original, dotPos - 1): ext = Mid$(original, dotPos)
    target = base & "_同期候補削除前_" & Format$(Now, "yyyymmdd_hhnnss") & ext
    Do While Len(Dir$(target)) > 0
        suffix = suffix + 1
        target = base & "_同期候補削除前_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & suffix & ext
    Loop
    ThisWorkbook.SaveCopyAs target
    If Len(Dir$(target)) = 0 Then Err.Raise vbObjectError + 6124, , "バックアップを作成できませんでした。"
    ReviewBackup = target
End Function

Public Function SyncReviewDeleteBatch(ByVal selected As Collection) As Boolean
    Dim dst As Worksheet, src As Worksheet, item As Object, ordered As New Collection
    Dim i As Long, j As Long, bestRow As Long, bestIndex As Long, rowNo As Long
    Dim selectedKeys As Object, summary As String, backup As String, prevEvents As Boolean, changing As Boolean
    Dim completed As Long, failureText As String, current As Collection, approvedKeys As Object
    Dim rowDeleted As Boolean
    On Error GoTo Failed
    If selected Is Nothing Then Err.Raise vbObjectError + 6125, , "削除候補を選択してください。"
    If selected.Count = 0 Then Err.Raise vbObjectError + 6125, , "削除候補を選択してください。"
    Set dst = ThisWorkbook.Worksheets(DEST_SHEET): Set src = ThisWorkbook.Worksheets(CSV_SHEET)
    If ThisWorkbook.ReadOnly Or dst.ProtectContents Or dst.FilterMode Then _
        Err.Raise vbObjectError + 6126, , "読み取り専用・シート保護・フィルター適用中は削除しません。"
    Set selectedKeys = CreateObject("Scripting.Dictionary")
    Set current = SyncReviewCandidates(CLng(selected(1)("year")))
    For Each item In selected
        If CStr(item("category")) <> "C" Or CLng(item("sourceRow")) <> 0 Then _
            Err.Raise vbObjectError + 6127, , "削除可以外の候補が混ざっています。"
        rowNo = CLng(item("rows"))
        If selectedKeys.Exists(CStr(rowNo)) Then Err.Raise vbObjectError + 6128, , "同じ行を二重に選択しています。"
        selectedKeys.Add CStr(rowNo), True
        If Len(SyncReviewDecision(item)) > 0 Then Err.Raise vbObjectError + 6138, , "保留・別物と判断済みの候補は削除できません。"
        If Not ReviewCurrent(item, current) Then Err.Raise vbObjectError + 6129, , "CSV・対象行・候補が変更されました。再読み込みしてください。"
        summary = summary & vbCrLf & ReviewDescription(dst, rowNo, False)
    Next item
    If MsgBox("次の " & selected.Count & " 行を完全に削除します。取り消しは事前バックアップから復旧してください。" & _
        vbCrLf & summary & vbCrLf & vbCrLf & "本当に一括削除しますか？", _
        vbYesNo + vbExclamation + vbDefaultButton2, "最終確認：一括削除") <> vbYes Then Exit Function
    Set current = SyncReviewCandidates(CLng(selected(1)("year")))
    For Each item In selected
        If Len(SyncReviewDecision(item)) > 0 Or Not ReviewCurrent(item, current) Then _
            Err.Raise vbObjectError + 6130, , "最終確認中に判断または内容が変わりました。"
    Next item
    backup = ReviewBackup()
    Set approvedKeys = CreateObject("Scripting.Dictionary")
    For Each item In selected
        ReviewWriteAudit item, "削除承認", "最終確認で選択・同意済み。データは未変更。", backup
        approvedKeys.Add CStr(item("key")), True
    Next item
    For i = selected.Count To 1 Step -1
        bestRow = 0: bestIndex = 0
        For j = 1 To selected.Count
            Set item = selected(j)
            rowNo = CLng(item("rows"))
            If Not ReviewCollectionHasRow(ordered, rowNo) And rowNo > bestRow Then
                bestRow = rowNo: bestIndex = j
            End If
        Next j
        If bestIndex > 0 Then ordered.Add selected(bestIndex)
    Next i
    prevEvents = Application.EnableEvents: Application.EnableEvents = False: changing = True
    For Each item In ordered
        rowDeleted = False
        rowNo = CLng(item("rows"))
        If Not approvedKeys.Exists(CStr(item("key"))) Or Len(SyncReviewDecision(item)) > 0 Then _
            Err.Raise vbObjectError + 6131, , "未承認または保留済みの候補です。残りは停止します。"
        If ReviewSignature(dst, rowNo) <> CStr(item("destinationSignature")) Then _
            Err.Raise vbObjectError + 6131, , "対象行の内容が変わりました。残りは停止します。"
        ReviewWriteAudit item, "削除着手", ReviewDescription(dst, rowNo, False), backup
        dst.Rows(rowNo).Delete
        rowDeleted = True
        completed = completed + 1
        ReviewWriteAudit item, "削除完了", "Excel行" & rowNo & "を削除。復元はバックアップから", backup
        rowDeleted = False
    Next item
    changing = False: Application.EnableEvents = prevEvents
    MsgBox completed & " 行を削除しました。履歴シートと事前バックアップを確認して保存してください。" & vbCrLf & backup, vbInformation
    SyncReviewDeleteBatch = True
    Exit Function
Failed:
    failureText = Err.Description
    If changing Then Application.EnableEvents = prevEvents
    If Len(backup) > 0 Then
        On Error Resume Next
        Err.Clear
        If rowDeleted Then
            ReviewWriteAudit item, "削除完了", "Excel行" & rowNo & "は削除済み。最初の完了履歴追記に失敗: " & failureText, backup
        Else
            ReviewWriteAudit item, "削除失敗", failureText & " / 実削除件数 " & completed, backup
        End If
        If Err.Number <> 0 Then failureText = failureText & " / 監査履歴への再記録にも失敗: " & Err.Description
        On Error GoTo 0
    End If
    MsgBox "一括削除を停止しました（物理削除済み " & completed & " 件）。" & failureText & vbCrLf & _
        "バックアップ: " & backup & vbCrLf & "続行せず、履歴と対象行を確認してください。", vbCritical
End Function

Private Function ReviewCollectionHasRow(ByVal rows As Collection, ByVal rowNo As Long) As Boolean
    Dim item As Object
    For Each item In rows
        If CLng(item("rows")) = rowNo Then ReviewCollectionHasRow = True: Exit Function
    Next item
End Function

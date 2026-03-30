Option Explicit

'========================================================
' frmSeal 差し替え用（新基準表対応）
'--------------------------------------------------------
' 既存の frmSeal に以下の関数を貼り付けて置き換えてください。
' - WriteOneSeal
' - BuildIndexes（必要なら置換）
' - 補助関数群
'
' 前提:
' - GroupBase / PutMergeTopLeft / ClearMergeTarget は既存実装を流用
' - 転記元は「新ファイル基準表」
' - キーは「通し番号」
'========================================================

'--------------------------------------------------------
' 必要見出しを候補付きで解決し、通し番号→行番号インデックスを作成
' ※既存BuildIndexesを置き換える場合に使用
'--------------------------------------------------------
Public Sub BuildIndexes_NewSpec(ByVal wsSrc As Worksheet, ByRef serialIndex As Object, ByRef headerIndex As Object)

    Dim lastCol As Long, lastRow As Long
    Dim c As Long, r As Long
    Dim key As String

    Set serialIndex = CreateObject("Scripting.Dictionary")
    Set headerIndex = CreateObject("Scripting.Dictionary")

    lastCol = LastUsedCol(wsSrc)
    lastRow = LastUsedRow(wsSrc)

    ' 見出し候補（表記ゆれ吸収）
    headerIndex("通し番号") = FindHeaderByCandidates(wsSrc, lastCol, Array("通し番号"))
    headerIndex("タイトル") = FindHeaderByCandidates(wsSrc, lastCol, Array("タイトル"))
    headerIndex("分類名2") = FindHeaderByCandidates(wsSrc, lastCol, Array("分類名２", "分類名2"))
    headerIndex("分類名3") = FindHeaderByCandidates(wsSrc, lastCol, Array("分類名３", "分類名3"))
    headerIndex("年度和暦") = FindHeaderByCandidates(wsSrc, lastCol, Array("年度（和暦）", "年度(和暦)"))
    headerIndex("保存期間") = FindHeaderByCandidates(wsSrc, lastCol, Array("保存期間"))
    headerIndex("キャビネット番号") = FindHeaderByCandidates(wsSrc, lastCol, Array("キャビネット番号"))
    headerIndex("n+1") = FindHeaderByCandidates(wsSrc, lastCol, Array("ｎ＋1", "n+1"))

    If CLng(headerIndex("通し番号")) = 0 Then Exit Sub

    For r = 2 To lastRow
        key = Trim$(CStr(wsSrc.Cells(r, CLng(headerIndex("通し番号"))).Value))
        If Len(key) > 0 Then
            If Not serialIndex.Exists(key) Then
                serialIndex.Add key, r
            End If
        End If
    Next r

End Sub

'--------------------------------------------------------
' WriteOneSeal（新仕様）
' 左上ブロック基準:
'   A2 = 保存期間（継続なら「継」）
'   B2 = タイトル
'   H2 = 分類名2
'   A3 = 年度（和暦の数値部）
'   B3 = 分類名3
'--------------------------------------------------------
Public Sub WriteOneSeal_NewSpec( _
    ByVal wsSrc As Worksheet, _
    ByVal srcRow As Long, _
    ByVal wsSeal As Worksheet, _
    ByVal slot As Long, _
    ByVal colTitle As Long, _
    ByVal colClass2 As Long, _
    ByVal colClass3 As Long, _
    ByVal colWareki As Long, _
    ByVal colSaveTerm As Long)

    Dim baseCol As Long, baseRow As Long
    Dim vTitle As String, vClass2 As String, vClass3 As String
    Dim vWareki As String, vSave As String

    ' 既存のGroupBaseを流用（左6面＋右6面レイアウト）
    GroupBase slot, baseRow, baseCol

    vTitle = GetCellTrim(wsSrc, srcRow, colTitle)
    vClass2 = GetCellTrim(wsSrc, srcRow, colClass2)
    vClass3 = GetCellTrim(wsSrc, srcRow, colClass3)
    vWareki = GetCellTrim(wsSrc, srcRow, colWareki)
    vSave = GetCellTrim(wsSrc, srcRow, colSaveTerm)

    ' 念のため対象セルをクリア（結合セル対応は既存ヘルパーを流用）
    ClearMergeTarget wsSeal.Cells(baseRow, baseCol)       ' A2
    ClearMergeTarget wsSeal.Cells(baseRow, baseCol + 1)   ' B2
    ClearMergeTarget wsSeal.Cells(baseRow, baseCol + 7)   ' H2
    ClearMergeTarget wsSeal.Cells(baseRow + 1, baseCol)   ' A3
    ClearMergeTarget wsSeal.Cells(baseRow + 1, baseCol + 1) ' B3

    ' 新仕様の転記
    PutMergeTopLeft wsSeal.Cells(baseRow, baseCol), SaveTermToKei(vSave)          ' A2
    PutMergeTopLeft wsSeal.Cells(baseRow, baseCol + 1), vTitle                     ' B2
    PutMergeTopLeft wsSeal.Cells(baseRow, baseCol + 7), vClass2                    ' H2
    PutMergeTopLeft wsSeal.Cells(baseRow + 1, baseCol), ExtractWarekiNumber(vWareki) ' A3
    PutMergeTopLeft wsSeal.Cells(baseRow + 1, baseCol + 1), vClass3                ' B3

End Sub

'--------------------------------------------------------
' 年度（和暦）から数値部分のみ抽出
' 例: 令和7年度 / R7 / r07 / 7 -> "7"
'--------------------------------------------------------
Public Function ExtractWarekiNumber(ByVal s As String) As String

    Dim t As String, i As Long, ch As String, buf As String

    t = Trim$(CStr(s))
    If Len(t) = 0 Then
        ExtractWarekiNumber = ""
        Exit Function
    End If

    t = Replace(t, "令和", "")
    t = Replace(t, "年度", "")
    t = Replace(t, "R", "")
    t = Replace(t, "r", "")
    t = Replace(t, " ", "")
    t = Replace(t, "　", "")

    ' 全角数字を半角へ寄せる
    t = StrConv(t, vbNarrow)

    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "#" Then buf = buf & ch
    Next i

    If Len(buf) = 0 Then
        ExtractWarekiNumber = ""
    Else
        ExtractWarekiNumber = CStr(CLng(buf)) ' 07 -> 7
    End If

End Function

'--------------------------------------------------------
' 保存期間が「継続」の時だけ「継」を返す
'--------------------------------------------------------
Public Function SaveTermToKei(ByVal s As String) As String

    If StrComp(Trim$(CStr(s)), "継続", vbBinaryCompare) = 0 Then
        SaveTermToKei = "継"
    Else
        SaveTermToKei = ""
    End If

End Function

'--------------------------------------------------------
' ヘッダー候補から列番号取得（完全一致）
'--------------------------------------------------------
Public Function FindHeaderByCandidates(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal candidates As Variant) As Long

    Dim c As Long, i As Long
    Dim hv As String

    FindHeaderByCandidates = 0
    If lastCol <= 0 Then Exit Function

    For c = 1 To lastCol
        hv = Trim$(CStr(ws.Cells(1, c).Value))
        For i = LBound(candidates) To UBound(candidates)
            If StrComp(hv, Trim$(CStr(candidates(i))), vbBinaryCompare) = 0 Then
                FindHeaderByCandidates = c
                Exit Function
            End If
        Next i
    Next c

End Function

Private Function GetCellTrim(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String

    If c <= 0 Then
        GetCellTrim = ""
    Else
        GetCellTrim = Trim$(CStr(ws.Cells(r, c).Value))
    End If

End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long

    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = lastCell.Row
    End If

End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long

    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        LastUsedCol = 0
    Else
        LastUsedCol = lastCell.Column
    End If

End Function

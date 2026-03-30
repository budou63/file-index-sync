Option Explicit

'========================================================
' frmSeal 新仕様差し替えコード（既存構造維持版）
'--------------------------------------------------------
' 目的:
' - 既存 GroupBase / ClearMergeTarget / PutMergeTopLeft の
'   引数仕様に合わせてコンパイルエラーを回避
' - 新基準表見出しで通し番号検索し、12面へ転記
'
' 既存前提（変更しない）:
'   GroupBase(slot, baseCol, baseRow, offSlash, offG1)
'   ClearMergeTarget(ws, row, col)
'   PutMergeTopLeft(ws, row, col, value)
'========================================================

Private Const SHEET_NEW_STANDARD As String = "新ファイル基準表"
Private Const SHEET_SEAL As String = "個別フォルダシール"

'--------------------------------------------------------
' frmSeal の btnApply_Click から呼ぶ想定のエントリ
'--------------------------------------------------------
Public Sub ApplyBySerial_NewSpec(ByVal startSerialText As String)

    Dim wsSrc As Worksheet, wsSeal As Worksheet
    Dim serialIndex As Object, headerIndex As Object
    Dim startNo As Long, slot As Long
    Dim serialKey As String
    Dim srcRow As Long

    On Error GoTo ErrorHandler

    Set wsSrc = ThisWorkbook.Worksheets(SHEET_NEW_STANDARD)
    Set wsSeal = ThisWorkbook.Worksheets(SHEET_SEAL)

    BuildIndexes_NewSpec wsSrc, serialIndex, headerIndex
    If serialIndex Is Nothing Then Exit Sub

    startNo = CLng(Val(Trim$(startSerialText)))
    If startNo <= 0 Then
        MsgBox "通し番号（開始番号）を入力してください。", vbExclamation
        Exit Sub
    End If

    ' 12面へ順番に反映
    For slot = 1 To 12
        serialKey = CStr(startNo + slot - 1)

        If serialIndex.Exists(serialKey) Then
            srcRow = CLng(serialIndex(serialKey))
            WriteOneSeal_NewSpec wsSrc, wsSeal, slot, srcRow, headerIndex
        Else
            ClearOneSeal_NewSpec wsSeal, slot
        End If
    Next slot

    Exit Sub

ErrorHandler:
    MsgBox "シール反映中にエラーが発生しました: " & Err.Description, vbExclamation

End Sub

'--------------------------------------------------------
' 新仕様 WriteOneSeal
' 左上ブロック基準:
'   A2 = 保存期間（継続なら「継」）
'   B2 = タイトル
'   H2 = 分類名2
'   A3 = 年度（和暦の数値部）
'   B3 = 分類名3
'--------------------------------------------------------
Public Sub WriteOneSeal_NewSpec( _
    ByVal wsSrc As Worksheet, _
    ByVal wsSeal As Worksheet, _
    ByVal slot As Long, _
    ByVal srcRow As Long, _
    ByVal headerIndex As Object)

    Dim baseCol As Long, baseRow As Long
    Dim offSlash As Long, offG1 As Long
    Dim vTitle As String, vClass2 As String, vClass3 As String
    Dim vWareki As String, vSave As String
    Dim colTitle As Long, colClass2 As Long, colClass3 As Long
    Dim colWareki As Long, colSave As Long

    colTitle = DictCol(headerIndex, "タイトル")
    colClass2 = DictCol(headerIndex, "分類名2")
    colClass3 = DictCol(headerIndex, "分類名3")
    colWareki = DictCol(headerIndex, "年度和暦")
    colSave = DictCol(headerIndex, "保存期間")

    vTitle = GetCellTrim(wsSrc, srcRow, colTitle)
    vClass2 = GetCellTrim(wsSrc, srcRow, colClass2)
    vClass3 = GetCellTrim(wsSrc, srcRow, colClass3)
    vWareki = GetCellTrim(wsSrc, srcRow, colWareki)
    vSave = GetCellTrim(wsSrc, srcRow, colSave)

    ' 既存関数シグネチャのまま利用（ここがコンパイルエラー対策）
    GroupBase slot, baseCol, baseRow, offSlash, offG1

    ' クリア（既存関数シグネチャに合わせる）
    ClearMergeTarget wsSeal, baseRow, baseCol         ' A2
    ClearMergeTarget wsSeal, baseRow, baseCol + 1     ' B2
    ClearMergeTarget wsSeal, baseRow, baseCol + 7     ' H2
    ClearMergeTarget wsSeal, baseRow + 1, baseCol     ' A3
    ClearMergeTarget wsSeal, baseRow + 1, baseCol + 1 ' B3

    ' 新仕様転記（既存関数シグネチャに合わせる）
    PutMergeTopLeft wsSeal, baseRow, baseCol, SaveTermToKei(vSave)            ' A2
    PutMergeTopLeft wsSeal, baseRow, baseCol + 1, vTitle                       ' B2
    PutMergeTopLeft wsSeal, baseRow, baseCol + 7, vClass2                      ' H2
    PutMergeTopLeft wsSeal, baseRow + 1, baseCol, ExtractWarekiNumber(vWareki) ' A3
    PutMergeTopLeft wsSeal, baseRow + 1, baseCol + 1, vClass3                  ' B3

End Sub

'--------------------------------------------------------
' 対象面を新仕様項目だけクリア
'--------------------------------------------------------
Public Sub ClearOneSeal_NewSpec(ByVal wsSeal As Worksheet, ByVal slot As Long)

    Dim baseCol As Long, baseRow As Long
    Dim offSlash As Long, offG1 As Long

    GroupBase slot, baseCol, baseRow, offSlash, offG1

    ClearMergeTarget wsSeal, baseRow, baseCol
    ClearMergeTarget wsSeal, baseRow, baseCol + 1
    ClearMergeTarget wsSeal, baseRow, baseCol + 7
    ClearMergeTarget wsSeal, baseRow + 1, baseCol
    ClearMergeTarget wsSeal, baseRow + 1, baseCol + 1

End Sub

'--------------------------------------------------------
' 通し番号→行番号インデックス作成
'--------------------------------------------------------
Public Sub BuildIndexes_NewSpec(ByVal wsSrc As Worksheet, ByRef serialIndex As Object, ByRef headerIndex As Object)

    Dim lastCol As Long, lastRow As Long
    Dim r As Long, key As String
    Dim colSerial As Long

    Set serialIndex = CreateObject("Scripting.Dictionary")
    Set headerIndex = CreateObject("Scripting.Dictionary")

    lastCol = LastUsedCol(wsSrc)
    lastRow = LastUsedRow(wsSrc)

    headerIndex("通し番号") = FindHeaderByCandidates(wsSrc, lastCol, Array("通し番号"))
    headerIndex("タイトル") = FindHeaderByCandidates(wsSrc, lastCol, Array("タイトル"))
    headerIndex("分類名2") = FindHeaderByCandidates(wsSrc, lastCol, Array("分類名２", "分類名2"))
    headerIndex("分類名3") = FindHeaderByCandidates(wsSrc, lastCol, Array("分類名３", "分類名3"))
    headerIndex("年度和暦") = FindHeaderByCandidates(wsSrc, lastCol, Array("年度（和暦）", "年度(和暦)"))
    headerIndex("保存期間") = FindHeaderByCandidates(wsSrc, lastCol, Array("保存期間"))

    colSerial = DictCol(headerIndex, "通し番号")
    If colSerial = 0 Then
        MsgBox "見出し「通し番号」が見つかりません。", vbExclamation
        Exit Sub
    End If

    For r = 2 To lastRow
        key = Trim$(CStr(wsSrc.Cells(r, colSerial).Value))
        If Len(key) > 0 Then
            If Not serialIndex.Exists(key) Then
                serialIndex.Add key, r
            End If
        End If
    Next r

End Sub

'--------------------------------------------------------
' 和暦から数値部のみ抽出（令和7年度/R7/7 => 7）
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
    t = StrConv(t, vbNarrow) ' 全角数字対策

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
' 保存期間が継続の時だけ「継」
'--------------------------------------------------------
Public Function SaveTermToKei(ByVal s As String) As String

    If StrComp(Trim$(CStr(s)), "継続", vbBinaryCompare) = 0 Then
        SaveTermToKei = "継"
    Else
        SaveTermToKei = ""
    End If

End Function

'--------------------------------------------------------
' 見出し候補（完全一致）で列取得
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

Private Function DictCol(ByVal dict As Object, ByVal key As String) As Long

    If dict Is Nothing Then
        DictCol = 0
    ElseIf dict.Exists(key) Then
        DictCol = CLng(dict(key))
    Else
        DictCol = 0
    End If

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

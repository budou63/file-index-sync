VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSyncReview
   Caption         =   "同期候補の確認"
   ClientHeight    =   7700
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   11921
   OleObjectBlob   =   "frmSyncReview-native.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmSyncReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLblCounts As MSForms.Label
Private WithEvents mCboCategory As MSForms.ComboBox
Private mLblCase As MSForms.Label
Private WithEvents mLstCases As MSForms.ListBox
Private mLblRows As MSForms.Label
Private WithEvents mLstRows As MSForms.ListBox
Private mLblCompareRows As MSForms.Label
Private mFraCompare As MSForms.Frame
Private mCompareHeaders(0 To 2) As MSForms.Label
Private mCompareItemLabels(1 To 9) As MSForms.TextBox
Private mCompareCsvCells(1 To 9) As MSForms.TextBox
Private mCompareExcelCells(1 To 9) As MSForms.TextBox
Private mTxtDetails As MSForms.TextBox
Private WithEvents mChkDelete As MSForms.CheckBox
Private mLblStatus As MSForms.Label
Private mLblDWarning As MSForms.Label
Private WithEvents mCmdGoTo As MSForms.CommandButton
Private WithEvents mCmdApprove As MSForms.CommandButton
Private WithEvents mCmdOther As MSForms.CommandButton
Private WithEvents mCmdHold As MSForms.CommandButton
Private WithEvents mCmdNext As MSForms.CommandButton
Private WithEvents mCmdDelete As MSForms.CommandButton
Private WithEvents mCmdClear As MSForms.CommandButton
Private WithEvents mCmdRefresh As MSForms.CommandButton
Private WithEvents mCmdFinish As MSForms.CommandButton

Private cases As Collection
Private shown As Collection
Private deletionSelection As Object
Private targetYear As Long
Private rendering As Boolean
Private mDUnavailableReason As String
Private Const COMPARE_FIELD_COUNT As Long = 9
Private Const CSV_SHEET As String = "目録CSV"
Private Const DEST_SHEET As String = "新ファイル基準表"

Private Function EnsureControl(ByVal controlName As String, ByVal progId As String) As MSForms.Control
    Dim existing As MSForms.Control
    Dim expectedType As String
    expectedType = Split(progId, ".")(1)
    For Each existing In Me.Controls
        If existing.Name = controlName Then
            If TypeName(existing) <> expectedType Then Err.Raise vbObjectError + 513, _
                "frmSyncReview", "部品の種類が一致しません: " & controlName
            Set EnsureControl = existing
            Exit Function
        End If
    Next existing
    Set EnsureControl = Me.Controls.Add(progId, controlName, True)
End Function

Private Sub PlaceControl(ByVal control As MSForms.Control, ByVal x As Single, ByVal y As Single, _
                         ByVal w As Single, ByVal h As Single)
    control.Left = x
    control.Top = y
    control.Width = w
    control.Height = h
End Sub

Private Function EnsureFrameControl(ByVal parent As MSForms.Frame, ByVal controlName As String, _
                                    ByVal progId As String) As MSForms.Control
    Dim existing As MSForms.Control
    Dim expectedType As String
    expectedType = Split(progId, ".")(1)
    For Each existing In parent.Controls
        If existing.Name = controlName Then
            If TypeName(existing) <> expectedType Then Err.Raise vbObjectError + 514, _
                "frmSyncReview", "比較表部品の種類が一致しません: " & controlName
            Set EnsureFrameControl = existing
            Exit Function
        End If
    Next existing
    Set EnsureFrameControl = parent.Controls.Add(progId, controlName, True)
End Function

Private Sub BuildControls()
    Dim i As Long
    Me.Caption = "同期候補の確認"
    Me.Width = 760
    Me.Height = Me.Height + (440 - Me.InsideHeight)
    Set mLblCounts = EnsureControl("lblCounts", "Forms.Label.1")
    PlaceControl mLblCounts, 9, 6, 740, 19
    mLblCounts.Caption = "候補を読み込みます"
    Set mCboCategory = EnsureControl("cboCategory", "Forms.ComboBox.1")
    PlaceControl mCboCategory, 9, 28, 187.5, 16.5
    Set mLblCase = EnsureControl("lblCase", "Forms.Label.1")
    PlaceControl mLblCase, 9, 50, 187.5, 14
    mLblCase.Caption = "案件一覧（1件ずつ確認）"
    Set mLstCases = EnsureControl("lstCases", "Forms.ListBox.1")
    PlaceControl mLstCases, 9, 64.5, 187.5, 228
    Set mLblRows = EnsureControl("lblRows", "Forms.Label.1")
    PlaceControl mLblRows, 205, 50, 110, 14
    mLblRows.Caption = "候補のExcel行"
    Set mLstRows = EnsureControl("lstRows", "Forms.ListBox.1")
    PlaceControl mLstRows, 205, 64.5, 110, 228
    Set mLblCompareRows = EnsureControl("lblCompareRows", "Forms.Label.1")
    PlaceControl mLblCompareRows, 324, 28, 420, 18
    mLblCompareRows.Caption = "CSV行: ― / 新ファイル基準表行: ―"
    Set mFraCompare = EnsureControl("fraCompare", "Forms.Frame.1")
    PlaceControl mFraCompare, 324, 48, 420, 225
    mFraCompare.Caption = ""
    mFraCompare.ScrollBars = fmScrollBarsVertical
    mFraCompare.ScrollHeight = 430
    For i = 0 To 2
        Set mCompareHeaders(i) = EnsureFrameControl(mFraCompare, "lblCompareHeader" & CStr(i), "Forms.Label.1")
        Select Case i
            Case 0
                PlaceControl mCompareHeaders(i), 4, 3, 78, 18
                mCompareHeaders(i).Caption = "項目 / 表示比較"
            Case 1
                PlaceControl mCompareHeaders(i), 86, 3, 156, 18
                mCompareHeaders(i).Caption = "目録CSV"
            Case 2
                PlaceControl mCompareHeaders(i), 246, 3, 156, 18
                mCompareHeaders(i).Caption = "新ファイル基準表"
        End Select
        mCompareHeaders(i).Font.Bold = True
    Next i
    For i = 1 To COMPARE_FIELD_COUNT
        Set mCompareItemLabels(i) = EnsureFrameControl(mFraCompare, "txtCompareField" & Format$(i, "00"), "Forms.TextBox.1")
        Set mCompareCsvCells(i) = EnsureFrameControl(mFraCompare, "txtCompareCsv" & Format$(i, "00"), "Forms.TextBox.1")
        Set mCompareExcelCells(i) = EnsureFrameControl(mFraCompare, "txtCompareExcel" & Format$(i, "00"), "Forms.TextBox.1")
        PlaceControl mCompareItemLabels(i), 4, 24, 78, 32
        PlaceControl mCompareCsvCells(i), 86, 24, 156, 32
        PlaceControl mCompareExcelCells(i), 246, 24, 156, 32
        mCompareItemLabels(i).MultiLine = True
        mCompareCsvCells(i).MultiLine = True
        mCompareExcelCells(i).MultiLine = True
        mCompareItemLabels(i).WordWrap = True
        mCompareCsvCells(i).WordWrap = True
        mCompareExcelCells(i).WordWrap = True
        mCompareItemLabels(i).ScrollBars = fmScrollBarsVertical
        mCompareCsvCells(i).ScrollBars = fmScrollBarsVertical
        mCompareExcelCells(i).ScrollBars = fmScrollBarsVertical
        mCompareItemLabels(i).Locked = True
        mCompareCsvCells(i).Locked = True
        mCompareExcelCells(i).Locked = True
        mCompareItemLabels(i).TabStop = False
        mCompareCsvCells(i).TabStop = False
        mCompareExcelCells(i).TabStop = False
    Next i
    Set mTxtDetails = EnsureControl("txtDetails", "Forms.TextBox.1")
    PlaceControl mTxtDetails, 324, 278, 420, 50
    Set mChkDelete = EnsureControl("chkDelete", "Forms.CheckBox.1")
    PlaceControl mChkDelete, 9, 301.5, 300, 17.5
    mChkDelete.Caption = "選択したExcel行を削除対象にする"
    Set mLblStatus = EnsureControl("lblStatus", "Forms.Label.1")
    PlaceControl mLblStatus, 9, 333, 740, 21
    mLblStatus.Caption = "フォームを開いただけでは変更しません。"
    Set mLblDWarning = EnsureControl("lblDWarning", "Forms.Label.1")
    PlaceControl mLblDWarning, 9, 356, 740, 21
    mLblDWarning.Visible = False
    Set mCmdGoTo = EnsureControl("cmdGoTo", "Forms.CommandButton.1")
    PlaceControl mCmdGoTo, 9, 379.5, 70, 21
    mCmdGoTo.Caption = "対象行へ移動"
    Set mCmdApprove = EnsureControl("cmdApprove", "Forms.CommandButton.1")
    PlaceControl mCmdApprove, 84, 379.5, 78, 21
    mCmdApprove.Caption = "同一として処理"
    Set mCmdOther = EnsureControl("cmdOther", "Forms.CommandButton.1")
    PlaceControl mCmdOther, 167, 379.5, 50, 21
    mCmdOther.Caption = "別物"
    Set mCmdHold = EnsureControl("cmdHold", "Forms.CommandButton.1")
    PlaceControl mCmdHold, 222, 379.5, 50, 21
    mCmdHold.Caption = "保留"
    Set mCmdNext = EnsureControl("cmdNext", "Forms.CommandButton.1")
    PlaceControl mCmdNext, 277, 379.5, 60, 21
    mCmdNext.Caption = "次の候補"
    Set mCmdDelete = EnsureControl("cmdDelete", "Forms.CommandButton.1")
    PlaceControl mCmdDelete, 342, 379.5, 85.5, 21
    mCmdDelete.Caption = "選択行を一括削除"
    Set mCmdClear = EnsureControl("cmdClear", "Forms.CommandButton.1")
    PlaceControl mCmdClear, 433, 379.5, 125, 21
    mCmdClear.Caption = "選択行の番号をクリア"
    Set mCmdRefresh = EnsureControl("cmdRefresh", "Forms.CommandButton.1")
    PlaceControl mCmdRefresh, 9, 407, 70, 21
    mCmdRefresh.Caption = "再判定"
    Set mCmdFinish = EnsureControl("cmdFinish", "Forms.CommandButton.1")
    PlaceControl mCmdFinish, 508, 407, 82, 21
    mCmdFinish.Caption = "確認終了"
    mTxtDetails.MultiLine = True
    mTxtDetails.ScrollBars = fmScrollBarsVertical
    mTxtDetails.Locked = True
End Sub

Private Sub UserForm_Initialize()
    rendering = True
    BuildControls
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    mCboCategory.AddItem "A：判定困難・ID競合"
    mCboCategory.AddItem "B：同名・類似候補"
    mCboCategory.AddItem "C：削除可候補"
    mCboCategory.AddItem "D：コピー識別番号のクリア候補"
    mCboCategory.ListIndex = 0
    rendering = False
End Sub

Public Sub LoadCases(ByVal loaded As Collection, ByVal yearValue As Long, _
                     Optional ByVal unavailableReason As String = "")
    Set cases = loaded
    targetYear = yearValue
    mDUnavailableReason = unavailableReason
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    UpdateDWarning
    If Len(mDUnavailableReason) > 0 And mCboCategory.ListIndex = 3 Then
        rendering = True: mCboCategory.ListIndex = 0: rendering = False
    End If
    SelectDWhenOnlyCategory
    ShowCategory
    If Len(mDUnavailableReason) > 0 Then MsgBox "D候補を利用できません: " & mDUnavailableReason, vbExclamation
End Sub

Private Sub UpdateDWarning()
    mLblDWarning.Caption = "警告：D利用不可: " & mDUnavailableReason
    mLblDWarning.ControlTipText = mDUnavailableReason
    mLblDWarning.Visible = (Len(mDUnavailableReason) > 0)
    rendering = True
    If mLblDWarning.Visible Then
        mCboCategory.List(3) = "D：利用不可（必要な見出しなし）"
    Else
        mCboCategory.List(3) = "D：コピー識別番号のクリア候補"
    End If
    rendering = False
End Sub

Private Sub SelectDWhenOnlyCategory()
    If Len(mDUnavailableReason) = 0 And Not HasReviewCategory("A") And Not HasReviewCategory("B") And _
       Not HasReviewCategory("C") And HasReviewCategory("D") Then
        rendering = True: mCboCategory.ListIndex = 3: rendering = False
    End If
End Sub

Private Function HasReviewCategory(ByVal category As String) As Boolean
    Dim value As Object
    If cases Is Nothing Then Exit Function
    For Each value In cases
        If CStr(value("category")) = category Then HasReviewCategory = True: Exit Function
    Next value
End Function

Private Function SelectedCase() As Object
    If shown Is Nothing Then Exit Function
    If mLstCases.ListIndex < 0 Or mLstCases.ListIndex + 1 > shown.Count Then Exit Function
    Set SelectedCase = shown(mLstCases.ListIndex + 1)
End Function

Private Function SelectedRow() As Long
    If mLstRows.ListIndex < 0 Then Exit Function
    SelectedRow = CLng(Split(CStr(mLstRows.List(mLstRows.ListIndex)), " ")(0))
End Function

Private Function BNextUnconfirmedRow(ByVal item As Object, ByVal afterRow As Long) As Long
    Dim parts As Variant, i As Long, rowNo As Long, count As Long, currentIndex As Long
    Dim foundCurrent As Boolean
    If item Is Nothing Then Exit Function
    If CStr(item("category")) <> "B" Or Len(CStr(item("rows"))) = 0 Then Exit Function
    parts = Split(CStr(item("rows")), ";")
    count = UBound(parts) - LBound(parts) + 1
    If count <= 0 Then Exit Function
    If afterRow >= 2 Then
        For i = LBound(parts) To UBound(parts)
            If CLng(parts(i)) = afterRow Then
                currentIndex = i
                foundCurrent = True
                Exit For
            End If
        Next i
    End If
    If Not foundCurrent Then
        For i = LBound(parts) To UBound(parts)
            rowNo = CLng(parts(i))
            If Len(SyncReviewDecision(item, rowNo)) = 0 Then
                BNextUnconfirmedRow = rowNo
                Exit Function
            End If
        Next i
        Exit Function
    End If
    For i = 1 To count - 1
        rowNo = CLng(parts((currentIndex + i) Mod count))
        If Len(SyncReviewDecision(item, rowNo)) = 0 Then
            BNextUnconfirmedRow = rowNo
            Exit Function
        End If
    Next i
End Function

Private Sub UpdateBRowJudgments(ByVal item As Object)
    Dim parts As Variant, i As Long, rowNo As Long, totalCount As Long, confirmedCount As Long
    Dim rowDecision As String
    If item Is Nothing Then Exit Sub
    If CStr(item("category")) <> "B" Or Len(CStr(item("rows"))) = 0 Then Exit Sub
    parts = Split(CStr(item("rows")), ";")
    totalCount = UBound(parts) - LBound(parts) + 1
    mLblRows.Caption = "候補Excel行／判断"
    For i = LBound(parts) To UBound(parts)
        rowNo = CLng(parts(i))
        rowDecision = SyncReviewDecision(item, rowNo)
        If rowDecision = "別物" Then
            confirmedCount = confirmedCount + 1
        Else
            rowDecision = "未確認"
        End If
        If i < mLstRows.ListCount Then mLstRows.List(i) = CStr(rowNo) & " 行 [" & rowDecision & "]"
    Next i
    mLblStatus.Caption = "CSV" & CStr(item("sourceRow")) & "行：候補" & totalCount & "行／別物確認済" & _
        confirmedCount & "行／未確認" & (totalCount - confirmedCount) & "行"
    If SyncReviewDecision(item) = "保留" Then mLblStatus.Caption = mLblStatus.Caption & "／案件保留"
    If SyncReviewLegacyCaseDecision(item) = "別物" Then _
        mLblStatus.Caption = mLblStatus.Caption & "／旧方式判断あり"
End Sub

Private Sub CompareFieldSpec(ByVal fieldIndex As Long, ByRef fieldCaption As String, _
                              ByRef csvAliases As Variant, ByRef excelAliases As Variant)
    Select Case fieldIndex
        Case 1
            fieldCaption = "ファイルID": csvAliases = Array("ファイルID"): excelAliases = Array("ファイルID")
        Case 2
            fieldCaption = "ファイル番号": csvAliases = Array("ファイル番号"): excelAliases = Array("ファイル番号")
        Case 3
            fieldCaption = "バーコード番号": csvAliases = Array("バーコード番号"): excelAliases = Array("バーコード番号")
        Case 4
            fieldCaption = "年度": csvAliases = Array("年度"): excelAliases = Array("年度")
        Case 5
            fieldCaption = "作成年度": csvAliases = Array("作成年度（和暦）", "作成年度(和暦)")
            excelAliases = Array("作成年度（和暦）", "作成年度(和暦)")
        Case 6
            fieldCaption = "第1ガイド": csvAliases = Array("分類名２", "分類名2", "第1ガイド", "第１ガイド")
            excelAliases = Array("第1ガイド", "第１ガイド", "分類名２", "分類名2")
        Case 7
            fieldCaption = "第2ガイド": csvAliases = Array("分類名３", "分類名3", "第2ガイド", "第２ガイド")
            excelAliases = Array("第2ガイド", "第２ガイド", "分類名３", "分類名3")
        Case 8
            fieldCaption = "個別フォルダー": csvAliases = Array("タイトル", "個別フォルダー", "個別フォルダ")
            excelAliases = Array("個別フォルダー", "個別フォルダ", "タイトル")
        Case 9
            fieldCaption = "同期状態": csvAliases = Empty: excelAliases = Array("同期状態")
    End Select
End Sub

Private Function CompareHeaderKey(ByVal rawHeader As String) As String
    rawHeader = Replace(Trim$(rawHeader), "　", "")
    rawHeader = Replace(rawHeader, " ", "")
    On Error Resume Next
    rawHeader = StrConv(rawHeader, vbNarrow)
    On Error GoTo 0
    CompareHeaderKey = LCase$(rawHeader)
End Function

Private Function FindCompareColumn(ByVal ws As Worksheet, ByVal aliases As Variant) As Long
    Dim lastCol As Long, colNo As Long, foundCol As Long, foundCount As Long
    Dim headerText As String, alias As Variant
    If ws Is Nothing Or IsEmpty(aliases) Then Exit Function
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For colNo = 1 To lastCol
        If Not IsError(ws.Cells(1, colNo).Value2) Then
            headerText = CompareHeaderKey(CStr(ws.Cells(1, colNo).Value2))
            For Each alias In aliases
                If headerText = CompareHeaderKey(CStr(alias)) Then
                    foundCount = foundCount + 1
                    foundCol = colNo
                    Exit For
                End If
            Next alias
        End If
    Next colNo
    If foundCount = 1 Then FindCompareColumn = foundCol
End Function

Private Function CompareCellText(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal aliases As Variant, _
                                 ByRef headerFound As Boolean) As String
    Dim colNo As Long, rawValue As Variant, shownText As String
    headerFound = False
    If ws Is Nothing Or rowNo < 2 Then Exit Function
    colNo = FindCompareColumn(ws, aliases)
    If colNo <= 0 Then Exit Function
    headerFound = True
    rawValue = ws.Cells(rowNo, colNo).Value2
    If IsError(rawValue) Then
        CompareCellText = "セルエラー"
        Exit Function
    End If
    If IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    CompareCellText = CStr(rawValue)
    shownText = CStr(ws.Cells(rowNo, colNo).Text)
    If Len(shownText) > 0 And InStr(1, shownText, "#", vbBinaryCompare) = 0 Then _
        CompareCellText = shownText
End Function

Private Function CompareDisplayStatus(ByVal csvHeaderFound As Boolean, ByVal excelHeaderFound As Boolean, _
                                      ByVal csvText As String, ByVal excelText As String) As String
    If Not csvHeaderFound Or Not excelHeaderFound Then
        CompareDisplayStatus = "比較対象外"
    ElseIf StrComp(csvText, excelText, vbBinaryCompare) = 0 Then
        CompareDisplayStatus = "一致"
    Else
        CompareDisplayStatus = "相違"
    End If
End Function

Private Function CompareRowHeight(ByVal fieldCaption As String, ByVal csvText As String, _
                                  ByVal excelText As String, ByVal statusText As String) As Single
    Dim longest As Long, heightValue As Single
    longest = Application.Max(Len(fieldCaption) + Len(statusText), Len(csvText), Len(excelText))
    heightValue = 30 + (longest \ 16) * 12
    If heightValue > 66 Then heightValue = 66
    CompareRowHeight = heightValue
End Function

Private Function CompareRowLabel(ByVal rowNo As Long) As String
    If rowNo < 2 Then
        CompareRowLabel = "―"
    Else
        CompareRowLabel = CStr(rowNo)
    End If
End Function

Private Sub RenderComparison(ByVal item As Object, ByVal excelRow As Long)
    Dim src As Worksheet, dst As Worksheet
    Dim sourceRow As Long, fieldIndex As Long, y As Single, rowHeight As Single
    Dim fieldCaption As String, statusText As String, csvText As String, excelText As String
    Dim csvAliases As Variant, excelAliases As Variant
    Dim csvHeaderFound As Boolean, excelHeaderFound As Boolean
    Dim previousDecision As String, decisionComment As String, fillColor As Long
    Dim selectedDecision As String, selectedComment As String, legacyDecision As String
    If item Is Nothing Then
        mLblCompareRows.Caption = "CSV行: ― / 新ファイル基準表行: ―"
        mTxtDetails.Text = ""
        mFraCompare.ScrollTop = 0
        y = 24
        For fieldIndex = 1 To COMPARE_FIELD_COUNT
            mCompareItemLabels(fieldIndex).Text = "―" & vbCrLf & "比較対象外"
            mCompareCsvCells(fieldIndex).Text = "―"
            mCompareExcelCells(fieldIndex).Text = "―"
            mCompareItemLabels(fieldIndex).BackColor = RGB(242, 242, 242)
            mCompareCsvCells(fieldIndex).BackColor = RGB(242, 242, 242)
            mCompareExcelCells(fieldIndex).BackColor = RGB(242, 242, 242)
            PlaceControl mCompareItemLabels(fieldIndex), 4, y, 78, 30
            PlaceControl mCompareCsvCells(fieldIndex), 86, y, 156, 30
            PlaceControl mCompareExcelCells(fieldIndex), 246, y, 156, 30
            y = y + 32
        Next fieldIndex
        mFraCompare.ScrollHeight = y + 4
        Exit Sub
    End If
    Set src = ThisWorkbook.Worksheets(CSV_SHEET)
    Set dst = ThisWorkbook.Worksheets(DEST_SHEET)
    sourceRow = CLng(item("sourceRow"))
    mLblCompareRows.Caption = "CSV行: " & CompareRowLabel(sourceRow) & _
        "    新ファイル基準表行: " & CompareRowLabel(excelRow)
    y = 24
    mFraCompare.ScrollTop = 0
    For fieldIndex = 1 To COMPARE_FIELD_COUNT
        csvAliases = Empty: excelAliases = Empty: fieldCaption = ""
        CompareFieldSpec fieldIndex, fieldCaption, csvAliases, excelAliases
        csvHeaderFound = False: excelHeaderFound = False
        csvText = "": excelText = ""
        If sourceRow >= 2 Then csvText = CompareCellText(src, sourceRow, csvAliases, csvHeaderFound)
        If excelRow >= 2 Then excelText = CompareCellText(dst, excelRow, excelAliases, excelHeaderFound)
        statusText = CompareDisplayStatus(csvHeaderFound, excelHeaderFound, csvText, excelText)
        If Not csvHeaderFound Or sourceRow < 2 Or Len(csvText) = 0 Then csvText = "―"
        If Not excelHeaderFound Or excelRow < 2 Or Len(excelText) = 0 Then excelText = "―"
        mCompareItemLabels(fieldIndex).Text = fieldCaption & vbCrLf & statusText
        mCompareCsvCells(fieldIndex).Text = csvText
        mCompareExcelCells(fieldIndex).Text = excelText
        Select Case statusText
            Case "相違": fillColor = RGB(255, 235, 235)
            Case "比較対象外": fillColor = RGB(242, 242, 242)
            Case Else: fillColor = RGB(255, 255, 255)
        End Select
        mCompareItemLabels(fieldIndex).BackColor = fillColor
        mCompareCsvCells(fieldIndex).BackColor = fillColor
        mCompareExcelCells(fieldIndex).BackColor = fillColor
        mCompareItemLabels(fieldIndex).ControlTipText = fieldCaption & "：" & statusText
        rowHeight = CompareRowHeight(fieldCaption, csvText, excelText, statusText)
        PlaceControl mCompareItemLabels(fieldIndex), 4, y, 78, rowHeight
        PlaceControl mCompareCsvCells(fieldIndex), 86, y, 156, rowHeight
        PlaceControl mCompareExcelCells(fieldIndex), 246, y, 156, rowHeight
        y = y + rowHeight + 2
    Next fieldIndex
    mFraCompare.ScrollHeight = y + 4
    previousDecision = ""
    decisionComment = ""
    If CStr(item("category")) <> "D" Then
        previousDecision = SyncReviewDecision(item)
        decisionComment = SyncReviewDecisionComment(item)
    End If
    If Len(previousDecision) = 0 Then previousDecision = "未確認"
    If CStr(item("category")) = "B" Then
        selectedDecision = SyncReviewDecision(item, excelRow)
        selectedComment = SyncReviewDecisionComment(item, excelRow)
        legacyDecision = SyncReviewLegacyCaseDecision(item)
        If Len(selectedDecision) = 0 Then selectedDecision = "未確認"
        mTxtDetails.Text = "【保留理由】" & CStr(item("reason")) & vbCrLf & _
            "【案件の判断】" & previousDecision & IIf(previousDecision = "保留" And Len(decisionComment) > 0, " / " & decisionComment, "") & vbCrLf & _
            "【選択Excel行の判断】" & selectedDecision & IIf(Len(selectedComment) > 0, " / " & selectedComment, "")
        If legacyDecision = "別物" Then
            mTxtDetails.Text = mTxtDetails.Text & vbCrLf & _
                "【旧方式の案件単位判断】別物（候補行ごとの判断としては扱いません）" & _
                IIf(Len(decisionComment) > 0, " / " & decisionComment, "")
        End If
    Else
        If Len(decisionComment) > 0 Then previousDecision = previousDecision & " / " & decisionComment
        mTxtDetails.Text = "【保留理由】" & CStr(item("reason")) & vbCrLf & _
            "【前回判断】" & previousDecision
    End If
End Sub

Private Sub ShowCategory()
    Dim i As Long, item As Object, category As String, counts(1 To 4) As Long, unresolved(1 To 4) As Long
    Dim decision As String, idx As Long, categoryIndex As Long, isD As Boolean
    If cases Is Nothing Then Exit Sub
    rendering = True
    Set shown = New Collection
    mLstCases.Clear: mLstRows.Clear
    categoryIndex = mCboCategory.ListIndex
    For i = 1 To cases.Count
        Set item = cases(i)
        category = CStr(item("category"))
        Select Case category
            Case "A": idx = 1
            Case "B": idx = 2
            Case "C": idx = 3
            Case "D": idx = 4
            Case Else: GoTo NextCase
        End Select
        counts(idx) = counts(idx) + 1
        isD = (category = "D")
        decision = ""
        If isD Then
            unresolved(idx) = unresolved(idx) + item("copyCandidateRows").Count
        Else
            decision = SyncReviewDecision(item)
            If Len(decision) = 0 Then unresolved(idx) = unresolved(idx) + 1
        End If
        If categoryIndex = idx - 1 Then
            shown.Add item
            If isD Then
                mLstCases.AddItem "CSV行" & CStr(item("sourceRow")) & " / 保持Excel " & _
                    CStr(item("copyKeeperRow")) & " / クリア候補 " & CStr(item("copyCandidateRows").Count) & "件"
            Else
                mLstCases.AddItem "CSV行" & CStr(item("sourceRow")) & " / Excel " & CStr(item("rows")) & _
                    IIf(Len(decision) = 0, " [未確認]", " [" & decision & "]")
            End If
        End If
NextCase:
    Next i
    mLblCounts.Caption = "対象年度 " & targetYear & " / A " & counts(1) & "件（未 " & unresolved(1) & "）" & _
        " / B " & counts(2) & "件（未 " & unresolved(2) & "）" & _
        " / C " & counts(3) & "件（未 " & unresolved(3) & "）" & _
        " / D " & counts(4) & "件（候補 " & unresolved(4) & "行）"
    mTxtDetails.Text = "候補を選んでください。"
    mChkDelete.Value = False
    If mLstCases.ListCount > 0 Then mLstCases.ListIndex = 0
    rendering = False
    ShowCurrent
End Sub

Private Sub ShowCurrent()
    Dim item As Object, v As Variant, rows As String, category As String
    Dim candidates As Collection, candidate As Variant, keeper As Long, rowNo As Long
    Dim decision As String, rowDecision As String
    Set item = SelectedCase()
    rendering = True
    mLstRows.Clear
    mChkDelete.Value = False
    If item Is Nothing Then
        mTxtDetails.Text = "このカテゴリには候補がありません。"
        RenderComparison Nothing, 0
        mCmdApprove.Enabled = False: mCmdGoTo.Enabled = False: mCmdClear.Enabled = False
        mCmdOther.Enabled = False: mCmdHold.Enabled = False: mCmdDelete.Enabled = False
        mChkDelete.Visible = False
        rendering = False
        Exit Sub
    End If
    category = CStr(item("category"))
    decision = ""
    If category <> "D" Then decision = SyncReviewDecision(item)
    If category = "C" Then
        mLstRows.MultiSelect = fmMultiSelectMulti
    Else
        mLstRows.MultiSelect = fmMultiSelectSingle
    End If
    keeper = 0
    If category = "D" Then
        keeper = CLng(item("copyKeeperRow"))
        Set candidates = item("copyCandidateRows")
        mLstRows.AddItem CStr(keeper) & " 行 - 番号保持行"
        For Each candidate In candidates
            mLstRows.AddItem CStr(CLng(candidate)) & " 行 - クリア候補"
        Next candidate
    Else
        rows = CStr(item("rows"))
        If Len(rows) > 0 Then
            For Each v In Split(rows, ";")
                If category = "B" Then
                    rowDecision = SyncReviewDecision(item, CLng(v))
                    If Len(rowDecision) = 0 Then rowDecision = "未確認"
                    mLstRows.AddItem CStr(v) & " 行 [" & rowDecision & "]"
                Else
                    mLstRows.AddItem CStr(v) & " 行"
                End If
            Next v
        End If
    End If
    If mLstRows.ListCount > 0 Then mLstRows.ListIndex = 0
    mTxtDetails.Text = ""
    rowNo = SelectedRow()
    mCmdGoTo.Enabled = (category <> "D" And rowNo >= 2)
    mCmdApprove.Enabled = (category = "B" And mLstRows.ListCount = 1)
    mCmdOther.Enabled = (category <> "D")
    mCmdHold.Enabled = (category <> "D")
    mCmdDelete.Enabled = (category <> "D")
    mCmdClear.Enabled = IsCopyCandidate(item, rowNo)
    If Len(mDUnavailableReason) > 0 Then mCmdClear.Enabled = False
    mChkDelete.Visible = (category = "C")
    If mChkDelete.Visible Then mChkDelete.Value = (Len(decision) = 0 And deletionSelection.Exists(CStr(item("key"))))
    If category = "D" Then
        mLblStatus.Caption = "D候補：保持行はクリアできません。選択した候補行のみ専用APIで再検証します。"
    ElseIf category = "B" Then
        UpdateBRowJudgments item
    Else
        mLblRows.Caption = "候補のExcel行"
        mLblStatus.Caption = "Aの重複ID・複数行の統合は自動承認不可。Cはチェック後に件数を確認します。"
    End If
    rendering = False
    RenderComparison item, rowNo
End Sub

Private Function IsCopyCandidate(ByVal item As Object, ByVal rowNo As Long) As Boolean
    Dim candidate As Variant
    If item Is Nothing Then Exit Function
    If CStr(item("category")) <> "D" Or rowNo < 2 Then Exit Function
    If rowNo = CLng(item("copyKeeperRow")) Then Exit Function
    For Each candidate In item("copyCandidateRows")
        If CLng(candidate) = rowNo Then IsCopyCandidate = True: Exit Function
    Next candidate
End Function

Private Sub mCboCategory_Change()
    If rendering Then Exit Sub
    If Len(mDUnavailableReason) > 0 Then
        If mCboCategory.ListIndex = 3 Then
            rendering = True: mCboCategory.ListIndex = 0: rendering = False
        End If
    End If
    ShowCategory
End Sub
Private Sub mLstCases_Click()
    If Not rendering Then ShowCurrent
End Sub
Private Sub mLstRows_Click()
    Dim item As Object
    If rendering Then Exit Sub
    Set item = SelectedCase()
    If Not item Is Nothing Then
        If CStr(item("category")) = "C" Then SelectDeleteRow item, SelectedRow()
    End If
    ShowCurrentSelection
End Sub

Private Sub ShowCurrentSelection()
    Dim item As Object, rowNo As Long
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    rowNo = SelectedRow()
    If CStr(item("category")) = "D" Then
        mCmdClear.Enabled = IsCopyCandidate(item, rowNo)
        If Len(mDUnavailableReason) > 0 Then mCmdClear.Enabled = False
        If rowNo = CLng(item("copyKeeperRow")) Then
            mLblStatus.Caption = "この行はCSVと一致する番号保持行です。クリアできません。"
        ElseIf mCmdClear.Enabled Then
            mLblStatus.Caption = "D候補行です。ボタンで専用APIの再照合・確認を開始します。"
        Else
            mLblStatus.Caption = "D候補ではありません。クリアできません。"
        End If
    End If
    RenderComparison item, rowNo
End Sub

Private Sub mCmdClear_Click()
    Dim item As Object, completed As Boolean, apiInvoked As Boolean
    On Error GoTo Failed
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If Len(mDUnavailableReason) > 0 Then Exit Sub
    If Not IsCopyCandidate(item, SelectedRow()) Then
        mLblStatus.Caption = "保持行または候補外の行です。番号クリアを中止しました。"
        Exit Sub
    End If
    mCmdClear.Enabled = False
    apiInvoked = True
    completed = SyncReviewClearCopyIdentifiers(item, SelectedRow())
    If completed Then
        mLblStatus.Caption = "番号クリアが完了しました。最新候補を再読み込みします。"
        mCmdRefresh_Click
    Else
        mLblStatus.Caption = "番号クリアは完了していません。データ状態を確認してください。"
        If Not completed Then mCmdClear.Enabled = IsCopyCandidate(item, SelectedRow())
    End If
    Exit Sub
Failed:
    If apiInvoked And Err.Number = vbObjectError + 6155 Then
        mLblStatus.Caption = "重大な復元失敗です。ブックを保存せず終了し、事前バックアップを保全してください。"
        MsgBox "番号クリア後の復元に失敗しました。ブックを保存せず終了し、バックアップを保全してください。" & _
            vbCrLf & Err.Description, vbCritical
    Else
        mLblStatus.Caption = Err.Description
        If apiInvoked And IsCopyCandidate(item, SelectedRow()) Then mCmdClear.Enabled = True
        If Not apiInvoked Then MsgBox "番号クリアに失敗しました: " & Err.Description, vbExclamation
    End If
End Sub

Private Sub mChkDelete_Click()
    Dim item As Object, key As String, rowNo As Long
    If rendering Then Exit Sub
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    key = CStr(item("key"))
    rowNo = SelectedRow()
    If CStr(item("category")) <> "C" Or rowNo < 2 Then
        rendering = True: mChkDelete.Value = False: rendering = False
        mLblStatus.Caption = "削除対象にできるExcel行を選択してください。"
        Exit Sub
    End If
    If mChkDelete.Value And Len(SyncReviewDecision(item)) > 0 Then
        rendering = True: mChkDelete.Value = False: rendering = False
        mLblStatus.Caption = "保留・別物の判断済み行は削除できません。"
        Exit Sub
    End If
    If mChkDelete.Value Then
        If Not deletionSelection.Exists(key) Then deletionSelection.Add key, rowNo
    ElseIf deletionSelection.Exists(key) Then
        deletionSelection.Remove key
    End If
    mLblStatus.Caption = "削除対象のExcel行：" & deletionSelection.Count & " 件。削除前に全件再確認・事前バックアップを実施します。"
End Sub
Private Sub SelectDeleteRow(ByVal item As Object, ByVal rowNo As Long)
    Dim key As String
    If item Is Nothing Then Exit Sub
    If CStr(item("category")) <> "C" Or rowNo < 2 Then Exit Sub
    key = CStr(item("key"))
    If Len(SyncReviewDecision(item)) > 0 Then
        mLblStatus.Caption = "保留・別物と判断済みの行は削除対象にできません。"
        Exit Sub
    End If
    If Not deletionSelection.Exists(key) Then deletionSelection.Add key, rowNo
    rendering = True: mChkDelete.Value = True: rendering = False
    mLblStatus.Caption = "Excel行" & rowNo & "を削除対象に追加しました。別の候補行も選択できます。"
End Sub
Private Sub mCmdGoTo_Click()
    Dim item As Object
    On Error GoTo Failed
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If CStr(item("category")) = "D" Or SelectedRow() < 2 Then Exit Sub
    SyncReviewNavigate item, SelectedRow()
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
Private Sub mCmdApprove_Click()
    Dim item As Object
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If CStr(item("category")) = "D" Then Exit Sub
    If SyncReviewApprove(item, SelectedRow()) Then mCmdRefresh_Click
End Sub
Private Sub mCmdOther_Click()
    RecordChoice "別物"
End Sub
Private Sub mCmdHold_Click()
    RecordChoice "保留"
End Sub
Private Sub RecordChoice(ByVal decision As String)
    Dim item As Object, rowNo As Long, nextRow As Long
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If CStr(item("category")) = "D" Then Exit Sub
    If CStr(item("category")) = "B" And decision = "別物" Then
        rowNo = SelectedRow()
        If SyncReviewDecide(item, decision, rowNo) Then
            UpdateBRowJudgments item
            RenderComparison item, rowNo
            nextRow = BNextUnconfirmedRow(item, rowNo)
            If nextRow > 0 Then
                SelectCaseRow nextRow
            Else
                mCmdNext_Click
            End If
        End If
        Exit Sub
    End If
    If SyncReviewDecide(item, decision) Then
        If CStr(item("category")) = "C" Then
            If deletionSelection.Exists(CStr(item("key"))) Then deletionSelection.Remove CStr(item("key"))
            rendering = True: mChkDelete.Value = False: rendering = False
        End If
        ShowCategory
        mCmdNext_Click
    End If
End Sub
Private Sub mCmdNext_Click()
    Dim i As Long, startAt As Long, item As Object, candidate As Variant
    Dim candidates As Collection, selected As Long, j As Long, rowNo As Long
    If shown Is Nothing Then Exit Sub
    If shown.Count = 0 Then Exit Sub
    Set item = SelectedCase()
    If Not item Is Nothing Then
        If CStr(item("category")) = "D" Then
            Set candidates = item("copyCandidateRows")
            selected = SelectedRow()
            If Not IsCopyCandidate(item, selected) Then
                If candidates.Count > 0 Then SelectCaseRow CLng(candidates(1))
                Exit Sub
            End If
            For j = 1 To candidates.Count
                If CLng(candidates(j)) = selected And j < candidates.Count Then
                    SelectCaseRow CLng(candidates(j + 1)): Exit Sub
                End If
            Next j
            startAt = mLstCases.ListIndex
            For i = 1 To shown.Count
                Set item = shown(((startAt + i) Mod shown.Count) + 1)
                If CStr(item("category")) = "D" Then
                    Set candidates = item("copyCandidateRows")
                    If candidates.Count > 0 Then
                        mLstCases.ListIndex = (startAt + i) Mod shown.Count
                        ShowCurrent
                        SelectCaseRow CLng(candidates(1))
                        Exit Sub
                    End If
                End If
            Next i
            mLblStatus.Caption = "このカテゴリの未処理クリア候補はありません。"
            Exit Sub
        ElseIf CStr(item("category")) = "B" And Len(SyncReviewDecision(item)) = 0 Then
            rowNo = BNextUnconfirmedRow(item, SelectedRow())
            If rowNo > 0 Then
                SelectCaseRow rowNo
                Exit Sub
            End If
        End If
    End If
    startAt = mLstCases.ListIndex
    For i = 1 To shown.Count
        Set item = shown(((startAt + i) Mod shown.Count) + 1)
        If Len(SyncReviewDecision(item)) = 0 Then
            mLstCases.ListIndex = (startAt + i) Mod shown.Count
            ShowCurrent
            Exit Sub
        End If
    Next i
    mLblStatus.Caption = "このカテゴリの未確認候補はありません。"
End Sub

Private Sub SelectCaseRow(ByVal rowNo As Long)
    Dim i As Long
    For i = 0 To mLstRows.ListCount - 1
        If CLng(Split(CStr(mLstRows.List(i)), " ")(0)) = rowNo Then
            mLstRows.ListIndex = i
            ShowCurrentSelection
            Exit Sub
        End If
    Next i
End Sub
Private Sub mCmdDelete_Click()
    Dim selected As New Collection, i As Long, item As Object, deletionItem As Object, completed As Boolean
    If cases Is Nothing Then Exit Sub
    Set item = SelectedCase()
    If Not item Is Nothing Then
        If CStr(item("category")) = "D" Then Exit Sub
    End If
    For i = 1 To cases.Count
        Set item = cases(i)
        If CStr(item("category")) = "C" Then
            If deletionSelection.Exists(CStr(item("key"))) And Len(SyncReviewDecision(item)) = 0 Then
                Set deletionItem = SyncReviewSelectedDeletionItem(item, CLng(deletionSelection(CStr(item("key")))))
                If deletionItem Is Nothing Then Err.Raise vbObjectError + 6142, , "削除対象のExcel行を選択してください。"
                selected.Add deletionItem
            End If
        End If
    Next i
    If selected.Count = 0 Then
        mLblStatus.Caption = "削除対象のExcel行を選択してください。"
        MsgBox "削除対象のExcel行を選択してください。CSV行ではなく、C候補のExcel行を選択します。", vbExclamation
        Exit Sub
    End If
    completed = SyncReviewDeleteBatch(selected)
    If completed Then
        Set deletionSelection = CreateObject("Scripting.Dictionary")
        mCmdRefresh_Click
    Else
        mLblStatus.Caption = "削除は完了していません。選択内容とバックアップ表示を確認してください。"
    End If
End Sub
Private Sub mCmdRefresh_Click()
    Dim refreshed As Collection, copyCases As Collection, copyItem As Object, previousCases As Collection
    Dim unavailableReason As String, detectorNumber As Long
    Dim detectorDescription As String, detectorSource As String, newSelection As Object
    Dim previousSelection As Object, previousReason As String, failureText As String
    Dim previousCategory As Long, committed As Boolean
    On Error GoTo Failed
    Set refreshed = SyncReviewCandidates(targetYear)
    Err.Clear
    On Error Resume Next
    Set copyCases = SyncReviewCopyIdCandidates(targetYear)
    detectorNumber = Err.Number
    detectorDescription = Err.Description
    detectorSource = Err.Source
    Err.Clear
    On Error GoTo Failed
    If detectorNumber <> 0 Then
        If detectorNumber <> vbObjectError + 6141 Then Err.Raise detectorNumber, detectorSource, detectorDescription
        unavailableReason = detectorDescription
    Else
        For Each copyItem In copyCases
            refreshed.Add copyItem
        Next copyItem
    End If
    For Each copyItem In refreshed
        If CStr(copyItem("category")) <> "D" Then detectorDescription = SyncReviewDecision(copyItem)
    Next copyItem
    Set newSelection = CreateObject("Scripting.Dictionary")
    Set previousCases = cases
    Set previousSelection = deletionSelection
    previousReason = mDUnavailableReason
    previousCategory = mCboCategory.ListIndex
    committed = True
    Set cases = refreshed
    Set deletionSelection = newSelection
    mDUnavailableReason = unavailableReason
    UpdateDWarning
    If Len(mDUnavailableReason) > 0 And mCboCategory.ListIndex = 3 Then
        rendering = True: mCboCategory.ListIndex = 0: rendering = False
    End If
    SelectDWhenOnlyCategory
    ShowCategory
    If Len(mDUnavailableReason) > 0 Then MsgBox "D候補を利用できません: " & mDUnavailableReason, vbExclamation
    Exit Sub
Failed:
    failureText = Err.Description
    If committed Then
        On Error Resume Next
        Err.Clear
        Set cases = previousCases
        Set deletionSelection = previousSelection
        mDUnavailableReason = previousReason
        rendering = True: mCboCategory.ListIndex = previousCategory: rendering = False
        UpdateDWarning
        ShowCategory
        If Err.Number <> 0 Then
            Unload Me
            MsgBox "再読み込みと旧表示の復元に失敗しました。フォームを閉じて再度開いてください: " & failureText, vbCritical
            Exit Sub
        End If
        On Error GoTo 0
    End If
    Unload Me
    MsgBox "再読み込みを安全に中止し、フォームを閉じました。再度開いてください: " & failureText, vbExclamation
End Sub
Private Sub mCmdFinish_Click()
    Unload Me
End Sub

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
Private mTxtDetails As MSForms.TextBox
Private WithEvents mChkDelete As MSForms.CheckBox
Private mLblStatus As MSForms.Label
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

Private Sub BuildControls()
    Me.Caption = "同期候補の確認"
    Me.Width = 605
    Me.Height = 410
    Set mLblCounts = EnsureControl("lblCounts", "Forms.Label.1")
    PlaceControl mLblCounts, 9, 6, 580, 19
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
    Set mTxtDetails = EnsureControl("txtDetails", "Forms.TextBox.1")
    PlaceControl mTxtDetails, 324, 28, 269, 263.5
    Set mChkDelete = EnsureControl("chkDelete", "Forms.CheckBox.1")
    PlaceControl mChkDelete, 9, 301.5, 300, 17.5
    mChkDelete.Caption = "この削除可候補を一括削除の選択に加える"
    Set mLblStatus = EnsureControl("lblStatus", "Forms.Label.1")
    PlaceControl mLblStatus, 9, 323.5, 580, 23
    mLblStatus.Caption = "フォームを開いただけでは変更しません。"
    Set mCmdGoTo = EnsureControl("cmdGoTo", "Forms.CommandButton.1")
    PlaceControl mCmdGoTo, 9, 354.5, 70, 21
    mCmdGoTo.Caption = "対象行へ移動"
    Set mCmdApprove = EnsureControl("cmdApprove", "Forms.CommandButton.1")
    PlaceControl mCmdApprove, 84, 354.5, 78, 21
    mCmdApprove.Caption = "同一として処理"
    Set mCmdOther = EnsureControl("cmdOther", "Forms.CommandButton.1")
    PlaceControl mCmdOther, 167, 354.5, 50, 21
    mCmdOther.Caption = "別物"
    Set mCmdHold = EnsureControl("cmdHold", "Forms.CommandButton.1")
    PlaceControl mCmdHold, 222, 354.5, 50, 21
    mCmdHold.Caption = "保留"
    Set mCmdNext = EnsureControl("cmdNext", "Forms.CommandButton.1")
    PlaceControl mCmdNext, 277, 354.5, 60, 21
    mCmdNext.Caption = "次の候補"
    Set mCmdDelete = EnsureControl("cmdDelete", "Forms.CommandButton.1")
    PlaceControl mCmdDelete, 342, 354.5, 85.5, 21
    mCmdDelete.Caption = "選択行を一括削除"
    Set mCmdClear = EnsureControl("cmdClear", "Forms.CommandButton.1")
    PlaceControl mCmdClear, 433, 354.5, 125, 21
    mCmdClear.Caption = "選択行の番号をクリア"
    Set mCmdRefresh = EnsureControl("cmdRefresh", "Forms.CommandButton.1")
    PlaceControl mCmdRefresh, 9, 382, 70, 21
    mCmdRefresh.Caption = "再読み込み"
    Set mCmdFinish = EnsureControl("cmdFinish", "Forms.CommandButton.1")
    PlaceControl mCmdFinish, 508, 382, 82, 21
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

Public Sub LoadCases(ByVal loaded As Collection, ByVal yearValue As Long)
    Set cases = loaded
    targetYear = yearValue
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    SelectDWhenOnlyCategory
    ShowCategory
End Sub

Private Sub SelectDWhenOnlyCategory()
    If Not HasReviewCategory("A") And Not HasReviewCategory("B") And _
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
    Dim item As Object, v As Variant, rows As String, decision As String, category As String
    Dim candidates As Collection, candidate As Variant, keeper As Long, rowNo As Long
    Set item = SelectedCase()
    rendering = True
    mLstRows.Clear
    mChkDelete.Value = False
    If item Is Nothing Then
        mTxtDetails.Text = "このカテゴリには候補がありません。"
        mCmdApprove.Enabled = False: mCmdGoTo.Enabled = False: mCmdClear.Enabled = False
        mCmdOther.Enabled = False: mCmdHold.Enabled = False: mCmdDelete.Enabled = False
        mChkDelete.Visible = False
        rendering = False
        Exit Sub
    End If
    category = CStr(item("category"))
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
                mLstRows.AddItem CStr(v) & " 行"
            Next v
        End If
    End If
    If mLstRows.ListCount > 0 Then mLstRows.ListIndex = 0
    decision = ""
    If category <> "D" Then decision = SyncReviewDecision(item)
    mTxtDetails.Text = CStr(item("details")) & vbCrLf & vbCrLf & "前回判断：" & _
        IIf(Len(decision) = 0, "未確認", decision)
    rowNo = SelectedRow()
    mCmdGoTo.Enabled = (category <> "D" And rowNo >= 2)
    mCmdApprove.Enabled = (category = "B" And mLstRows.ListCount = 1)
    mCmdOther.Enabled = (category <> "D")
    mCmdHold.Enabled = (category <> "D")
    mCmdDelete.Enabled = (category <> "D")
    mCmdClear.Enabled = IsCopyCandidate(item, rowNo)
    mChkDelete.Visible = (category = "C")
    If mChkDelete.Visible Then mChkDelete.Value = (Len(decision) = 0 And deletionSelection.Exists(CStr(item("key"))))
    If category = "D" Then
        mLblStatus.Caption = "D候補：保持行はクリアできません。選択した候補行のみ専用APIで再検証します。"
    Else
        mLblStatus.Caption = "Aの重複ID・複数行の統合は自動承認不可。Cはチェック後に件数を確認します。"
    End If
    rendering = False
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
    If Not rendering Then ShowCategory
End Sub
Private Sub mLstCases_Click()
    If Not rendering Then ShowCurrent
End Sub
Private Sub mLstRows_Click()
    If Not rendering Then ShowCurrentSelection
End Sub

Private Sub ShowCurrentSelection()
    Dim item As Object, rowNo As Long
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If CStr(item("category")) <> "D" Then Exit Sub
    rowNo = SelectedRow()
    mCmdClear.Enabled = IsCopyCandidate(item, rowNo)
    If rowNo = CLng(item("copyKeeperRow")) Then
        mLblStatus.Caption = "この行はCSVと一致する番号保持行です。クリアできません。"
    ElseIf mCmdClear.Enabled Then
        mLblStatus.Caption = "D候補行です。ボタンで専用APIの再照合・確認を開始します。"
    Else
        mLblStatus.Caption = "D候補ではありません。クリアできません。"
    End If
End Sub

Private Sub mCmdClear_Click()
    Dim item As Object, completed As Boolean, apiInvoked As Boolean
    On Error GoTo Failed
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
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
    Dim item As Object, key As String
    If rendering Then Exit Sub
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    key = CStr(item("key"))
    If mChkDelete.Value And Len(SyncReviewDecision(item)) > 0 Then
        rendering = True: mChkDelete.Value = False: rendering = False
        mLblStatus.Caption = "保留・別物の判断済み行は削除できません。"
        Exit Sub
    End If
    If mChkDelete.Value Then
        If Not deletionSelection.Exists(key) Then deletionSelection.Add key, True
    ElseIf deletionSelection.Exists(key) Then
        deletionSelection.Remove key
    End If
    mLblStatus.Caption = "削除選択：" & deletionSelection.Count & " 件。削除前に全件再確認・事前バックアップを実施します。"
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
    Dim item As Object
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If CStr(item("category")) = "D" Then Exit Sub
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
    Dim selected As New Collection, i As Long, item As Object, completed As Boolean
    If cases Is Nothing Then Exit Sub
    Set item = SelectedCase()
    If Not item Is Nothing Then
        If CStr(item("category")) = "D" Then Exit Sub
    End If
    For i = 1 To cases.Count
        Set item = cases(i)
        If CStr(item("category")) = "C" Then
            If deletionSelection.Exists(CStr(item("key"))) And Len(SyncReviewDecision(item)) = 0 Then selected.Add item
        End If
    Next i
    completed = SyncReviewDeleteBatch(selected)
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    mCmdRefresh_Click
End Sub
Private Sub mCmdRefresh_Click()
    Dim copyCases As Collection, copyItem As Object
    On Error GoTo Failed
    Set cases = SyncReviewCandidates(targetYear)
    Set copyCases = SyncReviewCopyIdCandidates(targetYear)
    For Each copyItem In copyCases
        cases.Add copyItem
    Next copyItem
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    SelectDWhenOnlyCategory
    ShowCategory
    Exit Sub
Failed:
    MsgBox "再読み込みに失敗: " & Err.Description, vbExclamation
End Sub
Private Sub mCmdFinish_Click()
    Unload Me
End Sub

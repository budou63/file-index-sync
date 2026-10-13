VERSION 5.00
Begin VB.UserForm frmSyncReview
   Caption = "同期候補の確認"
   ClientHeight = 8200
   ClientWidth = 12100
   ShowModal = 0
   StartUpPosition = 0
   Begin VB.Label lblCounts
      Caption = "候補を読み込みます"
      Height = 380
      Left = 180
      Top = 120
      Width = 11600
   End
   Begin VB.ComboBox cboCategory
      Height = 330
      Left = 180
      Top = 560
      Width = 3750
   End
   Begin VB.Label lblCase
      Caption = "案件一覧（1件ずつ確認）"
      Height = 280
      Left = 180
      Top = 1000
      Width = 3750
   End
   Begin VB.ListBox lstCases
      Height = 4560
      Left = 180
      Top = 1290
      Width = 3750
   End
   Begin VB.Label lblRows
      Caption = "候補のExcel行"
      Height = 280
      Left = 4100
      Top = 1000
      Width = 2200
   End
   Begin VB.ListBox lstRows
      Height = 4560
      Left = 4100
      Top = 1290
      Width = 2200
   End
   Begin VB.TextBox txtDetails
      Height = 5270
      Left = 6480
      Top = 560
      Width = 5380
      MultiLine = -1
      ScrollBars = 2
      Locked = -1
   End
   Begin VB.CheckBox chkDelete
      Caption = "この削除可候補を一括削除の選択に加える"
      Height = 350
      Left = 180
      Top = 6030
      Width = 6000
   End
   Begin VB.Label lblStatus
      Caption = "フォームを開いただけでは変更しません。"
      Height = 460
      Left = 180
      Top = 6470
      Width = 11600
   End
   Begin VB.CommandButton cmdGoTo
      Caption = "対象行へ移動"
      Height = 420
      Left = 180
      Top = 7090
      Width = 1400
   End
   Begin VB.CommandButton cmdApprove
      Caption = "同一として処理"
      Height = 420
      Left = 1680
      Top = 7090
      Width = 1560
   End
   Begin VB.CommandButton cmdOther
      Caption = "別物"
      Height = 420
      Left = 3340
      Top = 7090
      Width = 1000
   End
   Begin VB.CommandButton cmdHold
      Caption = "保留"
      Height = 420
      Left = 4440
      Top = 7090
      Width = 1000
   End
   Begin VB.CommandButton cmdNext
      Caption = "次の候補"
      Height = 420
      Left = 5540
      Top = 7090
      Width = 1200
   End
   Begin VB.CommandButton cmdDelete
      Caption = "選択行を一括削除"
      Height = 420
      Left = 6840
      Top = 7090
      Width = 1710
   End
   Begin VB.CommandButton cmdRefresh
      Caption = "再読み込み"
      Height = 420
      Left = 8660
      Top = 7090
      Width = 1400
   End
   Begin VB.CommandButton cmdFinish
      Caption = "確認終了"
      Height = 420
      Left = 10160
      Top = 7090
      Width = 1640
   End
End
Attribute VB_Name = "frmSyncReview"
Attribute VB_PredeclaredId = True
Option Explicit

Private mLblCounts As MSForms.Label
Private WithEvents mCboCategory As MSForms.ComboBox
Private mLblCase As MSForms.Label
Private WithEvents mLstCases As MSForms.ListBox
Private mLblRows As MSForms.Label
Private mLstRows As MSForms.ListBox
Private mTxtDetails As MSForms.TextBox
Private WithEvents mChkDelete As MSForms.CheckBox
Private mLblStatus As MSForms.Label
Private WithEvents mCmdGoTo As MSForms.CommandButton
Private WithEvents mCmdApprove As MSForms.CommandButton
Private WithEvents mCmdOther As MSForms.CommandButton
Private WithEvents mCmdHold As MSForms.CommandButton
Private WithEvents mCmdNext As MSForms.CommandButton
Private WithEvents mCmdDelete As MSForms.CommandButton
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
    Set mCmdRefresh = EnsureControl("cmdRefresh", "Forms.CommandButton.1")
    PlaceControl mCmdRefresh, 433, 354.5, 70, 21
    mCmdRefresh.Caption = "再読み込み"
    Set mCmdFinish = EnsureControl("cmdFinish", "Forms.CommandButton.1")
    PlaceControl mCmdFinish, 508, 354.5, 82, 21
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
    mCboCategory.ListIndex = 0
    rendering = False
End Sub

Public Sub LoadCases(ByVal loaded As Collection, ByVal yearValue As Long)
    Set cases = loaded
    targetYear = yearValue
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    ShowCategory
End Sub

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
    Dim i As Long, item As Object, category As String, counts(1 To 3) As Long, unresolved(1 To 3) As Long
    Dim decision As String, idx As Long
    If cases Is Nothing Then Exit Sub
    rendering = True
    Set shown = New Collection
    mLstCases.Clear: mLstRows.Clear
    For i = 1 To cases.Count
        Set item = cases(i)
        category = CStr(item("category"))
        Select Case category
            Case "A": idx = 1
            Case "B": idx = 2
            Case "C": idx = 3
        End Select
        counts(idx) = counts(idx) + 1
        decision = SyncReviewDecision(item)
        If Len(decision) = 0 Then unresolved(idx) = unresolved(idx) + 1
        If (mCboCategory.ListIndex = 0 And category = "A") Or _
           (mCboCategory.ListIndex = 1 And category = "B") Or _
           (mCboCategory.ListIndex = 2 And category = "C") Then
            shown.Add item
            mLstCases.AddItem "CSV行" & CStr(item("sourceRow")) & " / Excel " & CStr(item("rows")) & _
                IIf(Len(decision) = 0, " [未確認]", " [" & decision & "]")
        End If
    Next i
    mLblCounts.Caption = "対象年度 " & targetYear & " / A " & counts(1) & "件（未 " & unresolved(1) & "）" & _
        " / B " & counts(2) & "件（未 " & unresolved(2) & "）" & _
        " / C " & counts(3) & "件（未 " & unresolved(3) & "）"
    mTxtDetails.Text = "候補を選んでください。"
    mChkDelete.Value = False
    If mLstCases.ListCount > 0 Then mLstCases.ListIndex = 0
    rendering = False
    ShowCurrent
End Sub

Private Sub ShowCurrent()
    Dim item As Object, v As Variant, rows As String, decision As String
    Set item = SelectedCase()
    rendering = True
    mLstRows.Clear
    mChkDelete.Value = False
    If item Is Nothing Then
        mTxtDetails.Text = "このカテゴリには候補がありません。"
        mCmdApprove.Enabled = False: mCmdGoTo.Enabled = False
        mChkDelete.Visible = False
        rendering = False
        Exit Sub
    End If
    rows = CStr(item("rows"))
    If Len(rows) > 0 Then
        For Each v In Split(rows, ";")
            mLstRows.AddItem CStr(v) & " 行"
        Next v
        mLstRows.ListIndex = 0
    End If
    decision = SyncReviewDecision(item)
    mTxtDetails.Text = CStr(item("details")) & vbCrLf & vbCrLf & "前回判断：" & _
        IIf(Len(decision) = 0, "未確認", decision)
    mCmdGoTo.Enabled = (mLstRows.ListCount > 0)
    mCmdApprove.Enabled = (CStr(item("category")) = "B" And mLstRows.ListCount = 1)
    mChkDelete.Visible = (CStr(item("category")) = "C")
    If mChkDelete.Visible Then mChkDelete.Value = (Len(decision) = 0 And deletionSelection.Exists(CStr(item("key"))))
    mLblStatus.Caption = "Aの重複ID・複数行の統合は自動承認不可。Cはチェック後に件数を確認します。"
    rendering = False
End Sub

Private Sub mCboCategory_Change()
    If Not rendering Then ShowCategory
End Sub
Private Sub mLstCases_Click()
    If Not rendering Then ShowCurrent
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
    If item Is Nothing Or SelectedRow() < 2 Then Exit Sub
    SyncReviewNavigate item, SelectedRow()
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
Private Sub mCmdApprove_Click()
    Dim item As Object
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
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
    Dim i As Long, startAt As Long, item As Object
    If shown Is Nothing Then Exit Sub
    If shown.Count = 0 Then Exit Sub
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
Private Sub mCmdDelete_Click()
    Dim selected As New Collection, i As Long, item As Object, completed As Boolean
    If cases Is Nothing Then Exit Sub
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
    On Error GoTo Failed
    Set cases = SyncReviewCandidates(targetYear)
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    ShowCategory
    Exit Sub
Failed:
    MsgBox "再読み込みに失敗: " & Err.Description, vbExclamation
End Sub
Private Sub mCmdFinish_Click()
    Unload Me
End Sub

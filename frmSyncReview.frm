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

Private cases As Collection
Private shown As Collection
Private deletionSelection As Object
Private targetYear As Long
Private rendering As Boolean

Private Sub UserForm_Initialize()
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    cboCategory.AddItem "A：判定困難・ID競合"
    cboCategory.AddItem "B：同名・類似候補"
    cboCategory.AddItem "C：削除可候補"
    cboCategory.ListIndex = 0
End Sub

Public Sub LoadCases(ByVal loaded As Collection, ByVal yearValue As Long)
    Set cases = loaded
    targetYear = yearValue
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    ShowCategory
End Sub

Private Function SelectedCase() As Object
    If shown Is Nothing Then Exit Function
    If lstCases.ListIndex < 0 Or lstCases.ListIndex + 1 > shown.Count Then Exit Function
    Set SelectedCase = shown(lstCases.ListIndex + 1)
End Function

Private Function SelectedRow() As Long
    If lstRows.ListIndex < 0 Then Exit Function
    SelectedRow = CLng(Split(CStr(lstRows.List(lstRows.ListIndex)), " ")(0))
End Function

Private Sub ShowCategory()
    Dim i As Long, item As Object, category As String, counts(1 To 3) As Long, unresolved(1 To 3) As Long
    Dim decision As String, idx As Long
    If cases Is Nothing Then Exit Sub
    rendering = True
    Set shown = New Collection
    lstCases.Clear: lstRows.Clear
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
        If (cboCategory.ListIndex = 0 And category = "A") Or _
           (cboCategory.ListIndex = 1 And category = "B") Or _
           (cboCategory.ListIndex = 2 And category = "C") Then
            shown.Add item
            lstCases.AddItem "CSV行" & CStr(item("sourceRow")) & " / Excel " & CStr(item("rows")) & _
                IIf(Len(decision) = 0, " [未確認]", " [" & decision & "]")
        End If
    Next i
    lblCounts.Caption = "対象年度 " & targetYear & " / A " & counts(1) & "件（未 " & unresolved(1) & "）" & _
        " / B " & counts(2) & "件（未 " & unresolved(2) & "）" & _
        " / C " & counts(3) & "件（未 " & unresolved(3) & "）"
    txtDetails.Text = "候補を選んでください。"
    chkDelete.Value = False
    rendering = False
    If lstCases.ListCount > 0 Then lstCases.ListIndex = 0
    ShowCurrent
End Sub

Private Sub ShowCurrent()
    Dim item As Object, v As Variant, rows As String, decision As String
    Set item = SelectedCase()
    rendering = True
    lstRows.Clear
    chkDelete.Value = False
    If item Is Nothing Then
        txtDetails.Text = "このカテゴリには候補がありません。"
        cmdApprove.Enabled = False: cmdGoTo.Enabled = False
        chkDelete.Visible = False
        rendering = False
        Exit Sub
    End If
    rows = CStr(item("rows"))
    If Len(rows) > 0 Then
        For Each v In Split(rows, ";")
            lstRows.AddItem CStr(v) & " 行"
        Next v
        lstRows.ListIndex = 0
    End If
    decision = SyncReviewDecision(item)
    txtDetails.Text = CStr(item("details")) & vbCrLf & vbCrLf & "前回判断：" & _
        IIf(Len(decision) = 0, "未確認", decision)
    cmdGoTo.Enabled = (lstRows.ListCount > 0)
    cmdApprove.Enabled = (CStr(item("category")) = "B" And lstRows.ListCount = 1)
    chkDelete.Visible = (CStr(item("category")) = "C")
    If chkDelete.Visible Then chkDelete.Value = (Len(decision) = 0 And deletionSelection.Exists(CStr(item("key"))))
    lblStatus.Caption = "Aの重複ID・複数行の統合は自動承認不可。Cはチェック後に件数を確認します。"
    rendering = False
End Sub

Private Sub cboCategory_Change()
    If Not rendering Then ShowCategory
End Sub
Private Sub lstCases_Click()
    If Not rendering Then ShowCurrent
End Sub
Private Sub chkDelete_Click()
    Dim item As Object, key As String
    If rendering Then Exit Sub
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    key = CStr(item("key"))
    If chkDelete.Value And Len(SyncReviewDecision(item)) > 0 Then
        rendering = True: chkDelete.Value = False: rendering = False
        lblStatus.Caption = "保留・別物の判断済み行は削除できません。"
        Exit Sub
    End If
    If chkDelete.Value Then
        If Not deletionSelection.Exists(key) Then deletionSelection.Add key, True
    ElseIf deletionSelection.Exists(key) Then
        deletionSelection.Remove key
    End If
    lblStatus.Caption = "削除選択：" & deletionSelection.Count & " 件。削除前に全件再確認・事前バックアップを実施します。"
End Sub
Private Sub cmdGoTo_Click()
    Dim item As Object
    On Error GoTo Failed
    Set item = SelectedCase()
    If item Is Nothing Or SelectedRow() < 2 Then Exit Sub
    SyncReviewNavigate item, SelectedRow()
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
Private Sub cmdApprove_Click()
    Dim item As Object
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If SyncReviewApprove(item, SelectedRow()) Then cmdRefresh_Click
End Sub
Private Sub cmdOther_Click()
    RecordChoice "別物"
End Sub
Private Sub cmdHold_Click()
    RecordChoice "保留"
End Sub
Private Sub RecordChoice(ByVal decision As String)
    Dim item As Object
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If SyncReviewDecide(item, decision) Then
        If CStr(item("category")) = "C" Then
            If deletionSelection.Exists(CStr(item("key"))) Then deletionSelection.Remove CStr(item("key"))
            rendering = True: chkDelete.Value = False: rendering = False
        End If
        ShowCategory
        cmdNext_Click
    End If
End Sub
Private Sub cmdNext_Click()
    Dim i As Long, startAt As Long, item As Object
    If shown Is Nothing Then Exit Sub
    If shown.Count = 0 Then Exit Sub
    startAt = lstCases.ListIndex
    For i = 1 To shown.Count
        Set item = shown(((startAt + i) Mod shown.Count) + 1)
        If Len(SyncReviewDecision(item)) = 0 Then
            lstCases.ListIndex = (startAt + i) Mod shown.Count
            ShowCurrent
            Exit Sub
        End If
    Next i
    lblStatus.Caption = "このカテゴリの未確認候補はありません。"
End Sub
Private Sub cmdDelete_Click()
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
    cmdRefresh_Click
End Sub
Private Sub cmdRefresh_Click()
    On Error GoTo Failed
    Set cases = SyncReviewCandidates(targetYear)
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    ShowCategory
    Exit Sub
Failed:
    MsgBox "再読み込みに失敗: " & Err.Description, vbExclamation
End Sub
Private Sub cmdFinish_Click()
    Unload Me
End Sub

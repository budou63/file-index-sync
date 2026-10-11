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
    If lstCases.ListIndex < 0 Or lstCases.ListIndex + 1 > shown.count Then Exit Function
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
    For i = 1 To cases.count
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
    txtDetails.text = "候補を選んでください。"
    chkDelete.value = False
    rendering = False
    If lstCases.ListCount > 0 Then lstCases.ListIndex = 0
    ShowCurrent
End Sub

Private Sub ShowCurrent()
    Dim item As Object, v As Variant, rows As String, decision As String
    Set item = SelectedCase()
    rendering = True
    lstRows.Clear
    chkDelete.value = False
    If item Is Nothing Then
        txtDetails.text = "このカテゴリには候補がありません。"
        cmdApprove.Enabled = False: cmdGoTo.Enabled = False
        chkDelete.visible = False
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
    txtDetails.text = CStr(item("details")) & vbCrLf & vbCrLf & "前回判断：" & _
        IIf(Len(decision) = 0, "未確認", decision)
    cmdGoTo.Enabled = (lstRows.ListCount > 0)
    cmdApprove.Enabled = (CStr(item("category")) = "B" And lstRows.ListCount = 1)
    chkDelete.visible = (CStr(item("category")) = "C")
    If chkDelete.visible Then chkDelete.value = (Len(decision) = 0 And deletionSelection.Exists(CStr(item("key"))))
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
    If chkDelete.value And Len(SyncReviewDecision(item)) > 0 Then
        rendering = True: chkDelete.value = False: rendering = False
        lblStatus.Caption = "保留・別物の判断済み行は削除できません。"
        Exit Sub
    End If
    If chkDelete.value Then
        If Not deletionSelection.Exists(key) Then deletionSelection.Add key, True
    ElseIf deletionSelection.Exists(key) Then
        deletionSelection.Remove key
    End If
    lblStatus.Caption = "削除選択：" & deletionSelection.count & " 件。削除前に全件再確認・事前バックアップを実施します。"
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
            rendering = True: chkDelete.value = False: rendering = False
        End If
        ShowCategory
        cmdNext_Click
    End If
End Sub
Private Sub cmdNext_Click()
    Dim i As Long, startAt As Long, item As Object
    If shown Is Nothing Then Exit Sub
    If shown.count = 0 Then Exit Sub
    startAt = lstCases.ListIndex
    For i = 1 To shown.count
        Set item = shown(((startAt + i) Mod shown.count) + 1)
        If Len(SyncReviewDecision(item)) = 0 Then
            lstCases.ListIndex = (startAt + i) Mod shown.count
            ShowCurrent
            Exit Sub
        End If
    Next i
    lblStatus.Caption = "このカテゴリの未確認候補はありません。"
End Sub
Private Sub cmdDelete_Click()
    Dim selected As New Collection, i As Long, item As Object, completed As Boolean
    If cases Is Nothing Then Exit Sub
    For i = 1 To cases.count
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

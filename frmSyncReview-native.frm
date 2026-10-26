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
Private WithEvents mCboEGroup As MSForms.ComboBox
Private mLblEVisibility As MSForms.Label
Private WithEvents mCboEVisibility As MSForms.ComboBox
Private mLblEWindowSize As MSForms.Label
Private WithEvents mCboEWindowSize As MSForms.ComboBox
Private mLblEPaneSize As MSForms.Label
Private WithEvents mCboEPaneSize As MSForms.ComboBox
Private mLblESelection As MSForms.Label
Private WithEvents mCmdEClear As MSForms.CommandButton
Private mLblCase As MSForms.Label
Private WithEvents mLstCases As MSForms.ListBox
Private mLblRows As MSForms.Label
Private WithEvents mLstRows As MSForms.ListBox
Private mLblCompareRows As MSForms.Label
Private mFraCompare As MSForms.Frame
Private mCompareHeaders(0 To 2) As MSForms.Label
Private mEColumnHeaders(0 To 4) As MSForms.Label
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
Private eFieldCases As Collection
Private eFieldItems As Collection
Private deletionSelection As Object
Private targetYear As Long
Private rendering As Boolean
' Display-only caches live for one synchronous render; never used by adoption validators.
Private eDisplayHeaderMaps As Object
Private eDisplayNormalizedKeys As Object
Private eRenderedSelection As String
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
    Dim eHeaderCaptions As Variant, eHeaderWidths As Variant, eHeaderLefts As Variant
    Me.Caption = "同期候補の確認"
    Me.Width = 760
    Me.Height = Me.Height + (440 - Me.InsideHeight)
    Set mLblCounts = EnsureControl("lblCounts", "Forms.Label.1")
    PlaceControl mLblCounts, 9, 6, 740, 19
    mLblCounts.Caption = "候補を読み込みます"
    Set mCboCategory = EnsureControl("cboCategory", "Forms.ComboBox.1")
    PlaceControl mCboCategory, 9, 28, 187.5, 16.5
    Set mCboEGroup = EnsureControl("cboEGroup", "Forms.ComboBox.1")
    PlaceControl mCboEGroup, 205, 28, 110, 16.5
    mCboEGroup.Style = fmStyleDropDownList
    mCboEGroup.ControlTipText = "項目グループで対象ファイルを絞り込みます。判断状態は変更しません。"
    mCboEGroup.Visible = False
    Set mLblEVisibility = EnsureControl("lblEVisibility", "Forms.Label.1")
    PlaceControl mLblEVisibility, 324, 28, 42, 16.5
    mLblEVisibility.Caption = "項目表示"
    mLblEVisibility.Visible = False
    Set mCboEVisibility = EnsureControl("cboEVisibility", "Forms.ComboBox.1")
    PlaceControl mCboEVisibility, 370, 28, 112, 16.5
    mCboEVisibility.Style = fmStyleDropDownList
    mCboEVisibility.Visible = False
    Set mLblEWindowSize = EnsureControl("lblEWindowSize", "Forms.Label.1")
    PlaceControl mLblEWindowSize, 525, 28, 42, 16.5
    mLblEWindowSize.Caption = "画面"
    mLblEWindowSize.Visible = False
    Set mCboEWindowSize = EnsureControl("cboEWindowSize", "Forms.ComboBox.1")
    PlaceControl mCboEWindowSize, 570, 28, 88, 16.5
    mCboEWindowSize.Style = fmStyleDropDownList
    mCboEWindowSize.ControlTipText = "画面サイズをプリセットから選びます（ウィンドウのドラッグ変更はできません）。"
    mCboEWindowSize.Visible = False
    Set mLblEPaneSize = EnsureControl("lblEPaneSize", "Forms.Label.1")
    PlaceControl mLblEPaneSize, 664, 28, 36, 16.5
    mLblEPaneSize.Caption = "幅"
    mLblEPaneSize.Visible = False
    Set mCboEPaneSize = EnsureControl("cboEPaneSize", "Forms.ComboBox.1")
    PlaceControl mCboEPaneSize, 704, 28, 112, 16.5
    mCboEPaneSize.Style = fmStyleDropDownList
    mCboEPaneSize.ControlTipText = "ファイル一覧と不一致比較領域の幅配分を選びます。"
    mCboEPaneSize.Visible = False
    Set mLblESelection = EnsureControl("lblESelection", "Forms.Label.1")
    PlaceControl mLblESelection, 9, 301.5, 190, 17.5
    mLblESelection.Visible = False
    Set mCmdEClear = EnsureControl("cmdEClear", "Forms.CommandButton.1")
    PlaceControl mCmdEClear, 205, 300, 110, 21
    mCmdEClear.Caption = "選択解除"
    mCmdEClear.Visible = False
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
    eHeaderCaptions = Array("項目名", "前回確認した値", "文書管理システム", "Excel基準表", "判断状況")
    eHeaderWidths = Array(130, 145, 145, 145, 180)
    eHeaderLefts = Array(263, 393, 538, 683, 828)
    For i = 0 To 4
        Set mEColumnHeaders(i) = EnsureControl("lblEColumnHeader" & CStr(i), "Forms.Label.1")
        PlaceControl mEColumnHeaders(i), CSng(eHeaderLefts(i)), 50, CSng(eHeaderWidths(i)), 14
        mEColumnHeaders(i).Caption = CStr(eHeaderCaptions(i))
        mEColumnHeaders(i).Font.Bold = True
        mEColumnHeaders(i).Visible = False
    Next i
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
    mCboCategory.AddItem "E：3方向編集競合"
    mCboCategory.ListIndex = 0
    mCboEVisibility.AddItem "入力項目のみ"
    mCboEVisibility.AddItem "全項目（管理コード含む）"
    mCboEVisibility.ListIndex = 0
    mCboEWindowSize.AddItem "コンパクト"
    mCboEWindowSize.AddItem "標準"
    mCboEWindowSize.AddItem "拡大"
    mCboEWindowSize.ListIndex = 1
    mCboEPaneSize.AddItem "標準"
    mCboEPaneSize.AddItem "一覧を広く"
    mCboEPaneSize.AddItem "比較を広く"
    mCboEPaneSize.ListIndex = 0
    rendering = False
End Sub

Public Sub LoadCases(ByVal loaded As Collection, ByVal yearValue As Long, _
                     Optional ByVal unavailableReason As String = "")
    EndEDisplayRead
    eRenderedSelection = ""
    Set cases = loaded
    targetYear = yearValue
    mDUnavailableReason = unavailableReason
    Set deletionSelection = CreateObject("Scripting.Dictionary")
    PopulateEGroups
    UpdateDWarning
    If Len(mDUnavailableReason) > 0 And mCboCategory.ListIndex = 3 Then
        rendering = True: mCboCategory.ListIndex = 0: rendering = False
    End If
    If mCboCategory.ListIndex = 4 And Not HasReviewCategory("E") Then
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
    ElseIf Not HasReviewCategory("A") And Not HasReviewCategory("B") And _
       Not HasReviewCategory("C") And Not HasReviewCategory("D") And HasReviewCategory("E") Then
        rendering = True: mCboCategory.ListIndex = 4: rendering = False
    End If
End Sub

Private Sub PopulateEGroups()
    Dim item As Object, groupNames As Object, name As Variant
    Set groupNames = CreateObject("Scripting.Dictionary")
    groupNames.CompareMode = vbTextCompare
    rendering = True
    mCboEGroup.Clear
    mCboEGroup.AddItem "すべて"
    If Not cases Is Nothing Then
        For Each item In cases
            If CStr(item("category")) = "E" And CLng(item("year")) = targetYear Then
                If Not groupNames.Exists(CStr(item("groupName"))) Then _
                    groupNames.Add CStr(item("groupName")), True
            End If
        Next item
    End If
    For Each name In groupNames.Keys: mCboEGroup.AddItem CStr(name): Next name
    mCboEGroup.ListIndex = 0
    rendering = False
End Sub

Private Function ESelectedCount() As Long
    Dim i As Long
    If mCboCategory.ListIndex <> 4 Then Exit Function
    For i = 0 To mLstCases.ListCount - 1
        If mLstCases.Selected(i) Then ESelectedCount = ESelectedCount + 1
    Next i
End Function

Private Sub UpdateESelection()
    Dim n As Long, item As Object, record As Object, selectedField As Object
    Dim identitySafe As Boolean, groupSafe As Boolean, systemSafe As Boolean, excelSafe As Boolean, holdSafe As Boolean
    Dim systemReason As String, excelReason As String, holdReason As String, pairPreview As String, pairReason As String
    If mCboCategory.ListIndex <> 4 Then Exit Sub
    n = ESelectedCount()
    Set item = SelectedCase()
    Set record = SelectedRecord()
    Set selectedField = SelectedEField()
    If n = 1 Then
        If Not record Is Nothing Then identitySafe = Not CBool(record("identityIssue"))
        If Not item Is Nothing And identitySafe Then
            If EIsClassification1Group(item) Then
                systemSafe = ESelectedDecisionIsSafe("システム採用", systemReason)
                excelSafe = ESelectedDecisionIsSafe("Excel採用", excelReason)
                holdSafe = Module1.BidirectionalSyncClassification1DecisionSafe(item, "保留", pairReason, pairPreview)
                If Not holdSafe And Len(holdReason) = 0 Then holdReason = pairReason
            ElseIf EIsRetentionGroup(item) Then
                systemSafe = ESelectedDecisionIsSafe("システム採用", systemReason)
                excelSafe = ESelectedDecisionIsSafe("Excel採用", excelReason)
                holdSafe = Module1.BidirectionalSyncRetentionDecisionSafe(item, "保留", pairReason, pairPreview)
                If Not holdSafe And Len(holdReason) = 0 Then holdReason = pairReason
            Else
                groupSafe = (EIsAllDisplayMode() Or Not ECaseHasOutOfScopeFields(item))
                systemSafe = groupSafe: excelSafe = groupSafe: holdSafe = groupSafe
            End If
        End If
    End If
    mLblESelection.Caption = CStr(n) & "件選択中"
    mCmdApprove.Enabled = (n = 1 And Not item Is Nothing And identitySafe And systemSafe)
    mCmdOther.Enabled = (n = 1 And Not item Is Nothing And identitySafe And excelSafe)
    mCmdHold.Enabled = (n = 1 And Not item Is Nothing And identitySafe And holdSafe)
    mCmdEClear.Enabled = (n > 0)
    If n > 1 Then
        mLblStatus.Caption = "複数選択は表示専用です。判断は1件だけ選び直してください。"
    ElseIf n = 1 And Not record Is Nothing And Not identitySafe Then
        mLblStatus.Caption = "識別情報の一貫性を確認できないため、採用・保留はできません。正式ID等を確認してください。"
    ElseIf n = 1 And Not item Is Nothing And EIsClassification1Group(item) Then
        mLblStatus.Caption = "分類名１と分類コード１は連動した判断単位です。採用前の確認画面で両方の値と採用元を確認してください。"
        If Not selectedField Is Nothing And Not record Is Nothing Then
            mTxtDetails.Text = ERecordIdentityDetails(record) & vbCrLf & EConflictFieldDetails(item, selectedField)
        End If
        If Len(pairPreview) > 0 Then mTxtDetails.Text = mTxtDetails.Text & vbCrLf & pairPreview
        If Len(pairReason) > 0 Then mTxtDetails.Text = mTxtDetails.Text & vbCrLf & "連動確認: " & pairReason
        If Not systemSafe And Len(systemReason) > 0 Then mTxtDetails.Text = mTxtDetails.Text & vbCrLf & "文書管理を採用できません: " & systemReason
        If Not excelSafe And Len(excelReason) > 0 Then mTxtDetails.Text = mTxtDetails.Text & vbCrLf & "Excel基準表を採用できません: " & excelReason
        If Not holdSafe And Len(holdReason) > 0 Then mTxtDetails.Text = mTxtDetails.Text & vbCrLf & "保留できません: " & holdReason
    ElseIf n = 1 And Not item Is Nothing And EIsRetentionGroup(item) Then
        mLblStatus.Caption = "保存期間と保存期間IDは連動した判断単位です。採用前に現在の目録CSV・コード管理CSVで両側の対応を一意検証します。"
        If Not selectedField Is Nothing And Not record Is Nothing Then
            mTxtDetails.Text = ERecordIdentityDetails(record) & vbCrLf & EConflictFieldDetails(item, selectedField)
        End If
        If Len(pairPreview) > 0 Then mTxtDetails.Text = mTxtDetails.Text & vbCrLf & pairPreview
        If Not systemSafe And Len(systemReason) > 0 Then mTxtDetails.Text = mTxtDetails.Text & vbCrLf & "文書管理を採用できません: " & systemReason
        If Not excelSafe And Len(excelReason) > 0 Then mTxtDetails.Text = mTxtDetails.Text & vbCrLf & "Excel基準表を採用できません: " & excelReason
        If Not holdSafe And Len(holdReason) > 0 Then mTxtDetails.Text = mTxtDetails.Text & vbCrLf & "保留できません: " & holdReason
    ElseIf n = 1 And Not item Is Nothing And Not groupSafe Then
        mLblStatus.Caption = "この判断グループには「入力項目のみ」の表示対象外の関連項目があります。「全項目（管理コード含む）」に切り替えてください。"
    End If
End Sub

Private Sub mCboEGroup_Change()
    If rendering Or mCboCategory.ListIndex <> 4 Then Exit Sub
    ShowCategory
End Sub

Private Sub mCboEVisibility_Change()
    If rendering Or mCboCategory.ListIndex <> 4 Then Exit Sub
    ShowECategory True
End Sub

Private Sub mCboEWindowSize_Change()
    If rendering Or mCboCategory.ListIndex <> 4 Then Exit Sub
    ShowECategory True
End Sub

Private Sub mCboEPaneSize_Change()
    If rendering Or mCboCategory.ListIndex <> 4 Then Exit Sub
    ShowECategory True
End Sub

Private Sub mCmdEClear_Click()
    Dim i As Long
    eRenderedSelection = ""
    If mCboCategory.ListIndex <> 4 Then Exit Sub
    rendering = True
    For i = 0 To mLstCases.ListCount - 1: mLstCases.Selected(i) = False: Next i
    mLstCases.ListIndex = -1
    rendering = False
    ShowCurrent
End Sub

Private Function HasReviewCategory(ByVal category As String) As Boolean
    Dim value As Object
    If cases Is Nothing Then Exit Function
    For Each value In cases
        If CStr(value("category")) = category Then HasReviewCategory = True: Exit Function
    Next value
End Function

Private Function SelectedCase() As Object
    Dim n As Long
    If mCboCategory.ListIndex = 4 Then
        If ESelectedCount() <> 1 Or mLstRows.ListIndex < 0 Then Exit Function
        If eFieldCases Is Nothing Then Exit Function
        If mLstRows.ListIndex + 1 <= eFieldCases.Count Then _
            Set SelectedCase = eFieldCases(mLstRows.ListIndex + 1)
        Exit Function
    End If
    If shown Is Nothing Then Exit Function
    If mLstCases.ListIndex >= 0 And mLstCases.ListIndex + 1 <= shown.Count Then
        If mCboCategory.ListIndex <> 4 Or mLstCases.Selected(mLstCases.ListIndex) Then
            Set SelectedCase = shown(mLstCases.ListIndex + 1)
            Exit Function
        End If
    End If
    If mCboCategory.ListIndex = 4 Then
        For n = 0 To mLstCases.ListCount - 1
            If mLstCases.Selected(n) Then Set SelectedCase = shown(n + 1): Exit Function
        Next n
    End If
End Function

Private Function SelectedRow() As Long
    If mCboCategory.ListIndex = 4 Then Exit Function
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
    Dim originalHeader As String
    originalHeader = rawHeader
    If Not eDisplayNormalizedKeys Is Nothing Then
        If eDisplayNormalizedKeys.Exists(originalHeader) Then
            CompareHeaderKey = CStr(eDisplayNormalizedKeys(originalHeader))
            Exit Function
        End If
    End If
    rawHeader = Replace(Trim$(rawHeader), "　", "")
    rawHeader = Replace(rawHeader, " ", "")
    On Error Resume Next
    rawHeader = StrConv(rawHeader, vbNarrow)
    On Error GoTo 0
    CompareHeaderKey = LCase$(rawHeader)
    If Not eDisplayNormalizedKeys Is Nothing Then eDisplayNormalizedKeys(originalHeader) = CompareHeaderKey
End Function

Private Function FindCompareColumn(ByVal ws As Worksheet, ByVal aliases As Variant) As Long
    Dim headerMap As Object, matchedColumns As Object, alias As Variant, headerKey As String
    Dim colNo As Long, columnKey As Variant
    If ws Is Nothing Or IsEmpty(aliases) Then Exit Function
    Set headerMap = EDisplayHeaderMap(ws)
    Set matchedColumns = CreateObject("Scripting.Dictionary")
    For Each alias In aliases
        headerKey = CompareHeaderKey(CStr(alias))
        If headerMap.Exists(headerKey) Then
            colNo = CLng(headerMap(headerKey))
            ' Zero denotes duplicate normalized headers: keep the original fail-closed behavior.
            If colNo = 0 Then Exit Function
            matchedColumns(CStr(colNo)) = True
        End If
    Next alias
    If matchedColumns.Count = 1 Then
        For Each columnKey In matchedColumns.Keys
            FindCompareColumn = CLng(columnKey)
        Next columnKey
    End If
End Function

Private Sub BeginEDisplayRead()
    Set eDisplayHeaderMaps = CreateObject("Scripting.Dictionary")
    Set eDisplayNormalizedKeys = CreateObject("Scripting.Dictionary")
    eDisplayHeaderMaps.CompareMode = vbBinaryCompare
    eDisplayNormalizedKeys.CompareMode = vbBinaryCompare
End Sub

Private Sub EndEDisplayRead()
    Set eDisplayHeaderMaps = Nothing
    Set eDisplayNormalizedKeys = Nothing
End Sub

Private Function EDisplayHeaderMap(ByVal ws As Worksheet) As Object
    Dim headerMap As Object, values As Variant, rawValue As Variant
    Dim lastCol As Long, colNo As Long, headerKey As String
    If Not eDisplayHeaderMaps Is Nothing Then
        If eDisplayHeaderMaps.Exists(ws.Name) Then
            Set EDisplayHeaderMap = eDisplayHeaderMaps(ws.Name)
            Exit Function
        End If
    End If
    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = vbBinaryCompare
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    values = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value2
    For colNo = 1 To lastCol
        If lastCol = 1 Then
            rawValue = values
        Else
            rawValue = values(1, colNo)
        End If
        If Not IsError(rawValue) Then
            headerKey = CompareHeaderKey(CStr(rawValue))
            If headerMap.Exists(headerKey) Then
                headerMap(headerKey) = 0
            Else
                headerMap.Add headerKey, colNo
            End If
        End If
    Next colNo
    If Not eDisplayHeaderMaps Is Nothing Then eDisplayHeaderMaps.Add ws.Name, headerMap
    Set EDisplayHeaderMap = headerMap
End Function

Private Function ESelectionRenderKey() As String
    Dim i As Long, result As String
    result = CStr(mCboEVisibility.ListIndex) & ":" & CStr(mLstCases.ListIndex) & ":"
    For i = 0 To mLstCases.ListCount - 1
        If mLstCases.Selected(i) Then result = result & CStr(i) & ";"
    Next i
    ESelectionRenderKey = result
End Function

Private Sub EDisplayFailed(ByVal failureText As String)
    EndEDisplayRead
    eRenderedSelection = ""
    rendering = False
    mCmdApprove.Enabled = False
    mCmdOther.Enabled = False
    mCmdHold.Enabled = False
    mLblStatus.Caption = "表示の更新に失敗しました。採用せず、再判定またはフォームの再読込を行ってください。"
    mTxtDetails.Text = failureText
End Sub

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
    Dim conflictFieldItems As Collection, conflictField As Object, conflictDetails As String
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
    If CStr(item("category")) = "E" Then
        conflictDetails = "【固定管理行ID】" & CStr(item("rowId")) & vbCrLf & _
            "【項目グループ】" & CStr(item("groupName")) & vbCrLf & _
            "【判断】" & CStr(item("decision"))
        Set conflictFieldItems = item("headers")
        For Each conflictField In conflictFieldItems
            conflictDetails = conflictDetails & vbCrLf & vbCrLf & "項目: " & CStr(conflictField("sourceHeader")) & vbCrLf & _
                "前回同期値: " & IIf(Len(CStr(conflictField("baselineValue"))) = 0, "（空欄）", CStr(conflictField("baselineValue"))) & vbCrLf & _
                "システム側: " & IIf(Len(CStr(conflictField("systemValue"))) = 0, "（空欄）", CStr(conflictField("systemValue"))) & vbCrLf & _
                "Excel側: " & IIf(Len(CStr(conflictField("excelValue"))) = 0, "（空欄）", CStr(conflictField("excelValue")))
        Next conflictField
        mTxtDetails.Text = conflictDetails
        Exit Sub
    End If
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
    Dim i As Long, item As Object, category As String, counts(1 To 5) As Long, unresolved(1 To 5) As Long
    Dim decision As String, displayDecision As String, idx As Long, categoryIndex As Long, isD As Boolean
    Dim filteredUnresolved As Long, filteredHeld As Long, selectedGroup As String
    If cases Is Nothing Then Exit Sub
    categoryIndex = mCboCategory.ListIndex
    If categoryIndex = 4 Then
        ShowECategory
        Exit Sub
    End If
    LayoutReviewForE False
    rendering = True
    Set shown = New Collection
    If categoryIndex = 4 Then
        mLstCases.MultiSelect = fmMultiSelectExtended
    Else
        mLstCases.MultiSelect = fmMultiSelectSingle
    End If
    mLstCases.Clear: mLstRows.Clear
    mCboEGroup.Visible = (categoryIndex = 4)
    mLblESelection.Visible = (categoryIndex = 4)
    mCmdEClear.Visible = (categoryIndex = 4)
    selectedGroup = "すべて"
    If categoryIndex = 4 And mCboEGroup.ListIndex >= 0 Then selectedGroup = CStr(mCboEGroup.Value)
    mLblCase.Caption = IIf(categoryIndex = 4, "E案件一覧（Ctrl/Shiftで複数選択）", "案件一覧（1件ずつ確認）")
    For i = 1 To cases.Count
        Set item = cases(i)
        category = CStr(item("category"))
        If category = "E" And CLng(item("year")) <> targetYear Then GoTo NextCase
        Select Case category
            Case "A": idx = 1
            Case "B": idx = 2
            Case "C": idx = 3
            Case "D": idx = 4
            Case "E": idx = 5
            Case Else: GoTo NextCase
        End Select
        counts(idx) = counts(idx) + 1
        isD = (category = "D")
        decision = ""
        If isD Then
            unresolved(idx) = unresolved(idx) + item("copyCandidateRows").Count
        Else
            decision = SyncReviewDecision(item)
            If Len(decision) = 0 Or (category = "E" And decision = "保留") Then unresolved(idx) = unresolved(idx) + 1
        End If
        If categoryIndex = idx - 1 And _
           (category <> "E" Or mCboEGroup.ListIndex <= 0 Or _
            StrComp(CStr(item("groupName")), selectedGroup, vbTextCompare) = 0) Then
            shown.Add item
            If category = "E" Then
                If Len(decision) = 0 Then filteredUnresolved = filteredUnresolved + 1
                If decision = "保留" Then filteredHeld = filteredHeld + 1
            End If
            If isD Then
                mLstCases.AddItem "CSV行" & CStr(item("sourceRow")) & " / 保持Excel " & _
                    CStr(item("copyKeeperRow")) & " / クリア候補 " & CStr(item("copyCandidateRows").Count) & "件"
            ElseIf category = "E" Then
                displayDecision = CStr(item("decision"))
                If Len(displayDecision) = 0 Then displayDecision = "未解決"
                mLstCases.AddItem "CSV行" & CStr(item("sourceRow")) & " / Excel " & CStr(item("rows")) & _
                    " / " & CStr(item("groupName")) & " [" & displayDecision & "]"
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
        " / D " & counts(4) & "件（候補 " & unresolved(4) & "行）" & _
        " / E " & counts(5) & "件（未解決/保留 " & unresolved(5) & "）"
    If categoryIndex = 4 Then
        mLblCounts.Caption = "対象年度 " & targetYear & " / " & selectedGroup & "：" & _
            CStr(shown.Count) & "件表示／未解決" & CStr(filteredUnresolved) & "件／保留" & _
            CStr(filteredHeld) & "件（E全体 " & CStr(counts(5)) & "件）"
    End If
    mTxtDetails.Text = "候補を選んでください。"
    mChkDelete.Value = False
    If categoryIndex <> 4 And mLstCases.ListCount > 0 Then mLstCases.ListIndex = 0
    rendering = False
    ShowCurrent
End Sub

Private Sub ShowCurrent()
    Dim item As Object, v As Variant, rows As String, category As String
    Dim candidates As Collection, candidate As Variant, keeper As Long, rowNo As Long
    Dim decision As String, rowDecision As String
    If mCboCategory.ListIndex = 4 Then
        ShowECurrent
        Exit Sub
    End If
    Set item = SelectedCase()
    rendering = True
    mCmdApprove.Caption = "同一として処理"
    mCmdOther.Caption = "別物"
    mCmdHold.Caption = "保留"
    mLstRows.Clear
    mChkDelete.Value = False
    If item Is Nothing Then
        mTxtDetails.Text = "このカテゴリには候補がありません。"
        RenderComparison Nothing, 0
        mCmdApprove.Enabled = False: mCmdGoTo.Enabled = False: mCmdClear.Enabled = False
        mCmdOther.Enabled = False: mCmdHold.Enabled = False: mCmdDelete.Enabled = False
        mChkDelete.Visible = False
        rendering = False
        UpdateESelection
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
    mCmdGoTo.Enabled = (category <> "D" And category <> "E" And rowNo >= 2)
    mCmdApprove.Enabled = (category = "E" Or (category = "B" And mLstRows.ListCount = 1))
    mCmdOther.Enabled = (category <> "D")
    mCmdHold.Enabled = (category <> "D")
    mCmdDelete.Enabled = (category <> "D" And category <> "E")
    mCmdClear.Enabled = IsCopyCandidate(item, rowNo)
    If Len(mDUnavailableReason) > 0 Then mCmdClear.Enabled = False
    mChkDelete.Visible = (category = "C")
    If mChkDelete.Visible Then mChkDelete.Value = (Len(decision) = 0 And deletionSelection.Exists(CStr(item("key"))))
    If category = "D" Then
        mLblStatus.Caption = "D候補：保持行はクリアできません。選択した候補行のみ専用APIで再検証します。"
    ElseIf category = "B" Then
        UpdateBRowJudgments item
    ElseIf category = "E" Then
        mLblRows.Caption = "固定管理行IDで特定したExcel行"
        mLblStatus.Caption = "承認前は変更しません。システム採用は確認後にExcelへ反映します。"
        mCmdApprove.Caption = "文書管理を採用"
        mCmdOther.Caption = "Excel基準表を採用"
        mCmdHold.Caption = "保留"
    Else
        mLblRows.Caption = "候補のExcel行"
        mLblStatus.Caption = "Aの重複ID・複数行の統合は自動承認不可。Cはチェック後に件数を確認します。"
    End If
    rendering = False
    RenderComparison item, rowNo
    UpdateESelection
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
Private Sub mLstCases_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal X As Single, ByVal Y As Single)
    If Not rendering Then eRenderedSelection = ""
End Sub

Private Sub mLstCases_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not rendering Then eRenderedSelection = ""
End Sub

Private Sub mLstCases_Click()
    If Not rendering Then ShowCurrent
End Sub

' In an extended-selection ListBox, changing the selected rows need not change Value.
' Click alone therefore does not cover Ctrl/Shift selections or keyboard navigation.
Private Sub mLstCases_Change()
    If Not rendering And mCboCategory.ListIndex = 4 Then ShowCurrent
End Sub

Private Sub mLstCases_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single)
    If Not rendering And mCboCategory.ListIndex = 4 Then ShowCurrent
End Sub

Private Sub mLstCases_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not rendering And mCboCategory.ListIndex = 4 Then ShowCurrent
End Sub

Private Sub mLstCases_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not rendering And mCboCategory.ListIndex = 4 Then UpdateESelection
End Sub
Private Sub mLstRows_Click()
    Dim item As Object
    If rendering Then Exit Sub
    If mCboCategory.ListIndex = 4 Then
        UpdateEFieldSelection
        Exit Sub
    End If
    Set item = SelectedCase()
    If Not item Is Nothing Then
        If CStr(item("category")) = "C" Then SelectDeleteRow item, SelectedRow()
    End If
    ShowCurrentSelection
End Sub

Private Sub mLstRows_Change()
    If Not rendering And mCboCategory.ListIndex = 4 Then UpdateEFieldSelection
End Sub

Private Sub mLstRows_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not rendering And mCboCategory.ListIndex = 4 Then UpdateEFieldSelection
End Sub

Private Sub ShowCurrentSelection()
    Dim item As Object, rowNo As Long
    If mCboCategory.ListIndex = 4 Then
        UpdateEFieldSelection
        Exit Sub
    End If
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
    Dim item As Object, reasonText As String
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If CStr(item("category")) = "E" Then
        If ESelectedCount() <> 1 Then Exit Sub
        If Not ESelectedDecisionIsSafe("システム採用", reasonText) Then
            If Len(reasonText) > 0 Then mLblStatus.Caption = reasonText
            Exit Sub
        End If
        If EIsClassification1Group(item) Then
            If Not EConfirmClassification1Decision(item, SelectedRecord(), "システム採用", reasonText) Then
                If Len(reasonText) > 0 Then mLblStatus.Caption = reasonText
                Exit Sub
            End If
        ElseIf EIsRetentionGroup(item) Then
            If Not EConfirmRetentionDecision(item, SelectedRecord(), "システム採用", reasonText) Then
                If Len(reasonText) > 0 Then mLblStatus.Caption = reasonText
                Exit Sub
            End If
        ElseIf MsgBox("文書管理システムの値をExcel基準表へ反映します。選択した関連項目グループ全体が対象です。続けますか？", _
                      vbQuestion + vbYesNo, "競合解決の確認") <> vbYes Then
            Exit Sub
        End If
        If Module1.BidirectionalSyncResolveConflict(item, "システム採用") Then mCmdRefresh_Click
        Exit Sub
    End If
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
    Dim item As Object, rowNo As Long, nextRow As Long, reasonText As String
    Set item = SelectedCase()
    If item Is Nothing Then Exit Sub
    If CStr(item("category")) = "E" Then
        If decision = "別物" Then decision = "Excel採用"
        If decision <> "Excel採用" And decision <> "保留" Then Exit Sub
        If ESelectedCount() <> 1 Then Exit Sub
        If Not ESelectedDecisionIsSafe(decision, reasonText) Then
            If Len(reasonText) > 0 Then mLblStatus.Caption = reasonText
            Exit Sub
        End If
        If EIsClassification1Group(item) Then
            If Not EConfirmClassification1Decision(item, SelectedRecord(), decision, reasonText) Then
                If Len(reasonText) > 0 Then mLblStatus.Caption = reasonText
                Exit Sub
            End If
        ElseIf EIsRetentionGroup(item) Then
            If Not EConfirmRetentionDecision(item, SelectedRecord(), decision, reasonText) Then
                If Len(reasonText) > 0 Then mLblStatus.Caption = reasonText
                Exit Sub
            End If
        End If
        If Module1.BidirectionalSyncResolveConflict(item, decision) Then
            ShowCategory
            mCmdNext_Click
        End If
        Exit Sub
    End If
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
Private Sub SelectSingleECase(ByVal newIndex As Long)
    Dim n As Long
    eRenderedSelection = ""
    rendering = True
    For n = 0 To mLstCases.ListCount - 1: mLstCases.Selected(n) = False: Next n
    mLstCases.Selected(newIndex) = True
    mLstCases.ListIndex = newIndex
    rendering = False
    ShowCurrent
End Sub

Private Sub mCmdNext_Click()
    Dim i As Long, startAt As Long, item As Object, candidate As Variant
    Dim candidates As Collection, selected As Long, j As Long, rowNo As Long
    If mCboCategory.ListIndex = 4 Then
        SelectNextECase
        Exit Sub
    End If
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
            If mCboCategory.ListIndex = 4 Then
                SelectSingleECase (startAt + i) Mod shown.Count
            Else
                mLstCases.ListIndex = (startAt + i) Mod shown.Count
                ShowCurrent
            End If
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
        If CStr(item("category")) = "D" Or CStr(item("category")) = "E" Then Exit Sub
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
    Dim refreshed As Collection, copyCases As Collection, conflictCases As Collection
    Dim copyItem As Object, conflictItem As Object, previousCases As Collection
    Dim unavailableReason As String, detectorNumber As Long
    Dim detectorDescription As String, detectorSource As String, newSelection As Object
    Dim previousSelection As Object, previousReason As String, failureText As String
    Dim previousCategory As Long, committed As Boolean
    eRenderedSelection = ""
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
    Set conflictCases = Module1.BidirectionalSyncConflictReviewCases(targetYear)
    For Each conflictItem In conflictCases
        refreshed.Add conflictItem
    Next conflictItem
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
    PopulateEGroups
    Set deletionSelection = newSelection
    mDUnavailableReason = unavailableReason
    UpdateDWarning
    If Len(mDUnavailableReason) > 0 And mCboCategory.ListIndex = 3 Then
        rendering = True: mCboCategory.ListIndex = 0: rendering = False
    End If
    If mCboCategory.ListIndex = 4 And Not HasReviewCategory("E") Then
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
        PopulateEGroups
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

Private Sub LayoutReviewForE(ByVal isE As Boolean)
    Dim i As Long, targetInsideHeight As Single, availableWidth As Single
    Dim leftWidth As Single, rightLeft As Single, rightWidth As Single
    Dim caseColumnArea As Single, guideColumnWidth As Long, folderColumnWidth As Long, yearColumnWidth As Long
    Dim itemNameWidth As Long, decisionWidth As Long, valueWidth As Long
    Dim detailsHeight As Single, footerTop As Single, detailsTop As Single, listHeight As Single
    Dim headerLeft As Single, columnWidths(0 To 4) As Single
    If isE Then
        Select Case mCboEWindowSize.ListIndex
            Case 0
                Me.Width = 960
                targetInsideHeight = 400
            Case 2
                Me.Width = 1280
                targetInsideHeight = 560
            Case Else
                Me.Width = 1100
                targetInsideHeight = 440
        End Select
        Me.Height = Me.Height + (targetInsideHeight - Me.InsideHeight)
        availableWidth = Me.InsideWidth
        Select Case mCboEPaneSize.ListIndex
            Case 1
                leftWidth = 320
            Case 2
                leftWidth = 205
            Case Else
                leftWidth = 245
        End Select
        If leftWidth > availableWidth - 575 Then leftWidth = availableWidth - 575
        If leftWidth < 170 Then leftWidth = 170
        rightLeft = 9 + leftWidth + 9
        rightWidth = availableWidth - rightLeft - 9
        caseColumnArea = leftWidth - 16
        guideColumnWidth = CLng(caseColumnArea * 0.28)
        folderColumnWidth = CLng(caseColumnArea * 0.56)
        yearColumnWidth = CLng(caseColumnArea) - guideColumnWidth - folderColumnWidth
        If yearColumnWidth < 55 Then
            folderColumnWidth = folderColumnWidth - (55 - yearColumnWidth)
            yearColumnWidth = 55
        End If
        itemNameWidth = CLng(rightWidth * 0.23)
        If itemNameWidth < 130 Then itemNameWidth = 130
        If itemNameWidth > 260 Then itemNameWidth = 260
        decisionWidth = CLng(rightWidth * 0.18)
        If decisionWidth < 130 Then decisionWidth = 130
        If decisionWidth > 180 Then decisionWidth = 180
        valueWidth = CLng((rightWidth - itemNameWidth - decisionWidth - 16) / 3)
        If valueWidth < 90 Then valueWidth = 90
        columnWidths(0) = itemNameWidth
        columnWidths(1) = valueWidth
        columnWidths(2) = valueWidth
        columnWidths(3) = valueWidth
        columnWidths(4) = decisionWidth
        PlaceControl mLblCounts, 9, 6, availableWidth - 18, 34
        mLblCounts.WordWrap = True
        PlaceControl mCboCategory, 9, 43, 187.5, 16.5
        mLblRows.Caption = "項目群"
        PlaceControl mLblRows, 205, 43, 42, 16.5
        mLblRows.Visible = True
        PlaceControl mCboEGroup, 251, 43, 160, 16.5
        mCboEGroup.Visible = True
        PlaceControl mLblEVisibility, 419, 43, 42, 16.5
        PlaceControl mCboEVisibility, 464, 43, 112, 16.5
        PlaceControl mLblEWindowSize, 582, 43, 35, 16.5
        PlaceControl mCboEWindowSize, 620, 43, 88, 16.5
        PlaceControl mLblEPaneSize, 716, 43, 28, 16.5
        PlaceControl mCboEPaneSize, 748, 43, 112, 16.5
        mLblEVisibility.Visible = True
        mCboEVisibility.Visible = True
        mLblEWindowSize.Visible = True
        mCboEWindowSize.Visible = True
        mLblEPaneSize.Visible = True
        mCboEPaneSize.Visible = True
        PlaceControl mLblCase, 9, 65, leftWidth, 14
        mLblCase.Caption = "確認が必要な入力項目"
        PlaceControl mLstCases, 9, 79.5, leftWidth, 228
        mLstCases.ColumnCount = 3
        mLstCases.ColumnWidths = CStr(guideColumnWidth) & " pt;" & CStr(folderColumnWidth) & _
                                 " pt;" & CStr(yearColumnWidth) & " pt"
        mLblCompareRows.Visible = False
        PlaceControl mLstRows, rightLeft, 79.5, rightWidth, 228
        mLstRows.ColumnCount = 5
        mLstRows.ColumnWidths = CStr(CLng(columnWidths(0))) & " pt;" & CStr(CLng(columnWidths(1))) & _
            " pt;" & CStr(CLng(columnWidths(2))) & " pt;" & CStr(CLng(columnWidths(3))) & _
            " pt;" & CStr(CLng(columnWidths(4))) & " pt"
        mLstRows.MultiSelect = fmMultiSelectSingle
        headerLeft = rightLeft
        For i = 0 To 4
            PlaceControl mEColumnHeaders(i), headerLeft, 65, columnWidths(i), 14
            headerLeft = headerLeft + columnWidths(i)
            mEColumnHeaders(i).Visible = True
        Next i
        If targetInsideHeight > 440 Then
            detailsHeight = 75
        Else
            detailsHeight = 55
        End If
        footerTop = Me.InsideHeight - 84
        detailsTop = footerTop - detailsHeight - 2
        listHeight = detailsTop - 85.5
        If listHeight < 150 Then listHeight = 150
        PlaceControl mLstCases, 9, 79.5, leftWidth, listHeight
        PlaceControl mLstRows, rightLeft, 79.5, rightWidth, listHeight
        PlaceControl mTxtDetails, rightLeft, detailsTop, rightWidth, detailsHeight
        PlaceControl mLblESelection, 9, detailsTop + 2.5, leftWidth - 120, 19
        PlaceControl mCmdEClear, 9 + leftWidth - 110, detailsTop + 2.5, 108, 22
        PlaceControl mLblStatus, 9, footerTop, availableWidth - 18, 21
        PlaceControl mLblDWarning, 9, footerTop + 24, availableWidth - 18, 21
        mLblDWarning.Visible = False
        mFraCompare.Visible = False
        mLblESelection.Visible = True
        mCmdEClear.Visible = True
        mCmdGoTo.Visible = False
        mCmdDelete.Visible = False
        mCmdClear.Visible = False
        PlaceControl mCmdApprove, 9, footerTop + 23.5, 110, 21
        mCmdApprove.Caption = "文書管理を採用"
        PlaceControl mCmdOther, 124, footerTop + 23.5, 125, 21
        mCmdOther.Caption = "Excel基準表を採用"
        PlaceControl mCmdHold, 254, footerTop + 23.5, 55, 21
        mCmdHold.Caption = "保留"
        PlaceControl mCmdNext, 314, footerTop + 23.5, 100, 21
        mCmdNext.Caption = "次の未解決項目"
        PlaceControl mCmdRefresh, 419, footerTop + 23.5, 75, 21
        PlaceControl mCmdFinish, 499, footerTop + 23.5, 82, 21
    Else
        Me.Width = 760
        Me.Height = Me.Height + (440 - Me.InsideHeight)
        mLblEVisibility.Visible = False
        mCboEVisibility.Visible = False
        mLblEWindowSize.Visible = False
        mCboEWindowSize.Visible = False
        mLblEPaneSize.Visible = False
        mCboEPaneSize.Visible = False
        PlaceControl mLblCounts, 9, 6, 740, 19
        PlaceControl mCboCategory, 9, 28, 187.5, 16.5
        PlaceControl mCboEGroup, 205, 28, 110, 16.5
        PlaceControl mLblCase, 9, 50, 187.5, 14
        PlaceControl mLstCases, 9, 64.5, 187.5, 228
        mLstCases.ColumnCount = 1
        mLstCases.ColumnWidths = ""
        PlaceControl mLblRows, 205, 50, 110, 14
        mLblRows.Caption = "候補のExcel行"
        mLblRows.Visible = True
        PlaceControl mLstRows, 205, 64.5, 110, 228
        mLstRows.ColumnCount = 1
        mLstRows.ColumnWidths = ""
        mLstRows.MultiSelect = fmMultiSelectSingle
        PlaceControl mLblCompareRows, 324, 28, 420, 18
        mLblCompareRows.Visible = True
        PlaceControl mFraCompare, 324, 48, 420, 225
        mFraCompare.Visible = True
        PlaceControl mTxtDetails, 324, 278, 420, 50
        PlaceControl mLblStatus, 9, 333, 740, 21
        PlaceControl mLblDWarning, 9, 356, 740, 21
        mLblDWarning.Caption = "警告：D利用不可: " & mDUnavailableReason
        mLblDWarning.ControlTipText = mDUnavailableReason
        mLblDWarning.Visible = (Len(mDUnavailableReason) > 0)
        PlaceControl mLblESelection, 9, 301.5, 190, 19
        PlaceControl mCmdEClear, 205, 300, 110, 22
        mLblESelection.Visible = False
        mCmdEClear.Visible = False
        mCmdGoTo.Visible = True
        PlaceControl mCmdGoTo, 9, 379.5, 70, 21
        mCmdGoTo.Caption = "対象行へ移動"
        mCmdApprove.Visible = True
        PlaceControl mCmdApprove, 84, 379.5, 78, 21
        mCmdApprove.Caption = "同一として処理"
        mCmdOther.Visible = True
        PlaceControl mCmdOther, 167, 379.5, 50, 21
        mCmdOther.Caption = "別物"
        mCmdHold.Visible = True
        PlaceControl mCmdHold, 222, 379.5, 50, 21
        mCmdHold.Caption = "保留"
        mCmdNext.Visible = True
        PlaceControl mCmdNext, 277, 379.5, 60, 21
        mCmdNext.Caption = "次の候補"
        mCmdDelete.Visible = True
        PlaceControl mCmdDelete, 342, 379.5, 85.5, 21
        mCmdDelete.Caption = "選択行を一括削除"
        mCmdClear.Visible = True
        PlaceControl mCmdClear, 433, 379.5, 125, 21
        mCmdClear.Caption = "選択行の番号をクリア"
        PlaceControl mCmdRefresh, 9, 407, 70, 21
        PlaceControl mCmdFinish, 508, 407, 82, 21
        For i = 0 To 4
            mEColumnHeaders(i).Visible = False
        Next i
    End If
End Sub

Private Sub ShowECategory(Optional ByVal preserveFileSelection As Boolean = False)
    Dim allRecords As Collection, record As Object, recordCases As Collection, item As Object
    Dim selectedGroup As String, filtered As Boolean, selectedRecordKeys As Object
    Dim priorCase As Object, priorField As Object, priorRecord As Object
    Dim existingItem As Object, existingField As Object
    Dim priorCaseKey As String, priorRecordKey As String, priorHeaderName As String
    Dim recordKey As String, existingRecordKey As String, priorTopIndex As Long
    Dim rowIndex As Long, i As Long, guide2Value As String, folderValue As String, creationYearValue As String
    Dim totalItems As Long, inputItemCount As Long, internalMismatchCount As Long
    Dim unresolvedInternalCount As Long, safeNoReviewCount As Long
    Dim internalOnlyFileCount As Long, identityIssueFileCount As Long
    Dim inputFileKeys As Object
    On Error GoTo DisplayFailed
    eRenderedSelection = ""
    If cases Is Nothing Then Exit Sub
    BeginEDisplayRead
    If preserveFileSelection Then
        If mLstCases.ListCount > 0 Then priorTopIndex = mLstCases.TopIndex
        Set selectedRecordKeys = CreateObject("Scripting.Dictionary")
        selectedRecordKeys.CompareMode = vbBinaryCompare
        For i = 0 To mLstCases.ListCount - 1
            If mLstCases.Selected(i) And Not shown Is Nothing Then
                If i + 1 <= shown.Count Then
                    Set record = shown(i + 1)
                    recordKey = ERecordSelectionKey(record)
                    If Len(recordKey) > 0 Then selectedRecordKeys(recordKey) = True
                End If
            End If
        Next i
        If ESelectedCount() = 1 Then
            Set priorCase = SelectedCase()
            Set priorField = SelectedEField()
            Set priorRecord = SelectedRecord()
            priorRecordKey = ERecordSelectionKey(priorRecord)
            If Len(priorRecordKey) > 0 And Not priorCase Is Nothing And Not priorField Is Nothing Then
                priorCaseKey = CStr(priorCase("key"))
                priorHeaderName = CStr(priorField("sourceHeader"))
            End If
        End If
    End If
    rendering = True
    LayoutReviewForE True
    Set allRecords = BuildERecords(cases, targetYear)
    Set shown = New Collection
    Set eFieldCases = New Collection
    Set eFieldItems = New Collection
    Set inputFileKeys = CreateObject("Scripting.Dictionary")
    inputFileKeys.CompareMode = vbBinaryCompare
    ' This stage does not auto-exclude any internal conflict; all such differences remain unresolved.
    safeNoReviewCount = 0
    mLstCases.MultiSelect = fmMultiSelectExtended
    mLstCases.Clear
    mLstRows.Clear
    mLstRows.ListIndex = -1
    mCboEGroup.Visible = True
    mLblCase.Caption = "確認が必要な入力項目"
    mLblRows.Visible = True
    mLblESelection.Caption = "0件選択中"
    mLblCompareRows.Caption = "ファイルを選択してください。"
    mTxtDetails.Text = "［入力項目のみ］では職員が確認する入力項目を表示します。内部項目だけの競合や識別不明は［全項目（管理コード含む）］で確認できます。"
    mLblStatus.Caption = "複数ファイルの選択は表示専用です。判断時は識別情報が一意なファイル1件を選んでください。"
    mChkDelete.Visible = False
    mCmdApprove.Caption = "文書管理を採用"
    mCmdOther.Caption = "Excel基準表を採用"
    mCmdHold.Caption = "保留"
    mCmdApprove.Enabled = False
    mCmdOther.Enabled = False
    mCmdHold.Enabled = False
    mCmdGoTo.Enabled = False
    mCmdClear.Enabled = False
    mCmdEClear.Enabled = False
    filtered = (mCboEGroup.ListIndex > 0)
    If filtered Then selectedGroup = CStr(mCboEGroup.Value)
    For Each record In allRecords
        Set recordCases = record("cases")
        totalItems = totalItems + recordCases.Count
        If CBool(record("identityIssue")) Then
            identityIssueFileCount = identityIssueFileCount + 1
        ElseIf ERecordNeedsConfirmation(record) Then
            recordKey = ERecordBucketKey(record)
            If Not inputFileKeys.Exists(recordKey) Then inputFileKeys.Add recordKey, True
        ElseIf ERecordNeedsInternalInvestigation(record) Then
            internalOnlyFileCount = internalOnlyFileCount + 1
        End If
        For Each item In recordCases
            If ECaseNeedsHumanDecision(item) Then
                If Not CBool(record("identityIssue")) And ECaseHasNormalDisplayFields(item) Then _
                    inputItemCount = inputItemCount + 1
                ECountInternalCase item, internalMismatchCount, unresolvedInternalCount
            End If
        Next item
        If (EIsAllDisplayMode() Or ERecordNeedsConfirmation(record)) And ERecordHasGroup(record, selectedGroup) Then
            shown.Add record
            ERecordDisplayValues record, guide2Value, folderValue, creationYearValue
            rowIndex = mLstCases.ListCount
            mLstCases.AddItem guide2Value
            mLstCases.List(rowIndex, 1) = folderValue
            mLstCases.List(rowIndex, 2) = creationYearValue
            If Not CBool(record("identityIssue")) And Not ERecordNeedsConfirmation(record) And _
               ERecordNeedsInternalInvestigation(record) Then
                mLstCases.List(rowIndex, 0) = "内部のみ・要調査: " & CStr(mLstCases.List(rowIndex, 0))
            End If
            recordKey = ERecordSelectionKey(record)
            If preserveFileSelection And Len(recordKey) > 0 And Not selectedRecordKeys Is Nothing Then
                If selectedRecordKeys.Exists(recordKey) Then mLstCases.Selected(rowIndex) = True
            End If
        End If
    Next record
    mLblCounts.Caption = "確認が必要なファイル: " & CStr(inputFileKeys.Count) & "件 / 確認が必要な入力項目: " & _
        CStr(inputItemCount) & "件" & vbCrLf & _
        "内部項目の不一致: " & CStr(internalMismatchCount) & "件 / 安全に確認不要: " & _
        CStr(safeNoReviewCount) & "件 / 未確認・要調査: " & CStr(unresolvedInternalCount) & _
        "件 / 内部のみ: " & CStr(internalOnlyFileCount) & "ファイル / 識別要確認: " & _
        CStr(identityIssueFileCount) & "ファイル"
    If filtered Then mLblStatus.Caption = mLblStatus.Caption & " 絞り込み後 " & CStr(shown.Count) & "ファイル／全競合項目 " & CStr(totalItems) & "件。"
    If shown.Count = 0 Then
        If filtered Then
            mTxtDetails.Text = "この絞り込み条件に該当するファイルはありません。"
            mLblStatus.Caption = "絞り込みを変更するか、［すべて］を選択してください。"
        ElseIf EIsAllDisplayMode() Then
            mTxtDetails.Text = "この絞り込み条件に該当するファイルはありません。"
            mLblStatus.Caption = "絞り込みを変更するか、［すべて］を選択してください。"
        Else
            mTxtDetails.Text = "確認が必要な入力項目はありません。内部項目の未確認分は件数に含めており、［全項目（管理コード含む）］で確認できます。"
            mLblStatus.Caption = "職員入力の確認待ちはありません。内部項目の調査が必要な場合は［全項目（管理コード含む）］に切り替えてください。"
        End If
    End If
    If preserveFileSelection And mLstCases.ListCount > 0 Then
        If priorTopIndex < 0 Then priorTopIndex = 0
        If priorTopIndex > mLstCases.ListCount - 1 Then priorTopIndex = mLstCases.ListCount - 1
        mLstCases.TopIndex = priorTopIndex
    End If
    EndEDisplayRead
    rendering = False
    If preserveFileSelection Then
        ShowCurrent
        If Len(priorCaseKey) > 0 And Len(priorRecordKey) > 0 And Len(priorHeaderName) > 0 Then
            For i = 1 To eFieldCases.Count
                Set existingItem = eFieldCases(i)
                Set existingField = eFieldItems(i)
                existingRecordKey = ERecordSelectionKey(existingItem)
                If StrComp(existingRecordKey, priorRecordKey, vbBinaryCompare) = 0 And _
                   StrComp(CStr(existingItem("key")), priorCaseKey, vbBinaryCompare) = 0 And _
                   StrComp(CompareHeaderKey(CStr(existingField("sourceHeader"))), _
                           CompareHeaderKey(priorHeaderName), vbBinaryCompare) = 0 Then
                    rendering = True
                    mLstRows.ListIndex = i - 1
                    rendering = False
                    UpdateEFieldSelection
                    Exit For
                End If
            Next i
        End If
    Else
        UpdateESelection
    End If
    Exit Sub
DisplayFailed:
    EDisplayFailed Err.Description
End Sub

Private Sub ECountInternalCase(ByVal caseItem As Object, ByRef internalMismatchCount As Long, _
                               ByRef unresolvedInternalCount As Long)
    Dim fields As Collection, fieldItem As Object
    Dim systemValue As String, excelValue As String
    Set fields = caseItem("headers")
    For Each fieldItem In fields
        If Len(EFieldNormalDisplayCaption(caseItem, fieldItem)) = 0 Then
            systemValue = CStr(fieldItem("systemValue"))
            excelValue = CStr(fieldItem("excelValue"))
            If StrComp(systemValue, excelValue, vbBinaryCompare) <> 0 Then
                internalMismatchCount = internalMismatchCount + 1
                unresolvedInternalCount = unresolvedInternalCount + 1
            End If
        End If
    Next fieldItem
End Sub

Private Function BuildERecords(ByVal sourceCases As Collection, ByVal yearValue As Long) As Collection
    Dim result As New Collection, recordIndexByKey As Object
    Dim rowIdCounts As Object, stableKeyCounts As Object
    Dim item As Object, record As Object, recordCases As Collection
    Dim i As Long, bucketKey As String, rowKey As String, stableKey As String
    Set recordIndexByKey = CreateObject("Scripting.Dictionary")
    recordIndexByKey.CompareMode = vbBinaryCompare
    Set rowIdCounts = CreateObject("Scripting.Dictionary")
    rowIdCounts.CompareMode = vbBinaryCompare
    Set stableKeyCounts = CreateObject("Scripting.Dictionary")
    stableKeyCounts.CompareMode = vbBinaryCompare
    If sourceCases Is Nothing Then
        Set BuildERecords = result
        Exit Function
    End If
    For i = 1 To sourceCases.Count
        Set item = sourceCases(i)
        If CStr(item("category")) = "E" And CLng(item("year")) = yearValue Then
            bucketKey = ERecordBucketKey(item)
            If Len(bucketKey) = 0 Then bucketKey = "UNVERIFIED:" & CStr(i)
            If recordIndexByKey.Exists(bucketKey) Then
                Set record = result(CLng(recordIndexByKey(bucketKey)))
            Else
                Set record = CreateObject("Scripting.Dictionary")
                record.Add "year", CLng(item("year"))
                record.Add "rowId", CStr(item("rowId"))
                record.Add "stableKey", CStr(item("stableKey"))
                record.Add "sourceRow", CLng(item("sourceRow"))
                record.Add "destinationRow", CLng(item("destinationRow"))
                record.Add "identityIssue", (Left$(bucketKey, 11) = "UNVERIFIED:")
                Set recordCases = New Collection
                record.Add "cases", recordCases
                result.Add record
                If Left$(bucketKey, 11) <> "UNVERIFIED:" Then recordIndexByKey.Add bucketKey, result.Count
            End If
            Set recordCases = record("cases")
            recordCases.Add item
        End If
    Next i
    For Each record In result
        rowKey = ERecordRowIdKey(record)
        If Len(rowKey) > 0 Then
            If rowIdCounts.Exists(rowKey) Then
                rowIdCounts(rowKey) = CLng(rowIdCounts(rowKey)) + 1
            Else
                rowIdCounts.Add rowKey, 1
            End If
        End If
        stableKey = ERecordStableKeyKey(record)
        If Len(stableKey) > 0 Then
            If stableKeyCounts.Exists(stableKey) Then
                stableKeyCounts(stableKey) = CLng(stableKeyCounts(stableKey)) + 1
            Else
                stableKeyCounts.Add stableKey, 1
            End If
        End If
    Next record
    For Each record In result
        rowKey = ERecordRowIdKey(record)
        stableKey = ERecordStableKeyKey(record)
        If Len(rowKey) > 0 And rowIdCounts.Exists(rowKey) Then
            If CLng(rowIdCounts(rowKey)) > 1 Then record("identityIssue") = True
        End If
        If Len(stableKey) > 0 And stableKeyCounts.Exists(stableKey) Then
            If CLng(stableKeyCounts(stableKey)) > 1 Then record("identityIssue") = True
        End If
    Next record
    Set BuildERecords = result
End Function

Private Function ERecordBucketKey(ByVal item As Object) As String
    Dim rowId As String, stableKey As String, yearValue As Long
    Dim sourceRow As Long, destinationRow As Long
    On Error GoTo InvalidIdentity
    If item Is Nothing Then Exit Function
    If Not item.Exists("rowId") Or Not item.Exists("stableKey") Or Not item.Exists("year") Or _
       Not item.Exists("sourceRow") Or Not item.Exists("destinationRow") Then Exit Function
    rowId = CStr(item("rowId"))
    stableKey = CStr(item("stableKey"))
    yearValue = CLng(item("year"))
    sourceRow = CLng(item("sourceRow"))
    destinationRow = CLng(item("destinationRow"))
    If Len(Trim$(rowId)) = 0 Or Len(stableKey) = 0 Or sourceRow < 2 Or destinationRow < 2 Then Exit Function
    ERecordBucketKey = CStr(yearValue) & ":" & CStr(Len(rowId)) & ":" & rowId & ":" & _
                       CStr(Len(stableKey)) & ":" & stableKey & ":" & CStr(sourceRow) & ":" & CStr(destinationRow)
    Exit Function
InvalidIdentity:
    Err.Clear
End Function

Private Function ERecordSelectionKey(ByVal item As Object) As String
    Dim rowId As String, stableKey As String, yearValue As Long
    On Error GoTo InvalidIdentity
    If item Is Nothing Then Exit Function
    If Not item.Exists("rowId") Or Not item.Exists("stableKey") Or Not item.Exists("year") Then Exit Function
    If item.Exists("identityIssue") Then
        If CBool(item("identityIssue")) Then Exit Function
    End If
    rowId = CStr(item("rowId"))
    stableKey = CStr(item("stableKey"))
    yearValue = CLng(item("year"))
    If yearValue <= 0 Or Len(Trim$(rowId)) = 0 Or Len(Trim$(stableKey)) = 0 Then Exit Function
    ERecordSelectionKey = CStr(yearValue) & ":" & CStr(Len(rowId)) & ":" & rowId & ":" & _
                           CStr(Len(stableKey)) & ":" & stableKey
    Exit Function
InvalidIdentity:
    Err.Clear
End Function

Private Function ERecordRowIdKey(ByVal record As Object) As String
    Dim rowId As String
    rowId = CStr(record("rowId"))
    If Len(Trim$(rowId)) = 0 Then Exit Function
    ERecordRowIdKey = CStr(record("year")) & ":" & CStr(Len(rowId)) & ":" & rowId
End Function

Private Function ERecordStableKeyKey(ByVal record As Object) As String
    Dim stableKey As String
    stableKey = CStr(record("stableKey"))
    If Len(stableKey) = 0 Then Exit Function
    ERecordStableKeyKey = CStr(record("year")) & ":" & CStr(Len(stableKey)) & ":" & stableKey
End Function

Private Function ERecordHasGroup(ByVal record As Object, ByVal selectedGroup As String) As Boolean
    Dim item As Object, recordCases As Collection
    If Len(selectedGroup) = 0 Then
        ERecordHasGroup = True
        Exit Function
    End If
    Set recordCases = record("cases")
    For Each item In recordCases
        If StrComp(CStr(item("groupName")), selectedGroup, vbTextCompare) = 0 Then
            ERecordHasGroup = True
            Exit Function
        End If
    Next item
End Function

Private Function EIsAllDisplayMode() As Boolean
    If mCboEVisibility Is Nothing Then Exit Function
    EIsAllDisplayMode = (mCboEVisibility.ListIndex = 1)
End Function

Private Function EHeaderNormalDisplayCaption(ByVal headerName As String) As String
    Dim headerKey As String
    headerKey = CompareHeaderKey(headerName)
    If headerKey = CompareHeaderKey("毎年度利用区分1:毎年作成 2:単年のみ") Then
        EHeaderNormalDisplayCaption = "毎年度利用区分（1：毎年作成、2：単年のみ）"
    ElseIf headerKey = CompareHeaderKey("原課") Then
        EHeaderNormalDisplayCaption = "原課"
    ElseIf headerKey = CompareHeaderKey("年度（和暦）") Or headerKey = CompareHeaderKey("年度(和暦)") Then
        EHeaderNormalDisplayCaption = "年度（和暦）"
    ElseIf headerKey = CompareHeaderKey("分類名１") Or headerKey = CompareHeaderKey("分類名1") Then
        EHeaderNormalDisplayCaption = "分類名１"
    ElseIf headerKey = CompareHeaderKey("第1ガイド") Or headerKey = CompareHeaderKey("第１ガイド") Or _
           headerKey = CompareHeaderKey("分類名２") Or headerKey = CompareHeaderKey("分類名2") Then
        EHeaderNormalDisplayCaption = "第1ガイド"
    ElseIf headerKey = CompareHeaderKey("第2ガイド") Or headerKey = CompareHeaderKey("第２ガイド") Or _
           headerKey = CompareHeaderKey("分類名３") Or headerKey = CompareHeaderKey("分類名3") Then
        EHeaderNormalDisplayCaption = "第2ガイド"
    ElseIf headerKey = CompareHeaderKey("作成年度（和暦）") Or headerKey = CompareHeaderKey("作成年度(和暦)") Then
        EHeaderNormalDisplayCaption = "作成年度（和暦）"
    ElseIf headerKey = CompareHeaderKey("個別フォルダー") Or headerKey = CompareHeaderKey("個別フォルダ") Or _
           headerKey = CompareHeaderKey("タイトル") Then
        EHeaderNormalDisplayCaption = "個別フォルダー（タイトル）"
    ElseIf headerKey = CompareHeaderKey("保存期間") Then
        EHeaderNormalDisplayCaption = "保存期間"
    ElseIf headerKey = CompareHeaderKey("媒体種別") Then
        EHeaderNormalDisplayCaption = "媒体種別"
    ElseIf headerKey = CompareHeaderKey("内容・取扱い") Then
        EHeaderNormalDisplayCaption = "内容・取扱い"
    ElseIf headerKey = CompareHeaderKey("備考") Then
        EHeaderNormalDisplayCaption = "備考"
    End If
End Function

Private Function EFieldNormalDisplayCaption(ByVal caseItem As Object, ByVal fieldItem As Object) As String
    Dim sourceCaption As String, destinationCaption As String
    sourceCaption = EHeaderNormalDisplayCaption(CStr(fieldItem("sourceHeader")))
    destinationCaption = EHeaderNormalDisplayCaption(CStr(fieldItem("destinationHeader")))
    If Len(sourceCaption) = 0 Or StrComp(sourceCaption, destinationCaption, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(CStr(fieldItem("systemValue")), CStr(fieldItem("excelValue")), vbBinaryCompare) = 0 Then Exit Function
    If EGroupMatchesNormalDisplay(CStr(caseItem("groupName")), CStr(fieldItem("sourceHeader")), sourceCaption) Then _
        EFieldNormalDisplayCaption = sourceCaption
End Function

Private Function EGroupMatchesNormalDisplay(ByVal groupName As String, ByVal sourceHeader As String, _
                                            ByVal categoryCaption As String) As Boolean
    Dim groupKey As String, sourceKey As String
    groupKey = CompareHeaderKey(groupName)
    sourceKey = CompareHeaderKey(sourceHeader)
    Select Case categoryCaption
        Case "原課"
            EGroupMatchesNormalDisplay = (groupKey = CompareHeaderKey("原課・原課ID") And _
                                          sourceKey = CompareHeaderKey("原課"))
        Case "分類名１"
            EGroupMatchesNormalDisplay = (groupKey = CompareHeaderKey("分類名１・分類コード１") And _
                (sourceKey = CompareHeaderKey("分類名１") Or sourceKey = CompareHeaderKey("分類名1")))
        Case "第1ガイド"
            EGroupMatchesNormalDisplay = (groupKey = CompareHeaderKey("第1ガイド・分類コード２") And _
                (sourceKey = CompareHeaderKey("第1ガイド") Or sourceKey = CompareHeaderKey("第１ガイド") Or _
                 sourceKey = CompareHeaderKey("分類名２") Or sourceKey = CompareHeaderKey("分類名2")))
        Case "第2ガイド"
            EGroupMatchesNormalDisplay = (groupKey = CompareHeaderKey("第2ガイド・分類コード３") And _
                (sourceKey = CompareHeaderKey("第2ガイド") Or sourceKey = CompareHeaderKey("第２ガイド") Or _
                 sourceKey = CompareHeaderKey("分類名３") Or sourceKey = CompareHeaderKey("分類名3")))
        Case "保存期間"
            EGroupMatchesNormalDisplay = (groupKey = CompareHeaderKey("保存期間・保存期間ID") And _
                                          sourceKey = CompareHeaderKey("保存期間"))
        Case "媒体種別"
            EGroupMatchesNormalDisplay = (groupKey = CompareHeaderKey("媒体種別・媒体種別ID") And _
                                          sourceKey = CompareHeaderKey("媒体種別"))
        Case Else
            EGroupMatchesNormalDisplay = (StrComp(groupKey, sourceKey, vbBinaryCompare) = 0)
    End Select
End Function

Private Function ECaseHasNormalDisplayFields(ByVal caseItem As Object) As Boolean
    Dim fields As Collection, fieldItem As Object
    Set fields = caseItem("headers")
    For Each fieldItem In fields
        If Len(EFieldNormalDisplayCaption(caseItem, fieldItem)) > 0 Then
            ECaseHasNormalDisplayFields = True
            Exit Function
        End If
    Next fieldItem
End Function

Private Function ECaseHasOutOfScopeFields(ByVal caseItem As Object) As Boolean
    Dim fields As Collection, fieldItem As Object
    Set fields = caseItem("headers")
    For Each fieldItem In fields
        If Len(EFieldNormalDisplayCaption(caseItem, fieldItem)) = 0 Then
            ECaseHasOutOfScopeFields = True
            Exit Function
        End If
    Next fieldItem
End Function

Private Function ESelectedDecisionIsSafe(ByVal decision As String, ByRef reasonText As String) As Boolean
    Dim item As Object, record As Object, previewText As String
    reasonText = ""
    If mCboCategory.ListIndex <> 4 Or ESelectedCount() <> 1 Then
        reasonText = "ファイル1件と判断項目1件を選択してください。"
        Exit Function
    End If
    Set item = SelectedCase()
    Set record = SelectedRecord()
    If item Is Nothing Or record Is Nothing Then
        reasonText = "判断対象を一意に特定できません。"
        Exit Function
    End If
    If CBool(record("identityIssue")) Then
        reasonText = "識別情報の一貫性を確認できないため、採用・保留はできません。"
        Exit Function
    End If
    If EIsClassification1Group(item) Then
        ESelectedDecisionIsSafe = Module1.BidirectionalSyncClassification1DecisionSafe(item, decision, reasonText, previewText)
        Exit Function
    End If
    If EIsRetentionGroup(item) Then
        ESelectedDecisionIsSafe = Module1.BidirectionalSyncRetentionDecisionSafe(item, decision, reasonText, previewText)
        Exit Function
    End If
    If Not EIsAllDisplayMode() And ECaseHasOutOfScopeFields(item) Then
        reasonText = "この判断グループには入力項目のみの表示対象外の関連項目があります。全項目（管理コード含む）に切り替えてください。"
        Exit Function
    End If
    ESelectedDecisionIsSafe = True
End Function

Private Function EIsClassification1Group(ByVal item As Object) As Boolean
    If item Is Nothing Then Exit Function
    EIsClassification1Group = (StrComp(CStr(item("groupName")), "分類名１・分類コード１", vbBinaryCompare) = 0)
End Function

Private Function EIsRetentionGroup(ByVal item As Object) As Boolean
    If item Is Nothing Then Exit Function
    EIsRetentionGroup = (StrComp(CStr(item("groupName")), "保存期間・保存期間ID", vbBinaryCompare) = 0)
End Function

Private Function EConfirmClassification1Decision( _
    ByVal item As Object, ByVal record As Object, ByVal decision As String, ByRef reasonText As String) As Boolean

    Dim previewText As String, sourceLabel As String, stateLabel As String
    Dim guide2Value As String, folderValue As String, creationYearValue As String, confirmationText As String, failureText As String

    On Error GoTo Failed
    reasonText = ""
    If item Is Nothing Or record Is Nothing Then
        reasonText = "対象ファイルまたは判断単位を確認できません。"
        EClearClassification1PreviewSignature item
        Exit Function
    End If
    If CBool(record("identityIssue")) Then
        reasonText = "識別情報の一貫性を確認できないため、採用・保留はできません。"
        EClearClassification1PreviewSignature item
        Exit Function
    End If
    EClearClassification1PreviewSignature item
    If Not Module1.BidirectionalSyncClassification1DecisionSafe(item, decision, reasonText, previewText, True, False) Then Exit Function
    ERecordDisplayValues record, guide2Value, folderValue, creationYearValue
    If Len(folderValue) = 0 Or folderValue = "―" Or folderValue = "（空欄）" Then
        reasonText = "確認画面に対象ファイル名を表示できないため、連動判断を停止しました。"
        EClearClassification1PreviewSignature item
        Exit Function
    End If
    Select Case decision
        Case "システム採用"
            sourceLabel = "文書管理システム"
            stateLabel = "文書管理を採用（関連項目グループ全体）"
        Case "Excel採用"
            sourceLabel = "Excel基準表"
            stateLabel = "Excel基準表を採用（関連項目グループ全体）"
        Case "保留"
            sourceLabel = "なし（値は変更しません）"
            stateLabel = "保留（関連項目グループ全体）"
        Case Else
            reasonText = "分類名１・分類コード１の判断方法が不正です。"
            EClearClassification1PreviewSignature item
            Exit Function
    End Select
    confirmationText = "対象ファイル（個別フォルダー）: " & folderValue & vbCrLf & _
        ERecordIdentityDetails(record) & vbCrLf & "採用元: " & sourceLabel & vbCrLf & _
        previewText & vbCrLf & "採用後の判断状態: " & stateLabel & vbCrLf & _
        "この内容で、このファイル・年度の判断単位だけを確定します。続けますか？"
    If MsgBox(confirmationText, vbQuestion + vbYesNo, "分類名１・分類コード１の連動判断確認") <> vbYes Then
        EClearClassification1PreviewSignature item
        Exit Function
    End If
    EConfirmClassification1Decision = True
    Exit Function

Failed:
    failureText = Err.Description
    EClearClassification1PreviewSignature item
    reasonText = "採用前の確認画面を完了できなかったため、連動判断を停止しました。" & failureText
    Err.Clear
End Function

Private Function EConfirmRetentionDecision( _
    ByVal item As Object, ByVal record As Object, ByVal decision As String, ByRef reasonText As String) As Boolean

    Dim previewText As String, sourceLabel As String, stateLabel As String
    Dim guide2Value As String, folderValue As String, creationYearValue As String
    Dim confirmationText As String, failureText As String

    On Error GoTo Failed
    reasonText = ""
    If item Is Nothing Or record Is Nothing Then
        reasonText = "対象ファイルまたは保存期間の判断単位を確認できません。"
        EClearRetentionPreviewSignature item
        Exit Function
    End If
    If Not EIsRetentionGroup(item) Or CBool(record("identityIssue")) Then
        reasonText = "識別情報または保存期間・保存期間IDの判断単位を確認できないため、処理を停止しました。"
        EClearRetentionPreviewSignature item
        Exit Function
    End If
    EClearRetentionPreviewSignature item
    If Not Module1.BidirectionalSyncRetentionDecisionSafe(item, decision, reasonText, previewText, True, False) Then
        EClearRetentionPreviewSignature item
        Exit Function
    End If
    ERecordDisplayValues record, guide2Value, folderValue, creationYearValue
    If Len(folderValue) = 0 Or folderValue = "―" Or folderValue = "（空欄）" Then
        reasonText = "確認画面に対象ファイル名を表示できないため、連動判断を停止しました。"
        EClearRetentionPreviewSignature item
        Exit Function
    End If
    Select Case decision
        Case "システム採用"
            sourceLabel = "文書管理システム"
            stateLabel = "文書管理を採用（保存期間・保存期間IDのグループ全体）"
        Case "Excel採用"
            sourceLabel = "Excel基準表"
            stateLabel = "Excel基準表を採用（保存期間・保存期間IDのグループ全体）"
        Case "保留"
            sourceLabel = "なし（値は変更しません）"
            stateLabel = "保留（保存期間・保存期間IDのグループ全体）"
        Case Else
            reasonText = "保存期間・保存期間IDの判断方法が不正です。"
            EClearRetentionPreviewSignature item
            Exit Function
    End Select
    confirmationText = "対象ファイル（個別フォルダー）: " & folderValue & vbCrLf & _
        ERecordIdentityDetails(record) & vbCrLf & "採用元: " & sourceLabel & vbCrLf & _
        previewText & vbCrLf & "採用後の判断状態: " & stateLabel & vbCrLf & _
        "この内容で、このファイル・年度の保存期間関連項目だけを確定します。続けますか？"
    If MsgBox(confirmationText, vbQuestion + vbYesNo, "保存期間・保存期間IDの連動判断確認") <> vbYes Then
        EClearRetentionPreviewSignature item
        Exit Function
    End If
    EConfirmRetentionDecision = True
    Exit Function

Failed:
    failureText = Err.Description
    EClearRetentionPreviewSignature item
    reasonText = "採用前の確認画面を完了できなかったため、連動判断を停止しました。" & failureText
    Err.Clear
End Function

Private Sub EClearRetentionPreviewSignature(ByVal item As Object)
    On Error Resume Next
    If Not item Is Nothing Then
        If item.Exists("linkedRetentionSignature") Then item.Remove "linkedRetentionSignature"
    End If
    Err.Clear
End Sub

Private Sub EClearClassification1PreviewSignature(ByVal item As Object)
    On Error Resume Next
    If Not item Is Nothing Then
        If item.Exists("linkedPairSignature") Then item.Remove "linkedPairSignature"
    End If
    Err.Clear
End Sub

Private Function ECaseNeedsHumanDecision(ByVal caseItem As Object) As Boolean
    Dim decision As String
    decision = SyncReviewDecision(caseItem)
    ECaseNeedsHumanDecision = (Len(decision) = 0 Or decision = "保留")
End Function

Private Function ERecordNeedsConfirmation(ByVal record As Object) As Boolean
    Dim item As Object, recordCases As Collection
    If CBool(record("identityIssue")) Then Exit Function
    Set recordCases = record("cases")
    For Each item In recordCases
        If ECaseNeedsHumanDecision(item) And ECaseHasNormalDisplayFields(item) Then
            ERecordNeedsConfirmation = True
            Exit Function
        End If
    Next item
End Function

Private Function ERecordNeedsInternalInvestigation(ByVal record As Object) As Boolean
    Dim item As Object, recordCases As Collection
    If CBool(record("identityIssue")) Then
        ERecordNeedsInternalInvestigation = True
        Exit Function
    End If
    Set recordCases = record("cases")
    For Each item In recordCases
        If ECaseNeedsHumanDecision(item) And ECaseHasOutOfScopeFields(item) Then
            ERecordNeedsInternalInvestigation = True
            Exit Function
        End If
    Next item
End Function

Private Sub ERecordDisplayValues(ByVal record As Object, ByRef guide2Value As String, _
                                 ByRef folderValue As String, ByRef creationYearValue As String)
    Dim src As Worksheet, dst As Worksheet, sourceRow As Long, destinationRow As Long
    On Error GoTo Unavailable
    Set src = ThisWorkbook.Worksheets(CSV_SHEET)
    Set dst = ThisWorkbook.Worksheets(DEST_SHEET)
    sourceRow = CLng(record("sourceRow"))
    destinationRow = CLng(record("destinationRow"))
    guide2Value = EIdentityPair(src, dst, sourceRow, destinationRow, 7)
    folderValue = EIdentityPair(src, dst, sourceRow, destinationRow, 8)
    creationYearValue = EIdentityPair(src, dst, sourceRow, destinationRow, 5)
    If CBool(record("identityIssue")) Then guide2Value = "識別要確認: " & guide2Value
    Exit Sub
Unavailable:
    Err.Clear
    guide2Value = "識別情報を取得できません"
    folderValue = "―"
    creationYearValue = "要確認"
End Sub

Private Function EIdentityPair(ByVal src As Worksheet, ByVal dst As Worksheet, _
                               ByVal sourceRow As Long, ByVal destinationRow As Long, _
                               ByVal fieldIndex As Long) As String
    Dim fieldCaption As String, csvAliases As Variant, excelAliases As Variant
    Dim sourceValue As String, excelValue As String
    Dim sourceHeaderFound As Boolean, excelHeaderFound As Boolean
    CompareFieldSpec fieldIndex, fieldCaption, csvAliases, excelAliases
    sourceValue = EIdentityFieldValue(src, sourceRow, csvAliases, sourceHeaderFound)
    excelValue = EIdentityFieldValue(dst, destinationRow, excelAliases, excelHeaderFound)
    If Not sourceHeaderFound Then sourceValue = "―"
    If Not excelHeaderFound Then excelValue = "―"
    If StrComp(sourceValue, excelValue, vbBinaryCompare) = 0 Then
        EIdentityPair = sourceValue
    Else
        EIdentityPair = "文書管理: " & sourceValue & " / Excel: " & excelValue
    End If
End Function

Private Function EIdentityFieldValue(ByVal ws As Worksheet, ByVal rowNo As Long, _
                                      ByVal aliases As Variant, ByRef headerFound As Boolean) As String
    EIdentityFieldValue = CompareCellText(ws, rowNo, aliases, headerFound)
    If Len(EIdentityFieldValue) = 0 And headerFound Then EIdentityFieldValue = "（空欄）"
End Function

Private Function SelectedRecord() As Object
    Dim i As Long
    If mCboCategory.ListIndex <> 4 Or shown Is Nothing Then Exit Function
    If ESelectedCount() = 0 Then Exit Function
    If mLstCases.ListIndex >= 0 And mLstCases.ListIndex + 1 <= shown.Count Then
        If mLstCases.Selected(mLstCases.ListIndex) Then
            Set SelectedRecord = shown(mLstCases.ListIndex + 1)
            Exit Function
        End If
    End If
    For i = 0 To mLstCases.ListCount - 1
        If mLstCases.Selected(i) Then
            Set SelectedRecord = shown(i + 1)
            Exit Function
        End If
    Next i
End Function

Private Function SelectedEField() As Object
    If mCboCategory.ListIndex <> 4 Or mLstRows.ListIndex < 0 Then Exit Function
    If eFieldItems Is Nothing Then Exit Function
    If mLstRows.ListIndex + 1 <= eFieldItems.Count Then _
        Set SelectedEField = eFieldItems(mLstRows.ListIndex + 1)
End Function

Private Sub ShowECurrent()
    Dim record As Object, item As Object, fieldItem As Object
    Dim recordCases As Collection, fieldItems As Collection
    Dim rowIndex As Long, decision As String, displayCaption As String, allMode As Boolean
    Dim normalTargetItems As Long, outOfTargetItems As Long, mixedGroups As Long
    Dim linkedGroups As Long, retentionLinkedGroups As Long
    Dim selectionKey As String
    If rendering Then Exit Sub
    On Error GoTo DisplayFailed
    selectionKey = ESelectionRenderKey()
    If selectionKey = eRenderedSelection Then Exit Sub
    BeginEDisplayRead
    rendering = True
    Set eFieldCases = New Collection
    Set eFieldItems = New Collection
    mLstRows.Clear
    mLstRows.ColumnCount = 5
    mLstRows.MultiSelect = fmMultiSelectSingle
    mLstRows.ListIndex = -1
    Set record = SelectedRecord()
    If record Is Nothing Then
        mLblCompareRows.Caption = "ファイルを選択してください。"
        mTxtDetails.Text = "確認対象ファイルを左側から選択すると、そのファイルの不一致項目を表示します。"
        mLblStatus.Caption = "複数ファイル選択は表示専用です。判断時はファイル1件と不一致項目1件を選択してください。"
        mCmdApprove.Enabled = False
        mCmdOther.Enabled = False
        mCmdHold.Enabled = False
        mCmdGoTo.Enabled = False
        mCmdClear.Enabled = False
        mChkDelete.Visible = False
        rendering = False
        UpdateESelection
        EndEDisplayRead
        eRenderedSelection = selectionKey
        Exit Sub
    End If
    Set recordCases = record("cases")
    allMode = EIsAllDisplayMode()
    For Each item In recordCases
        Set fieldItems = item("headers")
        decision = EConflictDecisionDisplay(item)
        If ECaseHasNormalDisplayFields(item) Then
            normalTargetItems = normalTargetItems + 1
            If ECaseHasOutOfScopeFields(item) Then
                If EIsClassification1Group(item) Then
                    linkedGroups = linkedGroups + 1
                ElseIf EIsRetentionGroup(item) Then
                    retentionLinkedGroups = retentionLinkedGroups + 1
                Else
                    mixedGroups = mixedGroups + 1
                End If
            End If
        Else
            outOfTargetItems = outOfTargetItems + 1
        End If
        For Each fieldItem In fieldItems
            displayCaption = EFieldNormalDisplayCaption(item, fieldItem)
            If allMode Or Len(displayCaption) > 0 Then
                rowIndex = mLstRows.ListCount
                If allMode Then
                    mLstRows.AddItem CStr(fieldItem("sourceHeader"))
                Else
                    mLstRows.AddItem displayCaption
                End If
                If CBool(fieldItem("baselineKnown")) Then
                    mLstRows.List(rowIndex, 1) = EConflictValueDisplay(CStr(fieldItem("baselineValue")))
                Else
                    mLstRows.List(rowIndex, 1) = "未確定"
                End If
                mLstRows.List(rowIndex, 2) = EConflictValueDisplay(CStr(fieldItem("systemValue")))
                mLstRows.List(rowIndex, 3) = EConflictValueDisplay(CStr(fieldItem("excelValue")))
                mLstRows.List(rowIndex, 4) = decision
                eFieldCases.Add item
                eFieldItems.Add fieldItem
            End If
        Next fieldItem
    Next item
    mLstRows.ListIndex = -1
    mLblCompareRows.Caption = "CSV行 " & CStr(record("sourceRow")) & " ／ Excel行 " & _
                              CStr(record("destinationRow")) & " ／ 判断単位 " & CStr(recordCases.Count) & "件"
    mTxtDetails.Text = ERecordIdentityDetails(record)
    If allMode Then
        If CBool(record("identityIssue")) Then
            mLblStatus.Caption = "識別情報が要確認のファイルです。採用・保留はできません。別のレコードを巻き込まないよう、識別情報を確認してください。"
        ElseIf Not ERecordNeedsConfirmation(record) And ERecordNeedsInternalInvestigation(record) Then
            mLblStatus.Caption = "内部項目だけに未解決差異があるファイルです。自動解決はせず、差異を保持しています。"
        Else
            mLblStatus.Caption = "不一致項目を選ぶと、理由と判断範囲を確認できます。"
        End If
    Else
        mLblStatus.Caption = "入力項目の不一致 " & CStr(normalTargetItems) & "件 ／ 表示対象外 " & _
            CStr(outOfTargetItems) & "件（既存の判断単位）"
        If mixedGroups > 0 Then mLblStatus.Caption = mLblStatus.Caption & _
            "。その他の関連項目を含む" & CStr(mixedGroups) & "件は、全項目を確認するまで判断できません。"
        If linkedGroups > 0 Then mLblStatus.Caption = mLblStatus.Caption & _
            "。分類名１・分類コード１の連動グループ" & CStr(linkedGroups) & _
            "件は、対応関係を検証し、採用前の確認画面で両方の値を確認します。"
        If retentionLinkedGroups > 0 Then mLblStatus.Caption = mLblStatus.Caption & _
            "。保存期間・保存期間IDの連動グループ" & CStr(retentionLinkedGroups) & _
            "件は、対応関係を検証し、採用前の確認画面で両方の値を確認します。"
        If mLstRows.ListCount = 0 Then
            mTxtDetails.Text = mTxtDetails.Text & vbCrLf & _
                "［入力項目のみ］の対象項目に不一致はありません。その他の項目に" & CStr(outOfTargetItems) & _
                "件の不一致があります。「全項目（管理コード含む）」に切り替えると確認できます。"
            mLblStatus.Caption = "このファイルの競合は［入力項目のみ］の対象外です。ファイル一覧には残しています。"
        End If
    End If
    mCmdApprove.Caption = "文書管理を採用"
    mCmdOther.Caption = "Excel基準表を採用"
    mCmdHold.Caption = "保留"
    mCmdGoTo.Enabled = False
    mCmdClear.Enabled = False
    mChkDelete.Visible = False
    rendering = False
    UpdateESelection
    EndEDisplayRead
    eRenderedSelection = selectionKey
    Exit Sub
DisplayFailed:
    EDisplayFailed Err.Description
End Sub

Private Sub UpdateEFieldSelection()
    Dim record As Object, item As Object, fieldItem As Object
    If rendering Or mCboCategory.ListIndex <> 4 Then Exit Sub
    Set record = SelectedRecord()
    If record Is Nothing Then Exit Sub
    On Error GoTo DisplayFailed
    BeginEDisplayRead
    Set item = SelectedCase()
    Set fieldItem = SelectedEField()
    If item Is Nothing Or fieldItem Is Nothing Then
        mTxtDetails.Text = ERecordIdentityDetails(record)
        If ESelectedCount() > 1 Then
            mLblStatus.Caption = "複数ファイルの選択は表示専用です。判断にはファイル1件と項目1件を選んでください。"
        Else
            mLblStatus.Caption = "判断する不一致項目を1件選択してください。"
        End If
        UpdateESelection
        EndEDisplayRead
        Exit Sub
    End If
    mTxtDetails.Text = ERecordIdentityDetails(record) & vbCrLf & EConflictFieldDetails(item, fieldItem)
    If EIsClassification1Group(item) Then
            mLblStatus.Caption = "分類名１と分類コード１は同じ判断単位です。採用前の確認画面で両方の値と採用元を確認します。"
        ElseIf EIsRetentionGroup(item) Then
            mLblStatus.Caption = "保存期間と保存期間IDは同じ判断単位です。採用前に対応関係を検証し、確認画面で両方の値を確認します。"
        ElseIf Not EIsAllDisplayMode() And ECaseHasOutOfScopeFields(item) Then
            mLblStatus.Caption = "この判断単位には［入力項目のみ］の表示対象外の関連項目が含まれます。「全項目（管理コード含む）」に切り替えてください。"
        ElseIf EIsAllDisplayMode() Then
            mLblStatus.Caption = "判断単位「" & CStr(item("groupName")) & "」の関連項目全体に選択した判断が適用されます。"
        Else
            mLblStatus.Caption = "［入力項目のみ］の項目を1件選択しています。判断は表示中の判断単位に適用されます。"
        End If
    UpdateESelection
    EndEDisplayRead
    Exit Sub
DisplayFailed:
    EDisplayFailed Err.Description
End Sub

Private Function ERecordIdentityDetails(ByVal record As Object) As String
    Dim src As Worksheet, dst As Worksheet, sourceRow As Long, destinationRow As Long
    Dim recordCases As Collection, firstCase As Object, matchKind As String
    On Error GoTo Unavailable
    Set src = ThisWorkbook.Worksheets(CSV_SHEET)
    Set dst = ThisWorkbook.Worksheets(DEST_SHEET)
    sourceRow = CLng(record("sourceRow"))
    destinationRow = CLng(record("destinationRow"))
    Set recordCases = record("cases")
    If recordCases.Count > 0 Then
        Set firstCase = recordCases(1)
        matchKind = CStr(firstCase("matchKind"))
    End If
    ERecordIdentityDetails = "第2ガイド: " & EIdentityPair(src, dst, sourceRow, destinationRow, 7) & _
        " ／ 個別フォルダー: " & EIdentityPair(src, dst, sourceRow, destinationRow, 8) & _
        " ／ 作成年度: " & EIdentityPair(src, dst, sourceRow, destinationRow, 5) & vbCrLf & _
        "第1ガイド: " & EIdentityPair(src, dst, sourceRow, destinationRow, 6) & _
        " ／ 正式ファイルID: " & EIdentityPair(src, dst, sourceRow, destinationRow, 1) & _
        " ／ ファイル番号: " & EIdentityPair(src, dst, sourceRow, destinationRow, 2) & _
        " ／ バーコード番号: " & EIdentityPair(src, dst, sourceRow, destinationRow, 3) & vbCrLf & _
        "Excel管理行ID: " & CStr(record("rowId")) & " ／ CSV安定識別キー: " & CStr(record("stableKey")) & _
        " ／ 照合方法: " & matchKind & vbCrLf & _
        "同期対象年度: " & CStr(record("year")) & " ／ CSV行: " & CStr(sourceRow) & _
        " ／ Excel行: " & CStr(destinationRow)
    If CBool(record("identityIssue")) Then
        ERecordIdentityDetails = ERecordIdentityDetails & vbCrLf & _
            "注意: 管理行ID・安定識別キー・行番号の一貫性を確認できません。別ファイルと統合せず、判断前に確認してください。"
    End If
    Exit Function
Unavailable:
    Err.Clear
    ERecordIdentityDetails = "ファイルの識別情報を取得できません。安全のため判断前に再読み込みしてください。"
End Function

Private Function EConflictFieldDetails(ByVal item As Object, ByVal selectedField As Object) As String
    Dim fieldItems As Collection, fieldItem As Object, fieldNames As String
    Dim previousValue As String, fieldCaption As String, selectedCaption As String
    Dim groupCaption As String, hiddenFieldCount As Long, allMode As Boolean
    Set fieldItems = item("headers")
    allMode = EIsAllDisplayMode()
    For Each fieldItem In fieldItems
        fieldCaption = EFieldNormalDisplayCaption(item, fieldItem)
        If allMode Then fieldCaption = CStr(fieldItem("sourceHeader"))
        If Len(fieldCaption) > 0 Then
            If Len(fieldNames) > 0 Then fieldNames = fieldNames & "、"
            fieldNames = fieldNames & fieldCaption
        Else
            hiddenFieldCount = hiddenFieldCount + 1
        End If
    Next fieldItem
    If allMode Then
        groupCaption = CStr(item("groupName"))
    ElseIf hiddenFieldCount > 0 Then
        groupCaption = "入力項目を含む判断グループ"
    Else
        groupCaption = fieldNames
    End If
    If Len(fieldNames) = 0 Then fieldNames = "（表示できる項目なし）"
    EConflictFieldDetails = "判断単位: " & groupCaption & " ／ 関連項目: " & fieldNames
    If Not allMode And hiddenFieldCount > 0 Then
        EConflictFieldDetails = EConflictFieldDetails & "（［入力項目のみ］の表示対象外の関連項目 " & _
            CStr(hiddenFieldCount) & "件は非表示）"
    End If
    selectedCaption = CStr(selectedField("sourceHeader"))
    If Not allMode Then selectedCaption = EFieldNormalDisplayCaption(item, selectedField)
    EConflictFieldDetails = EConflictFieldDetails & vbCrLf & _
        "理由: " & EConflictReasonDisplay(item, allMode) & vbCrLf & _
        "選択項目: " & selectedCaption & " ／ 文書管理システム: " & _
        EConflictValueDisplay(CStr(selectedField("systemValue"))) & " ／ Excel基準表: " & _
        EConflictValueDisplay(CStr(selectedField("excelValue")))
    If CBool(selectedField("baselineKnown")) Then
        previousValue = EConflictValueDisplay(CStr(selectedField("baselineValue")))
    Else
        previousValue = "未確定"
    End If
    EConflictFieldDetails = EConflictFieldDetails & vbCrLf & "前回同期値: " & previousValue & _
        " ／ この判断は同じ判断単位の関連項目全体に適用されます。"
    If Not allMode And hiddenFieldCount > 0 Then
        If EIsClassification1Group(item) Then
            EConflictFieldDetails = EConflictFieldDetails & vbCrLf & _
                "非表示の分類コード１も同じ判断単位に含まれます。対応関係を検証できた場合のみ、採用前の確認画面に両方の値を表示します。"
        ElseIf EIsRetentionGroup(item) Then
            EConflictFieldDetails = EConflictFieldDetails & vbCrLf & _
                "非表示の保存期間IDも同じ判断単位に含まれます。現在の目録CSVとコード管理CSVの両方で対応が一意な場合だけ採用を許可し、確認画面にIDを表示します。"
        Else
            EConflictFieldDetails = EConflictFieldDetails & vbCrLf & _
                "安全のため、この項目グループの採用・保留は無効です。「全項目（管理コード含む）」で関連項目を確認してください。"
        End If
    End If
    If InStr(1, CStr(selectedField("sourceHeader")), "コード", vbTextCompare) > 0 Or _
       StrComp(CStr(selectedField("sourceHeader")), "汎用項目１０", vbTextCompare) = 0 Then
        EConflictFieldDetails = EConflictFieldDetails & vbCrLf & _
            "コード欄: 上の値は保存済みの競合比較値です。ここでは値を変更しません。"
    End If
End Function

Private Function EConflictValueDisplay(ByVal valueText As String) As String
    If Len(valueText) = 0 Then
        EConflictValueDisplay = "（空欄）"
    Else
        EConflictValueDisplay = valueText
    End If
End Function

Private Function EConflictDecisionDisplay(ByVal item As Object) As String
    Select Case SyncReviewDecision(item)
        Case ""
            EConflictDecisionDisplay = "未判断"
        Case "システム採用"
            EConflictDecisionDisplay = "文書管理を採用"
        Case "Excel採用"
            EConflictDecisionDisplay = "Excel基準表を採用"
        Case "保留"
            EConflictDecisionDisplay = "保留中"
        Case Else
            EConflictDecisionDisplay = "要確認: " & SyncReviewDecision(item)
    End Select
End Function

Private Function EConflictReasonDisplay(ByVal item As Object, Optional ByVal includeTechnicalDetail As Boolean = True) As String
    Dim reasonText As String, summary As String
    reasonText = CStr(item("reason"))
    If InStr(1, reasonText, "前回同期確認値がない", vbTextCompare) > 0 Or _
       InStr(1, reasonText, "初回不一致", vbTextCompare) > 0 Then
        summary = "前回の同期確認値がありません。文書管理システムとExcelの値を確認し、採用する側を判断してください。"
    ElseIf InStr(1, reasonText, "両方が前回確認値", vbTextCompare) > 0 Then
        summary = "前回の確認後に両方の値が変更されています。自動では選ばず、両方の値を保持しています。"
    ElseIf InStr(1, reasonText, "汎用項目１０", vbTextCompare) > 0 Then
        summary = "汎用項目10のコード差を安全に判断できないため、自動更新を止めています。"
    ElseIf InStr(1, reasonText, "Excel管理行IDで対象を特定", vbTextCompare) > 0 Then
        summary = "保存済みの競合を再表示しています。前回同期値と両側の現在値を確認してから判断してください。"
    Else
        summary = "値の違いを自動で判断できないため、両方の値を保持しています。"
    End If
    If includeTechnicalDetail Then
        EConflictReasonDisplay = summary & vbCrLf & "検出時の説明: " & reasonText
    ElseIf InStr(1, reasonText, "汎用項目１０", vbTextCompare) > 0 Then
        EConflictReasonDisplay = "追加の関連情報があるため、全項目表示に切り替えて確認してください。"
    Else
        EConflictReasonDisplay = summary
    End If
End Function

Private Sub SelectNextECase()
    Dim recordIndex As Long, record As Object, recordCases As Collection, item As Object
    Dim foundOutOfTarget As Boolean
    If shown Is Nothing Then Exit Sub
    For recordIndex = 1 To shown.Count
        Set record = shown(recordIndex)
        Set recordCases = record("cases")
        For Each item In recordCases
            If Len(SyncReviewDecision(item)) = 0 Then
                If Not EIsAllDisplayMode() And _
                   (Not ECaseHasNormalDisplayFields(item) Or _
                    (ECaseHasOutOfScopeFields(item) And Not EIsClassification1Group(item) And _
                     Not EIsRetentionGroup(item))) Then
                    foundOutOfTarget = True
                Else
                    SelectSingleECase recordIndex - 1
                    SelectEFieldCase item
                    Exit Sub
                End If
            End If
        Next item
    Next recordIndex
    If Not EIsAllDisplayMode() And foundOutOfTarget Then
        mLblStatus.Caption = "［入力項目のみ］で操作できる未解決項目はありません。表示対象外または関連項目を含む不一致が残る場合は「全項目（管理コード含む）」に切り替えてください。"
    Else
        mLblStatus.Caption = "この絞り込み条件に未解決の不一致項目はありません。"
    End If
End Sub

Private Sub SelectEFieldCase(ByVal targetCase As Object)
    Dim i As Long, candidate As Object
    If targetCase Is Nothing Or eFieldCases Is Nothing Then Exit Sub
    For i = 1 To eFieldCases.Count
        Set candidate = eFieldCases(i)
        If StrComp(CStr(candidate("key")), CStr(targetCase("key")), vbBinaryCompare) = 0 Then
            rendering = True
            mLstRows.ListIndex = i - 1
            rendering = False
            UpdateEFieldSelection
            Exit Sub
        End If
    Next i
End Sub

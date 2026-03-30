Option Explicit

'========================================================
' frmSeal に貼り付ける btnApply_Click（新仕様）
'--------------------------------------------------------
' 前提:
' - 標準モジュール側に ApplyBySerial_NewSpec が存在
' - 開始通し番号は txtSerial（なければ txtStart）から取得
'========================================================
Private Sub btnApply_Click()

    Dim startSerial As String

    On Error GoTo ErrorHandler

    ' 既存フォーム名に合わせて必要に応じて調整
    startSerial = ""
    On Error Resume Next
    startSerial = Trim$(CStr(Me.Controls("txtSerial").Value))
    If Len(startSerial) = 0 Then
        startSerial = Trim$(CStr(Me.Controls("txtStart").Value))
    End If
    On Error GoTo ErrorHandler

    If Len(startSerial) = 0 Then
        MsgBox "通し番号（開始番号）を入力してください。", vbExclamation
        Exit Sub
    End If

    ApplyBySerial_NewSpec startSerial
    MsgBox "個別フォルダシールへ反映しました。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "反映処理でエラーが発生しました: " & Err.Description, vbExclamation

End Sub

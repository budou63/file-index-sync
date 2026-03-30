Option Explicit

'========================================================
' 貼り付け先:
'   VBAProject
'   └ Microsoft Excel Objects
'      └ 新ファイル基準表
'
' ※標準モジュールではなく、必ず「新ファイル基準表」の
'   シートコードモジュールへ配置してください。
'========================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ExitHandler
    Application.EnableEvents = False

    ' 標準モジュール(Module1)の補完処理を呼び出す
    新ファイル基準表_表示名からID自動補完 Target

ExitHandler:
    Application.EnableEvents = True

End Sub

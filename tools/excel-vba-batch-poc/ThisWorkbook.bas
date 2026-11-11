Option Explicit
Private Sub Workbook_Open()
    On Error GoTo Failed
    Dim f As Integer
    Dim val As String
    Dim number As Long
    Dim description As String
    val = CStr(ThisWorkbook.Worksheets(1).Range("A2").Value)
    ThisWorkbook.Worksheets(1).Range("A1").Value = val & ChrW(12539) & ChrW(20966) & ChrW(29702) & ChrW(23436) & ChrW(20102)
    ThisWorkbook.Save
    f = FreeFile
    Open ThisWorkbook.Path & "\vba-result.txt" For Output As #f
    Print #f, CStr(ThisWorkbook.Worksheets(1).Range("A3").Value) & "|ok|" & val & "|" & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Close #f
    Exit Sub
Failed:
    number = Err.Number
    description = Err.Description
    On Error Resume Next
    f = FreeFile
    Open ThisWorkbook.Path & "\vba-result.txt" For Output As #f
    Print #f, CStr(ThisWorkbook.Worksheets(1).Range("A3").Value) & "|error|" & CStr(number) & "|" & Replace(description, "|", "/")
    Close #f
End Sub

Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
    CompactOpen
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
    Dim oldEvents As Boolean
    oldEvents = Application.EnableEvents
    If Not oldEvents Then Exit Sub
    On Error GoTo Done
    Application.EnableEvents = False
    CompactHandleChange Sh, target
Done:
    Application.EnableEvents = oldEvents
End Sub

Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal target As Hyperlink)
    If Sh.name = "İzin Girişi" Then
        If target.Range.column = 8 And target.Range.row >= 10 Then CompactSelectPerson target.Range.row
    End If
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal target As Range, Cancel As Boolean)
    If Sh.name <> "İzin Girişi" Then Exit Sub
    If Not Intersect(target, Sh.Range("J9:L10")) Is Nothing Then
        Cancel = True: CompactDateStart
    ElseIf Not Intersect(target, Sh.Range("M9:O10")) Is Nothing Then
        Cancel = True: CompactDateEnd
    ElseIf target.row >= 26 And target.row <= 35 And target.column >= 10 And target.column <= 15 Then
        If Len(CStr(Sh.Cells(target.row, 10).Value2)) > 0 Then Cancel = True: CompactEdit
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    CompactPrepareForSave
    GoTo Done
Failed:
    Cancel = True
    MsgBox "Dosya kaydedilemedi: " & Err.Description, vbExclamation
Done:
    Application.EnableEvents = oldEvents
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Dim reportProblem As String
    On Error GoTo Failed
    If ActiveSheet.name = "Raporlar" Then
        reportProblem = CompactReportRefresh()
        If Len(reportProblem) > 0 Then Err.Raise vbObjectError + 311, , reportProblem
    End If
    CompactPrintLayout
    Exit Sub
Failed:
    Cancel = True
    MsgBox "Yazdırma alanı hazırlanamadı: " & Err.Description, vbExclamation
End Sub

Attribute VB_Name = "Module1"
Option Explicit

Private Function CC_Sheet() As Worksheet
    Set CC_Sheet = ThisWorkbook.Worksheets("İzin Girişi")
End Function

Private Function CC_Last(ByVal columnNumber As Long) As Long
    With CC_Sheet
        CC_Last = .Cells(.Rows.count, columnNumber).End(xlUp).row
    End With
End Function

Private Function CC_Dict() As Object
    Set CC_Dict = CreateObject("Scripting.Dictionary")
    CC_Dict.CompareMode = vbTextCompare
End Function

Private Function CC_Text(ByVal value As Variant) As String
    If Not IsError(value) And Not IsEmpty(value) Then CC_Text = Trim$(CStr(value))
End Function

Private Function CC_Bool(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbBoolean Then
        CC_Bool = CBool(value)
    ElseIf IsNumeric(value) Then
        CC_Bool = CDbl(value) <> 0
    Else
        CC_Bool = UCase$(CC_Text(value)) = "TRUE" Or UCase$(CC_Text(value)) = "EVET"
    End If
End Function

Private Function CC_Normal(ByVal value As String) As String
    value = Replace(Replace(Replace(value, "İ", "i"), "I", "i"), "ı", "i")
    value = LCase$(Trim$(value))
    value = Replace(Replace(Replace(value, "ş", "s"), "ğ", "g"), "ü", "u")
    CC_Normal = Replace(Replace(Replace(value, "ö", "o"), "ç", "c"), "_", " ")
End Function

Private Function CC_Active(ByVal value As Variant) As Boolean
    CC_Active = CC_Normal(CC_Text(value)) <> "iptal"
End Function

Private Function CC_DateOK(ByVal value As Variant) As Boolean
    If IsError(value) Or Not IsNumeric(value) Then Exit Function
    CC_DateOK = CDbl(value) >= CDbl(DateSerial(2000, 1, 1)) And CDbl(value) < CDbl(DateSerial(2101, 1, 1))
End Function

Public Function CompactTypeCode(ByVal value As String) As String
    Select Case CC_Normal(value)
        Case "yillik", "yillik izin": CompactTypeCode = "YILLIK"
        Case "rapor", "raporlu", "saglik raporu", "rapor izni": CompactTypeCode = "RAPOR"
        Case "mazeret", "mazeret izni": CompactTypeCode = "MAZERET"
        Case "babalik", "babalik izni": CompactTypeCode = "BABALIK"
        Case "evlilik", "evlilik izni": CompactTypeCode = "EVLILIK"
        Case "vefat", "vefat izni", "olum", "olum izni": CompactTypeCode = "VEFAT"
        Case "analik", "analik izni", "dogum izni": CompactTypeCode = "ANALIK"
        Case "refakat", "refakat izni": CompactTypeCode = "REFAKAT"
        Case "ucretsiz", "ucretsiz izin", "ayliksiz", "ayliksiz izin": CompactTypeCode = "UCRETSIZ"
        Case "idari", "idari izin": CompactTypeCode = "IDARI"
        Case "gorevli", "gorevlendirme": CompactTypeCode = "GOREVLI"
        Case "diger", "diger izin", "devamsizlik", "diger / devamsizlik": CompactTypeCode = "DIGER"
    End Select
End Function

Public Function CompactTypeLabel(ByVal value As String) As String
    Select Case CompactTypeCode(value)
        Case "YILLIK": CompactTypeLabel = "Yıllık İzin"
        Case "RAPOR": CompactTypeLabel = "Rapor"
        Case "MAZERET": CompactTypeLabel = "Mazeret İzni"
        Case "BABALIK": CompactTypeLabel = "Babalık İzni"
        Case "EVLILIK": CompactTypeLabel = "Evlilik İzni"
        Case "VEFAT": CompactTypeLabel = "Vefat İzni"
        Case "ANALIK": CompactTypeLabel = "Analık İzni"
        Case "REFAKAT": CompactTypeLabel = "Refakat İzni"
        Case "UCRETSIZ": CompactTypeLabel = "Aylıksız İzin"
        Case "IDARI": CompactTypeLabel = "İdari İzin"
        Case "GOREVLI": CompactTypeLabel = "Görevlendirme"
        Case "DIGER": CompactTypeLabel = "Diğer / Devamsızlık"
    End Select
End Function

Public Function CompactPersonId(ByVal sicil As String) As String
    Dim people As Variant, r As Long, lastRow As Long
    lastRow = CC_Last(2)
    If lastRow < 10 Then Exit Function
    people = CC_Sheet.Range("B10:I" & CStr(lastRow)).Value2
    For r = 1 To UBound(people, 1)
        If CC_Text(people(r, 1)) = Trim$(sicil) Then CompactPersonId = CC_Text(people(r, 8)): Exit Function
    Next r
End Function

Private Function CC_Holidays(ByVal yr As Long, ByVal mn As Long) As Object
    Dim dates As Variant, r As Long, serial As Long, result As Object
    Set result = CC_Dict()
    dates = ThisWorkbook.Worksheets("Puantaj").Range("K10:K109").Value2
    For r = 1 To UBound(dates, 1)
        If CC_DateOK(dates(r, 1)) Then
            serial = CLng(Fix(CDbl(dates(r, 1))))
            If VBA.Year(CDate(serial)) = yr And VBA.Month(CDate(serial)) = mn Then result(CStr(serial)) = True
        End If
    Next r
    Set CC_Holidays = result
End Function

Public Function CompactBaseDays(ByVal yr As Long, ByVal mn As Long) As Long
    Dim holidays As Object, d As Long, lastDay As Long
    If yr < 2000 Or yr > 2100 Or mn < 1 Or mn > 12 Then Exit Function
    Set holidays = CC_Holidays(yr, mn)
    lastDay = CLng(DateSerial(yr, mn + 1, 0))
    For d = CLng(DateSerial(yr, mn, 1)) To lastDay
        If Weekday(CDate(d), vbMonday) <= 5 And Not holidays.Exists(CStr(d)) Then CompactBaseDays = CompactBaseDays + 1
    Next d
End Function

Public Function CompactAttendance(ByVal yr As Long, ByVal mn As Long) As Variant
    Dim people As Variant, leaves As Variant, output() As Variant
    Dim holidays As Object, working As Object, absent As Object, totals As Object
    Dim r As Long, d As Long, firstDay As Long, lastDay As Long, a As Long, b As Long
    Dim count As Long, base As Long, pid As String, key As String, item As Variant, lastRow As Long
    If yr < 2000 Or yr > 2100 Or mn < 1 Or mn > 12 Then Err.Raise 5, "CompactAttendance", "Yıl ve ayı kontrol edin."
    lastRow = CC_Last(2)
    If lastRow < 10 Then Exit Function
    people = CC_Sheet.Range("A10:I" & CStr(lastRow)).Value2
    Set holidays = CC_Holidays(yr, mn)
    Set working = CC_Dict(): Set absent = CC_Dict(): Set totals = CC_Dict()
    firstDay = CLng(DateSerial(yr, mn, 1)): lastDay = CLng(DateSerial(yr, mn + 1, 0))
    For d = firstDay To lastDay
        If Weekday(CDate(d), vbMonday) <= 5 And Not holidays.Exists(CStr(d)) Then working(CStr(d)) = True: base = base + 1
    Next d
    lastRow = CC_Last(18)
    If lastRow >= 2 Then
        leaves = CC_Sheet.Range("R2:AB" & CStr(lastRow)).Value2
        For r = 1 To UBound(leaves, 1)
            If CC_Active(leaves(r, 8)) Then
                If CC_DateOK(leaves(r, 4)) And CC_DateOK(leaves(r, 5)) Then
                    a = CLng(Fix(CDbl(leaves(r, 4)))): b = CLng(Fix(CDbl(leaves(r, 5))))
                    If a <= lastDay And b >= firstDay And b >= a Then
                        If a < firstDay Then a = firstDay
                        If b > lastDay Then b = lastDay
                        pid = CC_Text(leaves(r, 2))
                        For d = a To b
                            If working.Exists(CStr(d)) Then
                                key = pid & "|" & CStr(d)
                                If Not absent.Exists(key) Then
                                    absent(key) = True
                                    If Not totals.Exists(pid) Then totals(pid) = 0&
                                    totals(pid) = CLng(totals(pid)) + 1
                                End If
                            End If
                        Next d
                    End If
                End If
            End If
        Next r
    End If
    For r = 1 To UBound(people, 1)
        If Len(CC_Text(people(r, 2))) > 0 And Len(CC_Text(people(r, 9))) > 0 Then count = count + 1
    Next r
    If count = 0 Then Exit Function
    ReDim output(1 To count, 1 To 7)
    count = 0
    For r = 1 To UBound(people, 1)
        pid = CC_Text(people(r, 9))
        If Len(CC_Text(people(r, 2))) > 0 And Len(pid) > 0 Then
            count = count + 1
            output(count, 1) = count: output(count, 2) = people(r, 2): output(count, 3) = people(r, 3)
            output(count, 5) = 0&
            If totals.Exists(pid) Then output(count, 5) = CLng(totals(pid))
            output(count, 4) = base - CLng(output(count, 5)): output(count, 6) = base: output(count, 7) = people(r, 4)
        End If
    Next r
    CompactAttendance = output
End Function

Private Function CC_Key(ByVal pid As String, ByVal yr As Long) As String
    CC_Key = pid & "|" & CStr(yr)
End Function

Private Function CC_Charge(ByVal pid As String, ByVal yr As Long, ByVal days As Double, ByVal balances As Object, ByVal used As Object, ByVal distribution As Object) As String
    Dim grantYear As Long, key As String, take As Double
    For grantYear = yr - 1 To yr
        key = CC_Key(pid, grantYear)
        If balances.Exists(key) Then
            take = CDbl(balances(key))
            If take > days Then take = days
            If take > 0 Then
                balances(key) = Round(CDbl(balances(key)) - take, 2)
                used(key) = Round(CDbl(used(key)) + take, 2)
                If Not distribution.Exists(CStr(grantYear)) Then distribution(CStr(grantYear)) = 0#
                distribution(CStr(grantYear)) = Round(CDbl(distribution(CStr(grantYear))) + take, 2)
                days = Round(days - take, 2)
            End If
        End If
        If days < 0.000001 Then Exit For
    Next grantYear
    If days > 0.000001 Then CC_Charge = "Kullanılabilir yıllık izin bakiyesi yetersiz."
End Function

Private Sub CC_SortLeaves(ByRef indices() As Long, ByRef leaves As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Double, swap As Long
    i = low: j = high: pivot = CDbl(leaves(indices((low + high) \ 2), 4))
    Do While i <= j
        Do While CDbl(leaves(indices(i), 4)) < pivot: i = i + 1: Loop
        Do While CDbl(leaves(indices(j), 4)) > pivot: j = j - 1: Loop
        If i <= j Then
            swap = indices(i): indices(i) = indices(j): indices(j) = swap
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then CC_SortLeaves indices, leaves, low, j
    If i < high Then CC_SortLeaves indices, leaves, i, high
End Sub

Public Function CompactRecalcBalances() As String
    Dim ws As Worksheet, people As Variant, rights As Variant, leaves As Variant
    Dim balances As Object, used As Object, distribution As Object, indices() As Long
    Dim usedOut() As Variant, allocOut() As Variant, rosterOut() As Variant
    Dim oldUsed As Variant, oldAlloc As Variant, oldRoster As Variant, wrote As Boolean
    Dim lastPeople As Long, lastRights As Long, lastLeaves As Long, leftYear As Long, rightYear As Long
    Dim r As Long, i As Long, count As Long, d As Long, firstDay As Long, lastDay As Long, grantYear As Long
    Dim pid As String, key As String, amount As Double, result As String, text As String, failure As String
    On Error GoTo Failed
    Set ws = CC_Sheet
    lastPeople = CC_Last(2): lastRights = CC_Last(30): lastLeaves = CC_Last(18)
    If lastPeople < 10 Then Exit Function
    people = ws.Range("A10:I" & CStr(lastPeople)).Value2
    Set balances = CC_Dict(): Set used = CC_Dict()
    If lastRights >= 2 Then
        rights = ws.Range("AD2:AG" & CStr(lastRights)).Value2
        ReDim usedOut(1 To UBound(rights, 1), 1 To 1)
        For r = 1 To UBound(rights, 1)
            If Not IsNumeric(rights(r, 2)) Or Not IsNumeric(rights(r, 3)) Then CompactRecalcBalances = "ERR:İzin hakkı yılını ve gününü kontrol edin.": Exit Function
            If CDbl(rights(r, 3)) < 0 Then CompactRecalcBalances = "ERR:İzin hakkı negatif olamaz.": Exit Function
            key = CC_Key(CC_Text(rights(r, 1)), CLng(rights(r, 2)))
            If balances.Exists(key) Then CompactRecalcBalances = "ERR:Aynı personel/yıl için iki hak kaydı var.": Exit Function
            balances(key) = CDbl(rights(r, 3)): used(key) = 0#
        Next r
    End If
    If lastLeaves >= 2 Then
        leaves = ws.Range("R2:AB" & CStr(lastLeaves)).Value2
        ReDim allocOut(1 To UBound(leaves, 1), 1 To 1): ReDim indices(1 To UBound(leaves, 1))
        For r = 1 To UBound(leaves, 1)
            allocOut(r, 1) = leaves(r, 10)
            If CC_Active(leaves(r, 8)) And CC_Bool(leaves(r, 7)) And CompactTypeCode(CC_Text(leaves(r, 3))) = "YILLIK" Then
                If Not CC_DateOK(leaves(r, 4)) Or Not CC_DateOK(leaves(r, 5)) Or Not IsNumeric(leaves(r, 6)) Then CompactRecalcBalances = "ERR:İzin kaydının tarihini/gününü kontrol edin.": Exit Function
                firstDay = CLng(leaves(r, 4)): lastDay = CLng(leaves(r, 5)): amount = CDbl(leaves(r, 6))
                If lastDay < firstDay Or amount <= 0 Or amount > lastDay - firstDay + 1 Then CompactRecalcBalances = "ERR:İzin tarih aralığını ve gününü kontrol edin.": Exit Function
                If VBA.Year(CDate(firstDay)) <> VBA.Year(CDate(lastDay)) And Abs(amount - (lastDay - firstDay + 1)) > 0.000001 Then CompactRecalcBalances = "ERR:Yıl aşan ve gün sayısı elle değiştirilen izni iki kayıt olarak girin.": Exit Function
                count = count + 1: indices(count) = r
            ElseIf Not CC_Active(leaves(r, 8)) Then
                allocOut(r, 1) = "İptal"
            ElseIf Not CC_Bool(leaves(r, 7)) Then
                If Len(CC_Text(allocOut(r, 1))) = 0 Then allocOut(r, 1) = "Açılışa dâhil / tekrar düşülmez"
            End If
        Next r
        If count > 1 Then CC_SortLeaves indices, leaves, 1, count
        For i = 1 To count
            r = indices(i): pid = CC_Text(leaves(r, 2))
            firstDay = CLng(leaves(r, 4)): lastDay = CLng(leaves(r, 5)): amount = CDbl(leaves(r, 6))
            Set distribution = CC_Dict()
            If VBA.Year(CDate(firstDay)) = VBA.Year(CDate(lastDay)) Then
                result = CC_Charge(pid, VBA.Year(CDate(firstDay)), amount, balances, used, distribution)
            Else
                result = ""
                For d = firstDay To lastDay
                    result = CC_Charge(pid, VBA.Year(CDate(d)), 1, balances, used, distribution)
                    If Len(result) > 0 Then Exit For
                Next d
            End If
            If Len(result) > 0 Then CompactRecalcBalances = "ERR:" & CC_Text(leaves(r, 1)) & " - " & result: Exit Function
            text = ""
            For grantYear = VBA.Year(CDate(firstDay)) - 1 To VBA.Year(CDate(lastDay))
                If distribution.Exists(CStr(grantYear)) Then
                    If Len(text) > 0 Then text = text & ";"
                    text = text & CStr(grantYear) & "=" & CStr(distribution(CStr(grantYear)))
                End If
            Next grantYear
            allocOut(r, 1) = text
        Next i
    End If
    If lastRights >= 2 Then
        For r = 1 To UBound(rights, 1)
            usedOut(r, 1) = used(CC_Key(CC_Text(rights(r, 1)), CLng(rights(r, 2))))
        Next r
    End If
    leftYear = CompactLeftYear(): rightYear = CompactRightYear()
    ReDim rosterOut(1 To UBound(people, 1), 1 To 3)
    For r = 1 To UBound(people, 1)
        pid = CC_Text(people(r, 9))
        If Len(pid) > 0 Then
            rosterOut(r, 1) = 0#: rosterOut(r, 2) = 0#
            key = CC_Key(pid, leftYear): If balances.Exists(key) Then rosterOut(r, 1) = balances(key)
            key = CC_Key(pid, rightYear): If balances.Exists(key) Then rosterOut(r, 2) = balances(key)
            rosterOut(r, 3) = CDbl(rosterOut(r, 1)) + CDbl(rosterOut(r, 2))
        End If
    Next r
    oldRoster = ws.Range("E10:G" & CStr(lastPeople)).Value2
    If lastRights >= 2 Then oldUsed = ws.Range("AG2:AG" & CStr(lastRights)).Value2
    If lastLeaves >= 2 Then oldAlloc = ws.Range("AA2:AA" & CStr(lastLeaves)).Value2
    wrote = True
    If lastRights >= 2 Then ws.Range("AG2:AG" & CStr(lastRights)).Value2 = usedOut
    If lastLeaves >= 2 Then ws.Range("AA2:AA" & CStr(lastLeaves)).Value2 = allocOut
    CompactAssignBlock ws, "E10:G" & CStr(lastPeople), rosterOut
    Exit Function
Failed:
    failure = Err.Description
    On Error Resume Next
    If wrote Then
        CompactAssignBlock ws, "E10:G" & CStr(lastPeople), oldRoster
        If lastRights >= 2 Then ws.Range("AG2:AG" & CStr(lastRights)).Value2 = oldUsed
        If lastLeaves >= 2 Then ws.Range("AA2:AA" & CStr(lastLeaves)).Value2 = oldAlloc
    End If
    CompactRecalcBalances = "ERR:İzin bakiyesi güncellenemedi: " & failure
End Function

Public Function CompactOpening(ByVal pid As String, ByVal yr As Long) As Double
    Dim rights As Variant, r As Long, lastRow As Long
    lastRow = CC_Last(30): If lastRow < 2 Then Exit Function
    rights = CC_Sheet.Range("AD2:AG" & CStr(lastRow)).Value2
    For r = 1 To UBound(rights, 1)
        If CC_Text(rights(r, 1)) = pid And CLng(rights(r, 2)) = yr Then CompactOpening = CDbl(rights(r, 3)): Exit Function
    Next r
End Function

Public Function CompactBalance(ByVal pid As String, ByVal yr As Long) As Double
    Dim rights As Variant, r As Long, lastRow As Long
    lastRow = CC_Last(30): If lastRow < 2 Then Exit Function
    rights = CC_Sheet.Range("AD2:AG" & CStr(lastRow)).Value2
    For r = 1 To UBound(rights, 1)
        If CC_Text(rights(r, 1)) = pid And CLng(rights(r, 2)) = yr Then CompactBalance = CDbl(rights(r, 3)) - CDbl(rights(r, 4)): Exit Function
    Next r
End Function

Public Function SaveCompactLeave(ByVal pid As String, ByVal typeCode As String, ByVal startSerial As Double, ByVal endSerial As Double, Optional ByVal manualDays As Double = -1, Optional ByVal editId As String = "") As String
    Dim ws As Worksheet, people As Variant, leaves As Variant, oldRow As Variant, proposed(1 To 1, 1 To 11) As Variant
    Dim r As Long, c As Long, lastRow As Long, targetRow As Long, n As Long, found As Boolean, imported As Boolean
    Dim id As String, code As String, days As Double, result As String, failure As String, wrote As Boolean, preserveOverlap As Boolean
    On Error GoTo Failed
    Set ws = CC_Sheet
    lastRow = CC_Last(2)
    If lastRow >= 10 Then
        people = ws.Range("A10:I" & CStr(lastRow)).Value2
        For r = 1 To UBound(people, 1)
            If CC_Text(people(r, 9)) = pid Then found = True: Exit For
        Next r
    End If
    If Not found Or Len(pid) = 0 Then SaveCompactLeave = "ERR:Önce bir personel seçin.": Exit Function
    code = CompactTypeCode(typeCode)
    If Len(code) = 0 Then SaveCompactLeave = "ERR:Geçerli bir izin türü seçin.": Exit Function
    If Not CC_DateOK(startSerial) Or Not CC_DateOK(endSerial) Then SaveCompactLeave = "ERR:Başlangıç ve bitiş tarihlerini kontrol edin.": Exit Function
    startSerial = Fix(startSerial): endSerial = Fix(endSerial)
    If endSerial < startSerial Then SaveCompactLeave = "ERR:Bitiş tarihi başlangıçtan önce olamaz.": Exit Function
    days = endSerial - startSerial + 1
    If manualDays <> -1 Then days = manualDays
    If days <= 0 Or Abs(days * 2 - Round(days * 2, 0)) > 0.000001 Then SaveCompactLeave = "ERR:İzin gününü tarih aralığını aşmadan tam veya yarım gün girin.": Exit Function
    If days > endSerial - startSerial + 1 Then
        If Not CC_ImportedValuesUnchanged(pid, code, startSerial, endSerial, days, editId) Then SaveCompactLeave = "ERR:İzin gününü tarih aralığını aşmadan tam veya yarım gün girin.": Exit Function
    End If
    If code = "YILLIK" And VBA.Year(CDate(startSerial)) <> VBA.Year(CDate(endSerial)) And Abs(days - (endSerial - startSerial + 1)) > 0.000001 Then SaveCompactLeave = "ERR:Yıl aşan ve gün sayısı elle değiştirilen izni iki kayıt olarak girin.": Exit Function
    lastRow = CC_Last(18)
    If lastRow >= 2 Then
        leaves = ws.Range("R2:AB" & CStr(lastRow)).Value2
        If Len(editId) > 0 Then
            For r = 1 To UBound(leaves, 1)
                If CC_Text(leaves(r, 1)) = editId Then
                    targetRow = r + 1
                    If CC_Text(leaves(r, 2)) <> pid Then SaveCompactLeave = "ERR:İzin başka personele ait.": Exit Function
                    imported = Not CC_Bool(leaves(r, 7)) And InStr(1, CC_Normal(CC_Text(leaves(r, 11))), "kaynak", vbBinaryCompare) > 0
                    preserveOverlap = imported And CDbl(leaves(r, 4)) = startSerial And CDbl(leaves(r, 5)) = endSerial
                    Exit For
                End If
            Next r
            If targetRow = 0 Then SaveCompactLeave = "ERR:Düzenlenecek izin bulunamadı.": Exit Function
        End If
        If Not preserveOverlap Then
            For r = 1 To UBound(leaves, 1)
                If CC_Text(leaves(r, 2)) = pid And CC_Text(leaves(r, 1)) <> editId And CC_Active(leaves(r, 8)) Then
                    If CC_DateOK(leaves(r, 4)) And CC_DateOK(leaves(r, 5)) Then
                        If CDbl(leaves(r, 4)) <= endSerial And CDbl(leaves(r, 5)) >= startSerial Then SaveCompactLeave = "ERR:Bu tarih aralığında başka izin var: " & CC_Text(leaves(r, 1)): Exit Function
                    End If
                End If
            Next r
        End If
    ElseIf Len(editId) > 0 Then
        SaveCompactLeave = "ERR:Düzenlenecek izin bulunamadı.": Exit Function
    End If
    If targetRow = 0 Then
        targetRow = Application.Max(2, lastRow + 1)
        id = CompactNextLeaveId()
        proposed(1, 11) = "Elle"
    Else
        oldRow = ws.Range("R" & targetRow & ":AB" & targetRow).Value2
        For c = 1 To 11: proposed(1, c) = oldRow(1, c): Next c
        id = editId
    End If
    proposed(1, 1) = id: proposed(1, 2) = pid: proposed(1, 3) = code
    proposed(1, 4) = startSerial: proposed(1, 5) = endSerial: proposed(1, 6) = days
    proposed(1, 7) = code = "YILLIK" And Not imported: proposed(1, 8) = "Onaylı"
    wrote = True
    ws.Range("R" & targetRow & ":AB" & targetRow).Value2 = proposed
    result = CompactRecalcBalances()
    If Len(result) > 0 Then
        If IsEmpty(oldRow) Then ws.Range("R" & targetRow & ":AB" & targetRow).ClearContents Else ws.Range("R" & targetRow & ":AB" & targetRow).Value2 = oldRow
        failure = CompactRecalcBalances()
        SaveCompactLeave = result: Exit Function
    End If
    SaveCompactLeave = id
    CompactReportsDirty
    Exit Function
Failed:
    failure = Err.Description
    On Error Resume Next
    If wrote Then
        If IsEmpty(oldRow) Then ws.Range("R" & targetRow & ":AB" & targetRow).ClearContents Else ws.Range("R" & targetRow & ":AB" & targetRow).Value2 = oldRow
        result = CompactRecalcBalances()
    End If
    SaveCompactLeave = "ERR:İzin kaydedilemedi: " & failure
End Function

Public Function CancelCompactLeave(ByVal id As String) As String
    CancelCompactLeave = CompactDeleteLeave(id)
End Function

Public Function CompactSetOpening(ByVal pid As String, ByVal yr As Long, ByVal amount As Double) As String
    Dim ws As Worksheet, rights As Variant, people As Variant, proposed(1 To 1, 1 To 4) As Variant, oldRow As Variant
    Dim lastRow As Long, r As Long, targetRow As Long, found As Boolean, result As String, failure As String, wrote As Boolean
    On Error GoTo Failed
    If amount < 0 Or Abs(amount * 2 - Round(amount * 2, 0)) > 0.000001 Then CompactSetOpening = "ERR:İzin hakkını sıfır veya pozitif tam/yarım gün girin.": Exit Function
    If yr < 2000 Or yr > 2100 Then CompactSetOpening = "ERR:Hakediş yılını kontrol edin.": Exit Function
    Set ws = CC_Sheet
    If CC_Last(2) >= 10 Then
        people = ws.Range("A10:I" & CStr(CC_Last(2))).Value2
        For r = 1 To UBound(people, 1)
            If CC_Text(people(r, 9)) = pid Then found = True: Exit For
        Next r
    End If
    If Not found Then CompactSetOpening = "ERR:Personel bulunamadı.": Exit Function
    lastRow = CC_Last(30)
    If lastRow >= 2 Then
        rights = ws.Range("AD2:AG" & CStr(lastRow)).Value2
        For r = 1 To UBound(rights, 1)
            If CC_Text(rights(r, 1)) = pid And CLng(rights(r, 2)) = yr Then targetRow = r + 1: Exit For
        Next r
    End If
    If targetRow = 0 Then
        targetRow = Application.Max(2, lastRow + 1)
    Else
        oldRow = ws.Range("AD" & targetRow & ":AG" & targetRow).Value2
    End If
    proposed(1, 1) = pid: proposed(1, 2) = yr: proposed(1, 3) = amount: proposed(1, 4) = 0#
    wrote = True
    ws.Range("AD" & targetRow & ":AG" & targetRow).Value2 = proposed
    result = CompactRecalcBalances()
    If Len(result) > 0 Then
        If IsEmpty(oldRow) Then ws.Range("AD" & targetRow & ":AG" & targetRow).ClearContents Else ws.Range("AD" & targetRow & ":AG" & targetRow).Value2 = oldRow
        failure = CompactRecalcBalances()
        CompactSetOpening = result
    End If
    Exit Function
Failed:
    failure = Err.Description
    On Error Resume Next
    If wrote Then
        If IsEmpty(oldRow) Then ws.Range("AD" & targetRow & ":AG" & targetRow).ClearContents Else ws.Range("AD" & targetRow & ":AG" & targetRow).Value2 = oldRow
        result = CompactRecalcBalances()
    End If
    CompactSetOpening = "ERR:İzin hakkı kaydedilemedi: " & failure
End Function

' Compact UI: two sheets, bounded ranges, no calculation when selecting a person.
Public Sub CompactOpen()
    Dim ws As Worksheet, oldEvents As Boolean, oldScreen As Boolean
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:="2022-tiski", UserInterfaceOnly:=True, AllowFiltering:=True
        ws.EnableSelection = xlNoRestrictions
    Next ws
    If Not CU_ValidPeriod Then
        With ThisWorkbook.Worksheets("Puantaj")
            .Range("B5").Value2 = Year(Date): .Range("D5").Value2 = Month(Date)
        End With
    End If
    CompactEnsurePersonLinks
    CU_YearLabels
    CompactHistoryInit
    CompactPrimeRoster
    CompactCalendarClose
    CU_RebindButtons
    CU_Search False, False
    CU_ShowSelected
    CompactRefreshAttendance
    GoTo Done
Failed:
    Application.StatusBar = "İzin dosyası açılamadı: " & Err.Description
Done:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
End Sub

Public Sub CompactHandleChange(ByVal Sh As Object, ByVal target As Range)
    Dim problem As String, remainingEditAddress As String, remainingEdits As Range
    On Error GoTo Failed
    If Sh.name = "Puantaj" Then
        If Not Intersect(target, Sh.Range("B5,D5,K10:K109")) Is Nothing Then
            If Not CU_ValidPeriod Then
                MsgBox "Yılı 2000-2100, ayı 1-12 arasında tam sayı olarak girin.", vbExclamation
                Sh.Range("B5").Value2 = Year(Date): Sh.Range("D5").Value2 = Month(Date)
            End If
            CompactRefreshAttendance
            CompactReportsDirty
        End If
    ElseIf Sh.name = "İzin Girişi" Then
        If Not Intersect(target, Sh.Range("E9:F9")) Is Nothing Then
            problem = CompactEditYears(target)
            CU_YearLabels
            CompactPrimeRoster
            CompactReportsDirty
            If Len(problem) > 0 Then MsgBox problem, vbExclamation, "Yıl bilgisi"
        ElseIf Not Intersect(target, Sh.Range("B10:F1048576")) Is Nothing Then
            Set remainingEdits = Intersect(target, Sh.Range("E10:F1048576"))
            If Not remainingEdits Is Nothing Then remainingEditAddress = remainingEdits.address
            If Not Intersect(target, Sh.Range("B10:D1048576")) Is Nothing Then
                problem = CompactSyncRoster(target)
                If Len(problem) > 0 Then MsgBox problem, vbExclamation, "Personel bilgisi": Exit Sub
            End If
            If Len(remainingEditAddress) > 0 Then
                problem = CompactEditRemaining(Sh.Range(remainingEditAddress))
                If Len(problem) > 0 Then MsgBox problem, vbExclamation, "Kalan izin"
            End If
            CompactPrimeRoster
            CU_ShowSelected
            CompactReportsDirty
        ElseIf Not Intersect(target, Sh.Range("C5")) Is Nothing Then
            CU_Search
        ElseIf Not Intersect(target, Sh.Range("J24,L24")) Is Nothing Then
            CompactHistoryReset
        ElseIf Not Intersect(target, Sh.Range("J9,M9,L12")) Is Nothing Then
            CU_Preview
        End If
    ElseIf Sh.name = "Raporlar" Then
        If Not Intersect(target, Sh.Range("B5,F5,B7,F7,B9,F9,B11,F11")) Is Nothing Then CompactReportsDirty
    End If
    Exit Sub
Failed:
    MsgBox "Giriş kontrolü: " & Err.Description, vbExclamation
End Sub

Public Sub CompactSelectPerson(ByVal rosterRow As Long)
    Dim ws As Worksheet, oldEvents As Boolean, errorText As String
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    If rosterRow < 10 Or rosterRow > CU_LastRow(ws, 2) Then Exit Sub
    If Len(CStr(ws.Cells(rosterRow, 9).Value2)) = 0 Then Exit Sub
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    ws.Range("AI2").Value2 = CStr(ws.Cells(rosterRow, 9).Value2)
    ws.Range("C5").Value2 = CStr(ws.Cells(rosterRow, 2).Value2) & " - " & CStr(ws.Cells(rosterRow, 3).Value2)
    CompactSearchClose
    CompactClear
    CompactHistoryReset False
    CU_ShowSelected
    ws.Activate
    ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
    ws.Range("J9").Select
Done:
    Application.EnableEvents = oldEvents
    If Len(errorText) > 0 Then MsgBox "Personel seçilemedi: " & errorText, vbExclamation
    Exit Sub
Failed:
    errorText = Err.Description
    Resume Done
End Sub

Public Sub CompactClear()
    Dim ws As Worksheet, oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Done
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    CompactCalendarClose
    ws.Range("AI3").ClearContents
    CU_Clear ws.Range("J9"): CU_Clear ws.Range("M9")
    CU_Clear ws.Range("L14"): CU_Clear ws.Range("J16")
    ws.Range("L12").Value2 = "Yıllık İzin"
    ws.Range("J5").Value2 = "YENİ İZİN"
Done:
    Application.EnableEvents = oldEvents
End Sub

Public Sub CompactSave()
    CU_Action "save"
End Sub

Public Sub CompactCancel()
    CU_Action "cancel"
End Sub

Public Sub CompactEdit()
    CU_Action "edit"
End Sub

Private Sub CU_Action(ByVal action As String)
    Dim ws As Worksheet, oldEvents As Boolean, oldScreen As Boolean, id As String, answer As String
    Dim pid As String, firstDay As Double, lastDay As Double, days As Double, code As String, r As Long
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    pid = CStr(ws.Range("AI2").Value2)
    If Len(pid) = 0 Then MsgBox "Önce personelin yanındaki + işaretine tıklayın.", vbExclamation: GoTo Done
    id = CStr(ws.Range("AI3").Value2)
    If action = "save" Then
        firstDay = CU_Date(ws.Range("J9").Value2): lastDay = CU_Date(ws.Range("M9").Value2)
        If firstDay <= 0 Or lastDay < firstDay Then MsgBox "Başlangıç ve bitiş tarihlerini seçin. Bitiş başlangıçtan önce olamaz.", vbExclamation: GoTo Done
        code = CU_TypeCode(CStr(ws.Range("L12").Value2))
        If Len(code) = 0 Then MsgBox "İzin türünü seçin.", vbExclamation: GoTo Done
        days = -1
        If Len(CStr(ws.Range("L14").Value2)) > 0 Then
            If Not IsNumeric(ws.Range("L14").Value2) Then MsgBox "İzin gününü sayı olarak girin.", vbExclamation: GoTo Done
            days = CDbl(ws.Range("L14").Value2)
        End If
        answer = SaveCompactLeave(pid, code, firstDay, lastDay, days, id)
        If left$(answer, 3) = "ERR" Then MsgBox answer, vbExclamation, "İzin kaydedilemedi": GoTo Done
        ws.Range("AI3").Value2 = answer
        r = CU_LeaveRow(answer)
        If r > 0 Then ws.Cells(r, 26).Value2 = CStr(ws.Range("J16").Value2)
        CompactPrimeRoster
        CU_ShowSelected
        CompactRefreshAttendance
        CompactReportsDirty
        ws.Range("J5").Value2 = "KAYITLI İZİN  /  " & answer
        MsgBox "İzin kaydedildi; puantaj güncellendi.", vbInformation
    Else
        If ActiveSheet.name = ws.name And ActiveCell.row >= 26 And ActiveCell.row <= 35 And ActiveCell.column >= 10 And ActiveCell.column <= 15 Then
            id = CStr(ws.Cells(ActiveCell.row, 10).Value2)
        End If
        r = CU_LeaveRow(id)
        If r = 0 Then MsgBox "Sağdaki son izinler listesinden bir kayıt seçin.", vbExclamation: GoTo Done
        If CStr(ws.Cells(r, 19).Value2) <> pid Then MsgBox "Seçili personele ait bir izin seçin.", vbExclamation: GoTo Done
        If action = "cancel" Then
            If MsgBox("""" & id & """ numaralı izin kaydı silinsin mi?", vbQuestion + vbYesNo + vbDefaultButton2, "İzin Kaydını Sil") <> vbYes Then GoTo Done
            answer = CompactDeleteLeave(id)
            If left$(answer, 3) = "ERR" Then MsgBox answer, vbExclamation: GoTo Done
            CompactClear
            CompactPrimeRoster
            CU_ShowSelected
            CompactRefreshAttendance
            CompactReportsDirty
            MsgBox "İzin kaydı silindi; puantaj güncellendi.", vbInformation
        Else
            ws.Range("AI3").Value2 = id
            ws.Range("J9").Value2 = ws.Cells(r, 21).Value2
            ws.Range("M9").Value2 = ws.Cells(r, 22).Value2
            ws.Range("L12").Value2 = CU_TypeLabel(CStr(ws.Cells(r, 20).Value2))
            ws.Range("L14").Value2 = ws.Cells(r, 23).Value2
            ws.Range("J16").Value2 = ws.Cells(r, 26).Value2
            ws.Range("J5").Value2 = "İZİN DÜZENLE  /  " & id
            ws.Activate: ActiveWindow.ScrollRow = 1
            ws.Range("J9").Select
        End If
    End If
    GoTo Done
Failed:
    MsgBox "İşlem tamamlanamadı: " & Err.Description, vbExclamation
Done:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
End Sub

Public Sub CompactRefreshAttendance()
    Dim ws As Worksheet, data As Variant, count As Long, yr As Long, mn As Long, oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("Puantaj")
    If Not CU_ValidPeriod Then Err.Raise vbObjectError + 200, , "Yıl veya ayı kontrol edin."
    yr = CLng(ws.Range("B5").Value2): mn = CLng(ws.Range("D5").Value2)
    data = CompactAttendance(yr, mn)
    count = 0
    If IsArray(data) Then
        count = UBound(data, 1)
    End If
    CompactAssignBlock ws, "A10:G" & CStr(Application.Max(10, 9 + count)), data, "A10:G" & Application.Max(10, CU_LastRow(ws, 2))
    ws.Range("F5").Value2 = CompactBaseDays(yr, mn)
    ws.Range("H5").Value2 = count
    GoTo Done
Failed:
    If Not ws Is Nothing Then ws.Range("F5").Value2 = "Kontrol"
    Application.EnableEvents = oldEvents
    Err.Raise vbObjectError + 201, "CompactRefreshAttendance", Err.Description
Done:
    Application.EnableEvents = oldEvents
End Sub

Public Sub CompactSearchClear()
    Dim ws As Worksheet, oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Done
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    CU_Clear ws.Range("C5")
    CU_Search False, False
Done:
    Application.EnableEvents = oldEvents
End Sub

Private Sub CU_Search(Optional ByVal choosePerson As Boolean = True, Optional ByVal showResults As Boolean = True)
    Dim ws As Worksheet, term As String, data As Variant, matches As Collection, result() As Variant
    Dim r As Long, i As Long, chosen As Long, label As String, exact As Long, lastRow As Long
    Dim ids() As Variant, unused As Variant
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    term = CC_Normal(CStr(ws.Range("C5").Value2))
    Set matches = New Collection
    lastRow = CU_LastRow(ws, 2)
    If lastRow >= 10 Then
        data = ws.Range("A10:I" & CStr(lastRow)).Value2
        For r = 1 To UBound(data, 1)
            If Len(CC_Text(data(r, 9))) > 0 Then
                label = CC_Text(data(r, 2)) & " - " & CC_Text(data(r, 3))
                If term = CC_Normal(CC_Text(data(r, 2))) Or term = CC_Normal(label) Then exact = r
                If Len(term) = 0 Or InStr(1, CC_Normal(label), term, vbBinaryCompare) > 0 Then matches.Add r
            End If
        Next r
    End If
    ws.Range("AT2:AU" & CStr(Application.Max(lastRow + 2, ws.Cells(ws.Rows.count, 47).End(xlUp).row))).ClearContents
    If matches.count > 0 Then
        ReDim result(1 To matches.count, 1 To 2)
        ReDim ids(1 To matches.count, 1 To 1)
        For i = 1 To matches.count
            r = CLng(matches(i))
            result(i, 1) = data(r, 9)
            ids(i, 1) = data(r, 9)
            result(i, 2) = CC_Text(data(r, 2)) & " - " & CC_Text(data(r, 3))
        Next i
        ws.Range("AT2:AT" & CStr(matches.count + 1)).Value2 = ids
        unused = CS_ResultCache(result)
    Else
        unused = CS_ResultCache(Empty)
    End If
    On Error Resume Next
    ThisWorkbook.names("CompactPersonSearch").Delete
    On Error GoTo 0
    ws.columns("AT:AU").Hidden = True
    CU_Clear ws.Range("D6")
    If Len(term) = 0 Then
        ws.Range("A7").Value2 = "İsim veya sicil yazıp Enter'a basın. ARA ile tüm personeli listeleyebilirsiniz."
    ElseIf matches.count = 0 Then
        ws.Range("A7").Value2 = "Sonuç bulunamadı. Farklı bir isim veya sicil deneyin."
    Else
        ws.Range("A7").Value2 = CStr(matches.count) & " personel bulundu · Seçmek için listeden bir isme tıklayın."
    End If
    If choosePerson And Len(term) > 0 Then
        chosen = exact
        If chosen = 0 And matches.count = 1 Then chosen = CLng(matches(1))
        If chosen > 0 Then
            CompactSearchClose
            ws.Range("C5").Value2 = CC_Text(data(chosen, 2)) & " - " & CC_Text(data(chosen, 3))
            CompactSelectPerson chosen + 9
            ws.Range("A7").Value2 = "Seçildi: " & CStr(ws.Range("C5").Value2) & " · Başlangıç ve bitiş tarihlerini seçebilirsiniz."
            Exit Sub
        Else
            ws.Range("AI2").ClearContents
            CompactClear
            CU_ShowSelected
        End If
    End If
    If showResults And matches.count > 0 And (Len(term) > 0 Or Not choosePerson) Then
        CS_ShowList matches.count
    Else
        CompactSearchClose
    End If
End Sub


Private Sub CU_ShowSelected()
    Dim ws As Worksheet, people As Variant, data As Variant, pid As String, r As Long, i As Long
    Dim lastPeople As Long, lastLeave As Long, found As Boolean, selectedYear As Long, selectedType As String
    Dim page As Long, pageCount As Long, count As Long
    Dim firstIndex As Long, lastIndex As Long, shown As Long, indices() As Long, output(1 To 10, 1 To 6) As Variant
    Dim wasProtected As Boolean, oldEvents As Boolean, errorNumber As Long, errorText As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects
    If wasProtected Then ws.Unprotect Password:="2022-tiski"
    ws.Range("L36").MergeArea.NumberFormat = "@"
    ws.Range("O26:O35").ClearComments
    ws.Range("J26:O35").ClearContents
    ws.Range("N24").Value2 = "0 kayıt"
    ws.Range("L36").Value2 = "0 / 0"
    ws.Range("J6").Value2 = "Personelin yanındaki + işaretine tıklayın."
    pid = CC_Text(ws.Range("AI2").Value2)
    If Len(pid) = 0 Then GoTo Done
    lastPeople = CU_LastRow(ws, 2)
    If lastPeople < 10 Then GoTo Done
    people = ws.Range("A10:I" & CStr(lastPeople)).Value2
    For r = 1 To UBound(people, 1)
        If CC_Text(people(r, 9)) = pid Then
            ws.Range("J6").Value2 = CC_Text(people(r, 2)) & " - " & CC_Text(people(r, 3))
            found = True
            Exit For
        End If
    Next r
    If Not found Then GoTo Done
    selectedYear = CH_ViewYear(ws.Range("J24").Value2)
    selectedType = ""
    If Not CH_All(ws.Range("L24").Value2) Then
        selectedType = CompactTypeCode(CC_Text(ws.Range("L24").Value2))
        If Len(selectedType) = 0 Then selectedType = "__INVALID__"
    End If
    lastLeave = CU_LastRow(ws, 18)
    data = ws.Range("R2:AB" & CStr(lastLeave)).Value2
    ReDim indices(1 To UBound(data, 1))
    For r = 1 To UBound(data, 1)
        If CC_Text(data(r, 2)) = pid And CC_Active(data(r, 8)) Then
            If Len(selectedType) = 0 Or CompactTypeCode(CC_Text(data(r, 3))) = selectedType Then
                If selectedYear = 0 Or Year(CDate(CU_Date(data(r, 4)))) = selectedYear Then
                    count = count + 1
                    indices(count) = r
                End If
            End If
        End If
    Next r
    ws.Range("N24").Value2 = CStr(count) & " kayıt"
    If count = 0 Then
        ws.Range("AI11").Value2 = 1
        GoTo Done
    End If
    If count > 1 Then CH_SortIndices indices, data, 1, count
    pageCount = (count + 9) \ 10
    page = CH_Page(ws.Range("AI11").Value2)
    If page > pageCount Then page = pageCount
    ws.Range("AI11").Value2 = page
    firstIndex = (page - 1) * 10 + 1
    lastIndex = Application.Min(count, firstIndex + 9)
    For i = firstIndex To lastIndex
        r = indices(i): shown = shown + 1
        output(shown, 1) = data(r, 1): output(shown, 2) = CU_TypeLabel(CStr(data(r, 3)))
        output(shown, 3) = data(r, 4): output(shown, 4) = data(r, 5)
        output(shown, 5) = data(r, 6): output(shown, 6) = data(r, 8)
    Next i
    ws.Range("J26:O35").Value2 = output
    For i = firstIndex To lastIndex
        r = indices(i)
        CU_HistoryNote ws.Cells(i - firstIndex + 26, 15), CStr(data(r, 9))
    Next i
    ws.Range("L36").Value2 = CStr(page) & " / " & CStr(pageCount)
Done:
    On Error GoTo 0
    If Not ws Is Nothing Then
        If wasProtected Then ws.Protect Password:="2022-tiski", UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    Application.EnableEvents = oldEvents
    If errorNumber <> 0 Then Err.Raise errorNumber, "CU_ShowSelected", errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume Done
End Sub

Private Sub CU_HistoryNote(ByVal target As Range, ByVal noteText As String)
    Dim note As Comment, lineCount As Long
    If Len(Trim$(Replace(Replace(noteText, vbCr, ""), vbLf, ""))) = 0 Then Exit Sub
    Set note = target.AddComment(text:=noteText)
    note.Visible = False
    lineCount = 1 + Len(noteText) \ 45 + Len(noteText) - Len(Replace(noteText, vbLf, ""))
    With note.Shape
        .TextFrame.AutoSize = False
        .TextFrame.Characters.Font.name = "Calibri"
        .TextFrame.Characters.Font.Size = 10
        .width = 280
        .Height = Application.Min(300, Application.Max(52, 26 + 14 * lineCount))
        .left = Application.Max(0, target.left + target.width - .width - 12)
        .top = target.top + target.Height
    End With
End Sub


Private Sub CU_Preview()
    Dim ws As Worksheet, a As Double, b As Double
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    a = CU_Date(ws.Range("J9").Value2): b = CU_Date(ws.Range("M9").Value2)
    If a > 0 And b >= a Then ws.Range("L14").Value2 = b - a + 1 Else CU_Clear ws.Range("L14")
End Sub

Private Function CU_LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    CU_LastRow = Application.Max(2, ws.Cells(ws.Rows.count, col).End(xlUp).row)
End Function

Private Function CU_LeaveRow(ByVal id As String) As Long
    Dim ws As Worksheet, data As Variant, r As Long
    If Len(id) = 0 Then Exit Function
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    data = ws.Range("R2:R" & CU_LastRow(ws, 18)).Value2
    If Not IsArray(data) Then
        If CStr(data) = id Then CU_LeaveRow = 2
        Exit Function
    End If
    For r = 1 To UBound(data, 1)
        If CStr(data(r, 1)) = id Then CU_LeaveRow = r + 1: Exit Function
    Next r
End Function

Private Sub CU_Clear(ByVal cell As Range)
    cell.MergeArea.ClearContents
End Sub

Private Function CU_Date(ByVal value As Variant) As Double
    On Error GoTo Done
    If Len(CStr(value)) = 0 Then Exit Function
    If IsNumeric(value) Then
        If CDbl(value) >= 36526 And CDbl(value) < 73416 Then CU_Date = Fix(CDbl(value))
    ElseIf IsDate(value) Then
        If Year(CDate(value)) >= 2000 And Year(CDate(value)) <= 2100 Then CU_Date = CDbl(DateValue(CDate(value)))
    End If
Done:
End Function

Private Function CU_ValidPeriod() As Boolean
    Dim y As Variant, m As Variant
    y = ThisWorkbook.Worksheets("Puantaj").Range("B5").Value2
    m = ThisWorkbook.Worksheets("Puantaj").Range("D5").Value2
    If Not IsNumeric(y) Or Not IsNumeric(m) Then Exit Function
    CU_ValidPeriod = (CDbl(y) = Fix(CDbl(y)) And CDbl(m) = Fix(CDbl(m)) And CDbl(y) >= 2000 And CDbl(y) <= 2100 And CDbl(m) >= 1 And CDbl(m) <= 12)
End Function

Private Sub CU_YearLabels()
    CompactRefreshYearHeaders
End Sub

Private Function CU_TypeCode(ByVal label As String) As String
    CU_TypeCode = CompactTypeCode(label)
End Function

Private Function CU_TypeLabel(ByVal code As String) As String
    CU_TypeLabel = CompactTypeLabel(code)
End Function

Private Sub CU_RebindButtons()
    Dim ws As Worksheet, shp As Shape, action As String, p As Long
    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            If left$(shp.name, 4) = "BTN_" Then
                action = shp.OnAction: p = InStrRev(action, "!")
                If p > 0 Then action = Mid$(action, p + 1)
                If Len(action) > 0 Then shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & action
            End If
        Next shp
    Next ws
End Sub

Public Sub CompactPrepareForSave()
    Dim i As Long, name As String, p As Long
    CompactCalendarClose
    CompactSearchClose
    For i = ThisWorkbook.names.count To 1 Step -1
        name = ThisWorkbook.names(i).name: p = InStrRev(name, "!")
        If p > 0 Then name = Mid$(name, p + 1)
        name = UCase$(Replace(Replace(name, "_xlnm.", "", 1, -1, vbTextCompare), "'", ""))
        If name = "PRINT_AREA" Or name = "PRINT_TITLES" Or name = "_FILTERDATABASE" Then ThisWorkbook.names(i).Delete
    Next i
End Sub

Public Sub CompactPrintLayout()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    ws.PageSetup.PrintArea = "$A$1:$H$" & CStr(CU_LastRow(ws, 2))
    ws.PageSetup.PrintTitleRows = "$9:$9"
    If ThisWorkbook.Worksheets.count >= 3 Then CompactReportPrintLayout
    Set ws = ThisWorkbook.Worksheets("Puantaj")
    ws.PageSetup.PrintArea = "$A$1:$G$" & CStr(CU_LastRow(ws, 2))
    ws.PageSetup.PrintTitleRows = "$9:$9"
End Sub

Public Sub CompactDateStart()
    CU_OpenCalendar "J9"
End Sub

Public Sub CompactDateEnd()
    CU_OpenCalendar "M9"
End Sub

Private Sub CU_OpenCalendar(ByVal target As String)
    Dim ws As Worksheet, d As Double
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    If Len(CStr(ws.Range("AI3").Value2)) > 0 Then d = CU_Date(ws.Range(target).Value2)
    If target = "M9" And d = 0 Then d = CU_Date(ws.Range("J9").Value2)
    If d = 0 Then d = CDbl(DateSerial(Year(Date), Month(Date), 1))
    ws.Range("AI5").Value2 = target
    ws.Range("AI6").Value2 = CDbl(DateSerial(Year(CDate(d)), Month(CDate(d)), 1))
    CU_DrawCalendar
End Sub

Public Sub CompactCalendarPrev()
    CU_ShiftMonth -1
End Sub

Public Sub CompactCalendarNext()
    CU_ShiftMonth 1
End Sub

Private Sub CU_ShiftMonth(ByVal delta As Long)
    Dim ws As Worksheet, d As Date
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    d = CDate(ws.Range("AI6").Value2)
    d = DateSerial(Year(d), Month(d) + delta, 1)
    If Year(d) < 2000 Or Year(d) > 2100 Then Exit Sub
    ws.Range("AI6").Value2 = CDbl(d)
    CU_DrawCalendar
End Sub

Public Sub CompactCalendarDay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    CompactApplyCalendarDate CDbl(ws.Shapes(CStr(Application.Caller)).AlternativeText)
End Sub

Public Sub CompactApplyCalendarDate(ByVal serial As Double)
    Dim ws As Worksheet, target As String, oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Done
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    target = CStr(ws.Range("AI5").Value2)
    If target <> "J9" And target <> "M9" Then GoTo Done
    If CU_Date(serial) = 0 Then GoTo Done
    ws.Range(target).Value2 = serial
    CompactCalendarClose
    CU_Preview
    If target = "J9" Then ws.Range("M9").Select
Done:
    Application.EnableEvents = oldEvents
End Sub

Public Sub CompactCalendarClose()
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    For i = ws.Shapes.count To 1 Step -1
        If left$(ws.Shapes(i).name, 3) = "DP_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub CU_DrawCalendar()
    Dim ws As Worksheet, area As Range, shp As Shape, x As Double, y As Double, d As Date, first As Long
    Dim names As Variant, months As Variant, i As Long, col As Long, row As Long, serial As Long, fill As Long, ink As Long
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    ws.Activate: CompactCalendarClose
    Set area = ActiveWindow.VisibleRange
    x = ws.Range("J9").left: y = ws.Range("J11").top + 3
    If x + 270 > area.left + area.width Then x = area.left + area.width - 274
    If x < area.left + 3 Then x = area.left + 3
    If y + 262 > area.top + area.Height Then y = area.top + area.Height - 265
    If y < area.top + 3 Then y = area.top + 3
    d = CDate(ws.Range("AI6").Value2)
    months = Split("Ocak,Şubat,Mart,Nisan,Mayıs,Haziran,Temmuz,Ağustos,Eylül,Ekim,Kasım,Aralık", ",")
    Set shp = CU_Shape("DP_Back", "", x, y, 270, 260, RGB(255, 255, 255), RGB(23, 62, 77), "")
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(160, 190, 198)
    Set shp = CU_Shape("DP_Prev", "<", x + 8, y + 8, 27, 28, RGB(236, 246, 248), RGB(8, 127, 140), "CompactCalendarPrev")
    Set shp = CU_Shape("DP_Title", months(Month(d) - 1) & " " & Year(d), x + 40, y + 8, 155, 28, RGB(236, 246, 248), RGB(23, 62, 77), "")
    Set shp = CU_Shape("DP_Next", ">", x + 200, y + 8, 27, 28, RGB(236, 246, 248), RGB(8, 127, 140), "CompactCalendarNext")
    Set shp = CU_Shape("DP_Close", "x", x + 235, y + 8, 27, 28, RGB(255, 237, 232), RGB(170, 70, 50), "CompactCalendarClose")
    names = Split("Pt,Sa,Ça,Pe,Cu,Ct,Pa", ",")
    For i = 0 To 6
        Set shp = CU_Shape("DP_W" & i, names(i), x + 8 + i * 36, y + 42, 34, 20, RGB(255, 255, 255), RGB(84, 110, 120), "")
    Next i
    first = CLng(d) - Weekday(d, vbMonday) + 1
    For i = 0 To 41
        col = i Mod 7: row = i \ 7: serial = first + i
        fill = RGB(241, 247, 249): ink = RGB(23, 62, 77)
        If col > 4 Then fill = RGB(254, 241, 236)
        If Month(CDate(serial)) <> Month(d) Then ink = RGB(170, 185, 190)
        If serial = CLng(Date) Then fill = RGB(219, 239, 231): ink = RGB(20, 112, 78)
        If serial = CLng(CU_Date(ws.Range(CStr(ws.Range("AI5").Value2)).Value2)) Then fill = RGB(8, 127, 140): ink = RGB(255, 255, 255)
        Set shp = CU_Shape("DP_D" & Format$(i, "00"), CStr(Day(CDate(serial))), x + 8 + col * 36, y + 66 + row * 30, 34, 27, fill, ink, "CompactCalendarDay")
        shp.AlternativeText = CStr(serial)
    Next i
End Sub

Private Function CU_Shape(ByVal name As String, ByVal text As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal fill As Long, ByVal ink As Long, ByVal macro As String) As Shape
    Dim shp As Shape
    Set shp = ThisWorkbook.Worksheets("İzin Girişi").Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    shp.name = name: shp.fill.ForeColor.RGB = fill: shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.text = text
    shp.TextFrame2.TextRange.Font.name = "Calibri": shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.fill.ForeColor.RGB = ink
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle: shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.MarginLeft = 0: shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 0: shp.TextFrame2.MarginBottom = 0
    shp.DrawingObject.PrintObject = False
    If Len(macro) > 0 Then shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & macro
    Set CU_Shape = shp
End Function

' Editable year window and remaining balances. Append to the compact standard module.
Private Function CY_ParseYear(ByVal value As Variant) As Long
    Dim text As String, amount As Double
    If IsError(value) Or IsEmpty(value) Then Exit Function
    text = Trim$(CStr(value))
    If Len(text) < 4 Then Exit Function
    If Not IsNumeric(left$(text, 4)) Then Exit Function
    If IsNumeric(text) Then
        amount = CDbl(text)
        If amount <> Fix(amount) Then Exit Function
    Else
        amount = CDbl(left$(text, 4))
        If Len(text) > 4 Then
            If Mid$(text, 5, 1) >= "0" And Mid$(text, 5, 1) <= "9" Then Exit Function
        End If
    End If
    If amount >= 2000 And amount <= 2100 Then CY_ParseYear = CLng(amount)
End Function

Public Sub CompactInitYears()
    Dim ws As Worksheet, a As Long, b As Long, oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = CC_Sheet
    a = CY_ParseYear(ws.Range("AI8").Value2): b = CY_ParseYear(ws.Range("AI9").Value2)
    If a < 2000 Or b <> a + 1 Then
        a = CY_ParseYear(ws.Range("E9").Value2): b = CY_ParseYear(ws.Range("F9").Value2)
        If a < 2000 Or b <> a + 1 Then
            b = CY_ParseYear(ThisWorkbook.Worksheets("Puantaj").Range("B5").Value2)
            If b < 2001 Then b = Year(Date)
            a = b - 1
        End If
        ws.Range("AI8").Value2 = a: ws.Range("AI9").Value2 = b
    End If
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    Application.EnableEvents = oldEvents
    Err.Raise Err.Number, "CompactInitYears", Err.Description
End Sub

Public Function CompactLeftYear() As Long
    CompactInitYears
    CompactLeftYear = CLng(CC_Sheet.Range("AI8").Value2)
End Function

Public Function CompactRightYear() As Long
    CompactInitYears
    CompactRightYear = CLng(CC_Sheet.Range("AI9").Value2)
End Function

Public Sub CompactRefreshYearHeaders()
    Dim oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    CompactInitYears
    CC_Sheet.Range("E9").Value2 = CStr(CompactLeftYear()) & " Kalan"
    CC_Sheet.Range("F9").Value2 = CStr(CompactRightYear()) & " Kalan"
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    Application.EnableEvents = oldEvents
    Err.Raise Err.Number, "CompactRefreshYearHeaders", Err.Description
End Sub

Public Sub CompactRefreshRemaining()
    Dim ws As Worksheet, people As Variant, rights As Variant, values() As Variant, balances As Object
    Dim lastPeople As Long, lastRights As Long, r As Long, a As Long, b As Long, pid As String, key As String
    Dim oldEvents As Boolean, errorNumber As Long, errorText As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = CC_Sheet
    a = CompactLeftYear(): b = CompactRightYear()
    lastPeople = CC_Last(2): lastRights = CC_Last(30)
    If lastPeople < 10 Then GoTo Done
    people = ws.Range("A10:I" & CStr(lastPeople)).Value2
    Set balances = CC_Dict()
    If lastRights >= 2 Then
        rights = ws.Range("AD2:AG" & CStr(lastRights)).Value2
        For r = 1 To UBound(rights, 1)
            If IsNumeric(rights(r, 2)) And IsNumeric(rights(r, 3)) And IsNumeric(rights(r, 4)) Then
                balances(CC_Key(CC_Text(rights(r, 1)), CLng(rights(r, 2)))) = CDbl(rights(r, 3)) - CDbl(rights(r, 4))
            End If
        Next r
    End If
    ReDim values(1 To UBound(people, 1), 1 To 3)
    For r = 1 To UBound(people, 1)
        pid = CC_Text(people(r, 9))
        If Len(pid) > 0 Then
            values(r, 1) = 0#: values(r, 2) = 0#
            key = CC_Key(pid, a): If balances.Exists(key) Then values(r, 1) = balances(key)
            key = CC_Key(pid, b): If balances.Exists(key) Then values(r, 2) = balances(key)
            values(r, 3) = CDbl(values(r, 1)) + CDbl(values(r, 2))
        End If
    Next r
    CompactAssignBlock ws, "E10:G" & CStr(lastPeople), values
Done:
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Application.EnableEvents = oldEvents
    Err.Raise errorNumber, "CompactRefreshRemaining", errorText
End Sub

Public Function CompactEditYears(ByVal target As Range) As String
    Dim ws As Worksheet, headers As Range, dataEdits As Range, area As Range, edits As Collection, item As Variant
    Dim a As Long, b As Long, oldA As Long, oldB As Long, lastRights As Long, lastLeaves As Long, newLast As Long
    Dim oldRights As Variant, oldAlloc As Variant, oldEvents As Boolean, changed As Boolean, result As String, failure As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = CC_Sheet
    If target Is Nothing Then GoTo Done
    If Not target.Parent Is ws Then result = "ERR:Yıl başlığını İzin Girişi sayfasında değiştirin.": GoTo Done
    Set headers = Application.Intersect(target, ws.Range("E9:F9"))
    If headers Is Nothing Then GoTo Done
    CompactInitYears
    oldA = CompactLeftYear(): oldB = CompactRightYear()
    If headers.Cells.CountLarge = 2 Then
        a = CY_ParseYear(ws.Range("E9").Value2): b = CY_ParseYear(ws.Range("F9").Value2)
    ElseIf Not Application.Intersect(headers, ws.Range("F9")) Is Nothing Then
        b = CY_ParseYear(ws.Range("F9").Value2): a = b - 1
    Else
        a = CY_ParseYear(ws.Range("E9").Value2): b = a + 1
    End If
    If a < 2000 Or b > 2100 Or b <> a + 1 Then
        result = "ERR:Ardışık iki yıl girin. Örnek: 2026 ve 2027."
        CompactRefreshYearHeaders
        CompactRefreshRemaining
        GoTo Done
    End If
    Set dataEdits = Application.Intersect(target, ws.Range("E10:F" & CStr(Application.Max(10, CC_Last(2)))))
    Set edits = New Collection
    If Not dataEdits Is Nothing Then
        For Each area In dataEdits.Areas
            edits.Add Array(area.address, area.Value2)
        Next area
    End If
    lastRights = Application.Max(1, CC_Last(30)): lastLeaves = Application.Max(2, CC_Last(18))
    oldRights = ws.Range("AD1:AG" & CStr(lastRights)).Value2
    oldAlloc = ws.Range("AA2:AA" & CStr(lastLeaves)).Value2
    ws.Range("AI8").Value2 = a: ws.Range("AI9").Value2 = b
    changed = True
    CompactRefreshYearHeaders
    result = CompactRecalcBalances()
    If Len(result) > 0 Then GoTo Reject
    If Not dataEdits Is Nothing Then
        For Each item In edits
            ws.Range(CStr(item(0))).Value2 = item(1)
        Next item
        result = CompactEditRemaining(dataEdits)
        If Len(result) > 0 Then GoTo Reject
    End If
    GoTo Done
Reject:
    newLast = CC_Last(30)
    ws.Range("AD1:AG" & CStr(lastRights)).Value2 = oldRights
    If newLast > lastRights Then ws.Range("AD" & CStr(lastRights + 1) & ":AG" & CStr(newLast)).ClearContents
    ws.Range("AA2:AA" & CStr(lastLeaves)).Value2 = oldAlloc
    ws.Range("AI8").Value2 = oldA: ws.Range("AI9").Value2 = oldB
    CompactRefreshYearHeaders
    CompactRefreshRemaining
Done:
    CompactEditYears = result
    Application.EnableEvents = oldEvents
    Exit Function
Failed:
    failure = Err.Description
    On Error Resume Next
    If changed Then
        newLast = CC_Last(30)
        ws.Range("AD1:AG" & CStr(lastRights)).Value2 = oldRights
        If newLast > lastRights Then ws.Range("AD" & CStr(lastRights + 1) & ":AG" & CStr(newLast)).ClearContents
        ws.Range("AA2:AA" & CStr(lastLeaves)).Value2 = oldAlloc
        ws.Range("AI8").Value2 = oldA: ws.Range("AI9").Value2 = oldB
    End If
    CompactRefreshYearHeaders
    CompactRefreshRemaining
    CompactEditYears = "ERR:Yıl değiştirilemedi: " & failure
    Application.EnableEvents = oldEvents
End Function

Public Function CompactRollYear(ByVal newRightYear As Long) As String
    Dim oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    CompactInitYears
    CC_Sheet.Range("F9").Value2 = newRightYear
    CompactRollYear = CompactEditYears(CC_Sheet.Range("F9"))
    Application.EnableEvents = oldEvents
    Exit Function
Failed:
    CompactRollYear = "ERR:Yıl değiştirilemedi: " & Err.Description
    Application.EnableEvents = oldEvents
End Function

Public Function CompactEditRemaining(ByVal target As Range) As String
    Dim ws As Worksheet, edits As Range, cell As Range, people As Variant, oldRights As Variant, oldAlloc As Variant, staged() As Variant, current As Variant, leaves As Variant
    Dim positions As Object, targets As Object, charges As Object, item As Variant, itemKey As Variant, key As String, pid As String
    Dim a As Long, b As Long, yr As Long, rowNumber As Long, lastPeople As Long, lastRights As Long, lastLeaves As Long
    Dim newLast As Long, r As Long, c As Long, baseline As Double, amount As Double, delta As Double
    Dim oldEvents As Boolean, wrote As Boolean, adjusted As Boolean, anyChange As Boolean, result As String, failure As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = CC_Sheet
    If target Is Nothing Then GoTo Done
    If Not target.Parent Is ws Then result = "ERR:İzin gününü İzin Girişi sayfasında girin.": GoTo Done
    lastPeople = CC_Last(2)
    If lastPeople < 10 Then GoTo Done
    Set edits = Application.Intersect(target, ws.Range("E10:F" & CStr(lastPeople)))
    If edits Is Nothing Then GoTo Done
    a = CompactLeftYear(): b = CompactRightYear()
    people = ws.Range("A10:I" & CStr(lastPeople)).Value2
    lastRights = Application.Max(1, CC_Last(30)): lastLeaves = Application.Max(2, CC_Last(18))
    oldRights = ws.Range("AD1:AG" & CStr(lastRights)).Value2
    oldAlloc = ws.Range("AA2:AA" & CStr(lastLeaves)).Value2
    Set positions = CC_Dict(): Set targets = CC_Dict()
    For r = 2 To UBound(oldRights, 1)
        If IsNumeric(oldRights(r, 2)) Then
            key = CC_Key(CC_Text(oldRights(r, 1)), CLng(oldRights(r, 2)))
            If positions.Exists(key) Then result = "ERR:Aynı personel/yıl için iki hak kaydı var.": GoTo Reject
            positions(key) = r
        Else
            result = "ERR:İzin hakkı yılını kontrol edin.": GoTo Reject
        End If
    Next r
    For Each cell In edits.Cells
        r = cell.row - 9
        pid = CC_Text(people(r, 9))
        If Len(pid) = 0 Or Len(CC_Text(people(r, 2))) = 0 Or Len(CC_Text(people(r, 3))) = 0 Then
            result = "ERR:" & cell.address(False, False) & " için önce geçerli bir personel ekleyin.": GoTo Reject
        End If
        For yr = a To b
            key = CC_Key(pid, yr)
            If Not targets.Exists(key) Then
                baseline = 0#
                If positions.Exists(key) Then
                    rowNumber = CLng(positions(key))
                    If Not IsNumeric(oldRights(rowNumber, 3)) Or Not IsNumeric(oldRights(rowNumber, 4)) Then result = "ERR:İzin hakkı gününü kontrol edin.": GoTo Reject
                    baseline = CDbl(oldRights(rowNumber, 3)) - CDbl(oldRights(rowNumber, 4))
                End If
                targets(key) = Array(pid, yr, baseline)
            End If
        Next yr
        If IsError(people(r, cell.column)) Then result = "ERR:" & cell.address(False, False) & " için sayısal gün girin.": GoTo Reject
        If IsEmpty(people(r, cell.column)) Or Not IsNumeric(people(r, cell.column)) Then result = "ERR:" & cell.address(False, False) & " boş bırakılamaz; izin yoksa 0 girin.": GoTo Reject
        amount = CDbl(people(r, cell.column))
        If amount < 0 Or Abs(amount * 2 - Round(amount * 2, 0)) > 0.000001 Then result = "ERR:" & cell.address(False, False) & " için sıfır veya pozitif tam/yarım gün girin.": GoTo Reject
        yr = a: If cell.column = 6 Then yr = b
        key = CC_Key(pid, yr): item = targets(key)
        If Abs(CDbl(item(2)) - amount) > 0.000001 Then anyChange = True
        targets(key) = Array(pid, yr, amount)
    Next cell
    If Not anyChange Then
        CompactRefreshRemaining
        GoTo Done
    End If
    Set charges = CC_Dict()
    leaves = ws.Range("R2:AB" & CStr(lastLeaves)).Value2
    For r = 1 To UBound(leaves, 1)
        If CC_Active(leaves(r, 8)) And CC_Bool(leaves(r, 7)) And CompactTypeCode(CC_Text(leaves(r, 3))) = "YILLIK" Then
            If IsNumeric(leaves(r, 6)) Then
                pid = CC_Text(leaves(r, 2))
                If Not charges.Exists(pid) Then charges(pid) = 0#
                charges(pid) = CDbl(charges(pid)) + CDbl(leaves(r, 6))
            End If
        End If
    Next r
    newLast = lastRights
    For Each itemKey In targets.keys
        If Not positions.Exists(itemKey) Then newLast = newLast + 1: positions(itemKey) = newLast
    Next itemKey
    ReDim staged(1 To newLast, 1 To 4)
    For r = 1 To lastRights
        For c = 1 To 4: staged(r, c) = oldRights(r, c): Next c
    Next r
    For Each itemKey In targets.keys
        item = targets(itemKey): r = CLng(positions(itemKey))
        amount = 0#: If charges.Exists(CStr(item(0))) Then amount = CDbl(charges(CStr(item(0))))
        staged(r, 1) = CStr(item(0)): staged(r, 2) = CLng(item(1))
        staged(r, 3) = CDbl(item(2)) + amount: staged(r, 4) = 0#
    Next itemKey
    wrote = True
    ws.Range("AD1:AG" & CStr(newLast)).Value2 = staged
    ' First pass covers every eligible historical charge from the displayed years.
    ' Then retain only the actual used amount plus the requested remaining balance.
    ' This gives an exact result in at most two passes, including half-day edits.
    result = CompactRecalcBalances()
    If Len(result) > 0 Then GoTo Reject
    current = ws.Range("AD1:AG" & CStr(newLast)).Value2
    adjusted = False
    For Each itemKey In targets.keys
        item = targets(itemKey): r = CLng(positions(itemKey))
        amount = CDbl(item(2)) + CDbl(current(r, 4))
        If Abs(CDbl(current(r, 3)) - amount) > 0.000001 Then current(r, 3) = amount: adjusted = True
    Next itemKey
    If adjusted Then
        ws.Range("AD1:AG" & CStr(newLast)).Value2 = current
        result = CompactRecalcBalances()
        If Len(result) > 0 Then GoTo Reject
    End If
    current = ws.Range("AD1:AG" & CStr(newLast)).Value2
    For Each itemKey In targets.keys
        item = targets(itemKey): r = CLng(positions(itemKey))
        delta = CDbl(item(2)) - (CDbl(current(r, 3)) - CDbl(current(r, 4)))
        If Abs(delta) > 0.000001 Then result = "ERR:Kalan günler mevcut izinlerle dengelenemedi.": GoTo Reject
    Next itemKey
    GoTo Done
Reject:
    If wrote Then
        ws.Range("AD1:AG" & CStr(lastRights)).Value2 = oldRights
        If newLast > lastRights Then ws.Range("AD" & CStr(lastRights + 1) & ":AG" & CStr(newLast)).ClearContents
        ws.Range("AA2:AA" & CStr(lastLeaves)).Value2 = oldAlloc
    End If
    CompactRefreshRemaining
Done:
    CompactEditRemaining = result
    Application.EnableEvents = oldEvents
    Exit Function
Failed:
    failure = Err.Description
    On Error Resume Next
    If wrote Then
        ws.Range("AD1:AG" & CStr(lastRights)).Value2 = oldRights
        If newLast > lastRights Then ws.Range("AD" & CStr(lastRights + 1) & ":AG" & CStr(newLast)).ClearContents
        ws.Range("AA2:AA" & CStr(lastLeaves)).Value2 = oldAlloc
    End If
    CompactRefreshRemaining
    CompactEditRemaining = "ERR:Kalan gün kaydedilemedi: " & failure
    Application.EnableEvents = oldEvents
End Function

' Personnel management for the two-sheet tracker. No worksheet rows are deleted.
Private Function CP_Cache(Optional ByVal replacement As Variant) As Variant
    Static savedRoster As Variant
    If Not IsMissing(replacement) Then savedRoster = replacement
    CP_Cache = savedRoster
End Function

Private Function CP_LastPerson() As Long
    Dim ws As Worksheet
    Set ws = CC_Sheet
    CP_LastPerson = Application.Max(9, ws.Cells(ws.Rows.count, 2).End(xlUp).row, _
        ws.Cells(ws.Rows.count, 3).End(xlUp).row, ws.Cells(ws.Rows.count, 9).End(xlUp).row)
End Function

Public Sub CompactPrimeRoster()
    Dim data As Variant, unused As Variant, lastRow As Long
    lastRow = CP_LastPerson()
    If lastRow >= 10 Then data = CC_Sheet.Range("A10:I" & CStr(lastRow)).Value2
    unused = CP_Cache(data)
End Sub

Private Function CP_PersonRow(ByVal pid As String) As Long
    Dim data As Variant, r As Long, lastRow As Long
    lastRow = CP_LastPerson()
    If lastRow < 10 Then Exit Function
    data = CC_Sheet.Range("A10:I" & CStr(lastRow)).Value2
    For r = 1 To UBound(data, 1)
        If CC_Text(data(r, 9)) = pid And Len(pid) > 0 Then CP_PersonRow = r + 9: Exit Function
    Next r
End Function

Private Function CP_SicilExists(ByVal sicil As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long, data As Variant
    Set ws = CC_Sheet
    lastRow = CP_LastPerson()
    If lastRow < 10 Then Exit Function
    data = ws.Range("A10:I" & CStr(lastRow)).Value2
    For r = 1 To UBound(data, 1)
        If StrComp(CC_Text(data(r, 2)), sicil, vbTextCompare) = 0 Then CP_SicilExists = True: Exit Function
    Next r
End Function

Private Function CP_DaysOK(ByVal amount As Variant) As Boolean
    If IsError(amount) Then Exit Function
    If IsEmpty(amount) Or Len(Trim$(CStr(amount))) = 0 Then CP_DaysOK = True: Exit Function
    If Not IsNumeric(amount) Then Exit Function
    CP_DaysOK = CDbl(amount) >= 0 And Abs(CDbl(amount) * 2 - Round(CDbl(amount) * 2, 0)) < 0.000001
End Function

Private Function CP_Days(ByVal amount As Variant) As Double
    If Len(CC_Text(amount)) > 0 Then CP_Days = CDbl(amount)
End Function

Private Function CP_NextPersonId() As String
    Dim ws As Worksheet, data As Variant, columns As Variant, column As Variant
    Dim r As Long, n As Long, value As String, lastRow As Long
    Set ws = CC_Sheet
    n = CLng(Val(CStr(ws.Range("AI10").Value2)))
    If n < 1 Then n = 1
    columns = Array(9, 19, 30)
    For Each column In columns
        lastRow = ws.Cells(ws.Rows.count, CLng(column)).End(xlUp).row
        If lastRow >= 2 Then
            data = ws.Range(ws.Cells(1, CLng(column)), ws.Cells(lastRow, CLng(column))).Value2
            For r = 2 To UBound(data, 1)
                value = CC_Text(data(r, 1))
                If left$(UCase$(value), 1) = "P" Then
                    If Val(Mid$(value, 2)) >= n Then n = CLng(Val(Mid$(value, 2))) + 1
                End If
            Next r
        End If
    Next column
    CP_NextPersonId = "P" & Format$(n, "000000")
    ws.Range("AI10").Value2 = n + 1
End Function

Private Function CP_LeftYear() As Long
    CP_LeftYear = CLng(Val(CStr(CC_Sheet.Range("AI8").Value2)))
    If CP_LeftYear < 2000 Then CP_LeftYear = CLng(Val(CStr(CC_Sheet.Range("E9").Value2)))
End Function

Private Function CP_RightYear() As Long
    CP_RightYear = CLng(Val(CStr(CC_Sheet.Range("AI9").Value2)))
    If CP_RightYear < 2000 Then CP_RightYear = CLng(Val(CStr(CC_Sheet.Range("F9").Value2)))
End Function

Private Sub CP_RegisterRow(ByVal rosterRow As Long, ByVal sicil As String, ByVal personName As String, _
        ByVal department As String, ByVal leftDays As Double, ByVal rightDays As Double)
    Dim ws As Worksheet, values(1 To 1, 1 To 9) As Variant, rights(1 To 2, 1 To 4) As Variant
    Dim pid As String, lastRight As Long, templateRow As Long
    Set ws = CC_Sheet
    pid = CP_NextPersonId()
    templateRow = 10
    If rosterRow <> templateRow Then ws.Range("A" & templateRow & ":I" & templateRow).Copy ws.Range("A" & rosterRow & ":I" & rosterRow)
    Application.CutCopyMode = False
    values(1, 2) = sicil: values(1, 3) = personName: values(1, 4) = department
    values(1, 5) = leftDays: values(1, 6) = rightDays: values(1, 7) = leftDays + rightDays
    values(1, 8) = "+": values(1, 9) = pid
    ws.Range("A" & rosterRow & ":I" & rosterRow).Value2 = values
    ws.Cells(rosterRow, 2).NumberFormat = "@"
    ws.Cells(rosterRow, 2).Value2 = sicil
    ws.Range("B" & rosterRow & ":F" & rosterRow).Locked = False
    ws.Cells(rosterRow, 9).Locked = True
    ws.Rows(rosterRow).RowHeight = ws.Rows(templateRow).RowHeight
    lastRight = Application.Max(1, ws.Cells(ws.Rows.count, 30).End(xlUp).row)
    rights(1, 1) = pid: rights(1, 2) = CP_LeftYear(): rights(1, 3) = leftDays: rights(1, 4) = 0#
    rights(2, 1) = pid: rights(2, 2) = CP_RightYear(): rights(2, 3) = rightDays: rights(2, 4) = 0#
    ws.Range("AD" & CStr(lastRight + 1) & ":AG" & CStr(lastRight + 2)).Value2 = rights
End Sub

Public Sub CompactRepairRoster(Optional ByVal refreshAttendance As Boolean = False)
    Dim ws As Worksheet, r As Long, n As Long, lastRow As Long, sequence() As Variant
    Dim oldEvents As Boolean, problem As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = CC_Sheet
    If ws.FilterMode Then ws.ShowAllData
    lastRow = CP_LastPerson()
    If lastRow >= 10 Then
        ReDim sequence(1 To lastRow - 9, 1 To 1)
        For r = 10 To lastRow
            If Len(CC_Text(ws.Cells(r, 2).Value2)) > 0 And Len(CC_Text(ws.Cells(r, 3).Value2)) > 0 And Len(CC_Text(ws.Cells(r, 9).Value2)) > 0 Then
                n = n + 1: sequence(r - 9, 1) = n
                ws.Cells(r, 8).Value2 = "+"
                If ws.Cells(r, 8).Hyperlinks.count = 0 Then
                    ws.Hyperlinks.Add Anchor:=ws.Cells(r, 8), address:="", SubAddress:="'İzin Girişi'!J9", TextToDisplay:="+"
                End If
            Else
                ws.Cells(r, 8).ClearContents
            End If
        Next r
        ws.Range("A10:A" & CStr(lastRow)).Value2 = sequence
    End If
    ws.AutoFilterMode = False
    CU_Search False
    CompactPrimeRoster
    CompactReportLists
    CompactReportsDirty
    If refreshAttendance Then CompactRefreshAttendance
Done:
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    problem = Err.Description
    Application.EnableEvents = oldEvents
    Err.Raise vbObjectError + 247, "CompactRepairRoster", problem
End Sub

Private Sub CP_RestoreRoster(ByVal data As Variant)
    Dim ws As Worksheet, lastRow As Long, oldLast As Long
    Set ws = CC_Sheet
    lastRow = CP_LastPerson()
    oldLast = 9
    If IsArray(data) Then oldLast = UBound(data, 1) + 9
    CompactAssignBlock ws, "A10:I" & CStr(Application.Max(10, oldLast)), data, "A10:I" & CStr(Application.Max(10, lastRow, oldLast))
End Sub

Public Function CompactAddPerson(ByVal sicil As String, ByVal personName As String, ByVal department As String, _
        Optional ByVal leftRemaining As Double = 0, Optional ByVal rightRemaining As Double = 0) As String
    Dim ws As Worksheet, oldRoster As Variant, oldRights As Variant, oldEvents As Boolean
    Dim targetRow As Long, rightsLast As Long, newRightsLast As Long, wrote As Boolean, result As String, pid As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = CC_Sheet
    sicil = Trim$(sicil): personName = Trim$(personName): department = Trim$(department)
    If Len(sicil) = 0 Or Len(personName) = 0 Then CompactAddPerson = "ERR:Sicil ve ad soyad zorunludur.": GoTo Done
    If CP_SicilExists(sicil) Then CompactAddPerson = "ERR:Bu sicil numarası zaten listede var.": GoTo Done
    If Not CP_DaysOK(leftRemaining) Or Not CP_DaysOK(rightRemaining) Then CompactAddPerson = "ERR:İzin gününü sıfır veya pozitif tam/yarım gün girin.": GoTo Done
    If CP_LeftYear() < 2000 Or CP_RightYear() <= CP_LeftYear() Then CompactAddPerson = "ERR:Önce iki izin yılı başlığını kontrol edin.": GoTo Done
    targetRow = Application.Max(10, CP_LastPerson() + 1)
    If targetRow > 10 Then oldRoster = ws.Range("A10:I" & CStr(targetRow - 1)).Value2
    rightsLast = Application.Max(1, ws.Cells(ws.Rows.count, 30).End(xlUp).row)
    oldRights = ws.Range("AD1:AG" & CStr(rightsLast)).Value2
    wrote = True
    CP_RegisterRow targetRow, sicil, personName, department, leftRemaining, rightRemaining
    pid = CStr(ws.Cells(targetRow, 9).Value2)
    CU_Clear ws.Range("C5")
    CompactRepairRoster False
    CompactRefreshAttendance
    CompactAddPerson = pid
Done:
    Application.EnableEvents = oldEvents
    Exit Function
Failed:
    result = Err.Description
    On Error Resume Next
    If wrote Then
        CP_RestoreRoster oldRoster
        newRightsLast = ws.Cells(ws.Rows.count, 30).End(xlUp).row
        ws.Range("AD1:AG" & CStr(Application.Max(rightsLast, newRightsLast))).ClearContents
        ws.Range("AD1:AG" & CStr(rightsLast)).Value2 = oldRights
        CompactRepairRoster False
        CompactRefreshAttendance
    End If
    CompactAddPerson = "ERR:Personel eklenemedi: " & result
    Application.EnableEvents = oldEvents
End Function

Public Function CompactRemovePerson(ByVal pid As String) As String
    Dim ws As Worksheet, r As Long, oldRoster As Variant, oldEvents As Boolean, result As String, wrote As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = CC_Sheet
    r = CP_PersonRow(pid)
    If r = 0 Then CompactRemovePerson = "ERR:Listeden çıkarılacak personel bulunamadı.": GoTo Done
    oldRoster = ws.Range("A10:I" & CStr(CP_LastPerson())).Value2
    If ws.FilterMode Then ws.ShowAllData
    wrote = True
    ws.Range("A" & CStr(r) & ":I" & CStr(r)).Delete Shift:=xlUp
    If CStr(ws.Range("AI2").Value2) = pid Then
        ws.Range("AI2").ClearContents
        CompactClear
        CU_ShowSelected
    End If
    CU_Clear ws.Range("C5")
    CompactRepairRoster False
    CompactRefreshAttendance
Done:
    Application.EnableEvents = oldEvents
    Exit Function
Failed:
    result = Err.Description
    On Error Resume Next
    If wrote Then
        CP_RestoreRoster oldRoster
        CompactRepairRoster False
        CompactRefreshAttendance
    End If
    CompactRemovePerson = "ERR:Personel listeden çıkarılamadı: " & result
    Application.EnableEvents = oldEvents
End Function

Public Function CompactSyncRoster(ByVal target As Range) As String
    Dim ws As Worksheet, data As Variant, oldRoster As Variant, oldRights As Variant
    Dim names As Object, ids As Object, r As Long, lastRow As Long, rightLast As Long, currentRightLast As Long
    Dim sicil As String, nm As String, dept As String, pid As String, result As String, problem As String
    Dim oldEvents As Boolean, wrote As Boolean, rightsSnapshot As Boolean, registeredNew As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = CC_Sheet
    If target Is Nothing Then GoTo Done
    If Not target.Parent Is ws Then GoTo Done
    If Intersect(target, ws.Range("B10:D1048576")) Is Nothing Then GoTo Done
    oldRoster = CP_Cache()
    lastRow = CP_LastPerson()
    If lastRow < 10 Then CompactRepairRoster True: GoTo Done
    data = ws.Range("A10:I" & CStr(lastRow)).Value2
    If Not IsArray(oldRoster) Then oldRoster = data
    Set names = CC_Dict(): Set ids = CC_Dict()
    For r = 1 To UBound(data, 1)
        sicil = CC_Text(data(r, 2)): nm = CC_Text(data(r, 3)): dept = CC_Text(data(r, 4)): pid = CC_Text(data(r, 9))
        If Len(pid) > 0 And (Len(sicil) = 0 Or Len(nm) = 0) Then
            If Len(sicil) > 0 Or Len(nm) > 0 Or Len(dept) > 0 Then problem = "Kayıtlı personelin sicili ve adı boş bırakılamaz. Silmek için Personel Sil düğmesini kullanın.": GoTo Rejected
        End If
        If Len(sicil) > 0 Then
            If names.Exists(sicil) Then problem = "Aynı sicil numarası iki personele verilemez: " & sicil: GoTo Rejected
            names(sicil) = True
        End If
        If Len(pid) > 0 Then
            If ids.Exists(pid) Then problem = "Personel satırlarını gizli kayıt kimliğiyle kopyalamayın; Personel Ekle düğmesini kullanın.": GoTo Rejected
            ids(pid) = True
        ElseIf Len(sicil) > 0 And Len(nm) > 0 Then
            If Not CP_DaysOK(data(r, 5)) Or Not CP_DaysOK(data(r, 6)) Then problem = "Yeni personelin izin günlerini sıfır veya pozitif tam/yarım gün girin.": GoTo Rejected
        End If
    Next r
    rightLast = Application.Max(1, ws.Cells(ws.Rows.count, 30).End(xlUp).row)
    oldRights = ws.Range("AD1:AG" & CStr(rightLast)).Value2
    rightsSnapshot = True
    wrote = True
    For r = 1 To UBound(data, 1)
        If Len(CC_Text(data(r, 9))) = 0 And Len(CC_Text(data(r, 2))) > 0 And Len(CC_Text(data(r, 3))) > 0 Then
            CP_RegisterRow r + 9, CC_Text(data(r, 2)), CC_Text(data(r, 3)), CC_Text(data(r, 4)), CP_Days(data(r, 5)), CP_Days(data(r, 6))
            registeredNew = True
        End If
    Next r
    For r = UBound(data, 1) To 1 Step -1
        If Len(CC_Text(data(r, 9))) > 0 And Len(CC_Text(data(r, 2))) = 0 And Len(CC_Text(data(r, 3))) = 0 And Len(CC_Text(data(r, 4))) = 0 Then
            pid = CC_Text(data(r, 9))
            If ws.FilterMode Then ws.ShowAllData
            ws.Range("A" & CStr(r + 9) & ":I" & CStr(r + 9)).Delete Shift:=xlUp
            If CStr(ws.Range("AI2").Value2) = pid Then ws.Range("AI2").ClearContents: CompactClear: CU_ShowSelected
        End If
    Next r
    If registeredNew Then CU_Clear ws.Range("C5")
    CompactRepairRoster False
    If Len(CStr(ws.Range("AI2").Value2)) > 0 Then CU_ShowSelected
    CompactRefreshAttendance
    GoTo Done
Rejected:
    result = "ERR:" & problem
    CP_RestoreRoster oldRoster
    problem = CompactRecalcBalances()
    CompactRepairRoster False
    CompactRefreshAttendance
    GoTo Done
Failed:
    result = "ERR:Personel bilgisi güncellenemedi: " & Err.Description
    On Error Resume Next
    If IsArray(oldRoster) Then CP_RestoreRoster oldRoster
    If wrote And rightsSnapshot Then
        currentRightLast = ws.Cells(ws.Rows.count, 30).End(xlUp).row
        ws.Range("AD1:AG" & CStr(Application.Max(rightLast, currentRightLast))).ClearContents
        ws.Range("AD1:AG" & CStr(rightLast)).Value2 = oldRights
    End If
    problem = CompactRecalcBalances()
    CompactRepairRoster False
    CompactRefreshAttendance
Done:
    CompactSyncRoster = result
    Application.EnableEvents = oldEvents
End Function

Public Sub CompactPersonAddUI()
    Dim sicil As Variant, nm As Variant, dept As Variant, result As String, rowNumber As Long
    sicil = Application.InputBox("Yeni personelin sicil numarası:", "Personel Ekle", Type:=2)
    If VarType(sicil) = vbBoolean Then Exit Sub
    If Len(Trim$(CStr(sicil))) = 0 Then Exit Sub
    nm = Application.InputBox("Adı soyadı:", "Personel Ekle", Type:=2)
    If VarType(nm) = vbBoolean Then Exit Sub
    If Len(Trim$(CStr(nm))) = 0 Then Exit Sub
    dept = Application.InputBox("Daire başkanlığı (isteğe bağlı):", "Personel Ekle", Type:=2)
    If VarType(dept) = vbBoolean Then Exit Sub
    result = CompactAddPerson(CStr(sicil), CStr(nm), CStr(dept))
    If left$(result, 3) = "ERR" Then MsgBox result, vbExclamation: Exit Sub
    rowNumber = CP_PersonRow(result)
    CompactSelectPerson rowNumber
    CC_Sheet.Cells(rowNumber, 5).Select
    ActiveWindow.ScrollRow = Application.Max(1, rowNumber - 4)
    MsgBox "Personel eklendi. İki yılın izin günlerini bu satırdaki gün alanlarına girebilirsiniz.", vbInformation
End Sub

Public Sub CompactPersonDeleteUI()
    Dim ws As Worksheet, pid As String, r As Long, result As String, label As String
    Set ws = CC_Sheet
    pid = CStr(ws.Range("AI2").Value2)
    If ActiveSheet.name = ws.name Then
        If ActiveCell.row >= 10 And ActiveCell.row <= CP_LastPerson() And ActiveCell.column <= 9 Then pid = CStr(ws.Cells(ActiveCell.row, 9).Value2)
    End If
    r = CP_PersonRow(pid)
    If r = 0 Then MsgBox "Önce listeden silinecek personelin satırını seçin.", vbExclamation: Exit Sub
    label = CStr(ws.Cells(r, 2).Value2) & " - " & CStr(ws.Cells(r, 3).Value2)
    If MsgBox(label & vbCrLf & "Bu personel izin ve puantaj listesinden çıkarılsın mı?", vbYesNo + vbQuestion + vbDefaultButton2, "Personel Sil") <> vbYes Then Exit Sub
    result = CompactRemovePerson(pid)
    If left$(result, 3) = "ERR" Then MsgBox result, vbExclamation: Exit Sub
    MsgBox "Personel izin ve puantaj listesinden çıkarıldı.", vbInformation
End Sub

' Physical leave deletion. Only the hidden R:AB ledger is compacted.
' Imported history is already included in the opening balances and is never refunded twice.
Public Sub CompactPrimeLeaveIds()
    Dim ws As Worksheet, data As Variant, lastRow As Long, r As Long, n As Long, suffix As Double, text As String
    Set ws = CC_Sheet
    n = CLng(Val(CStr(ws.Range("AI4").Value2)))
    If n < 1 Then n = 1
    lastRow = CC_Last(18)
    If lastRow >= 2 Then
        data = ws.Range("R2:R" & CStr(lastRow)).Value2
        If IsArray(data) Then
            For r = 1 To UBound(data, 1)
                text = CC_Text(data(r, 1))
                If Len(text) > 1 Then
                    suffix = Val(Mid$(text, 2))
                    If suffix >= n And suffix < 2147483646# Then n = CLng(suffix) + 1
                End If
            Next r
        Else
            text = CC_Text(data)
            If Len(text) > 1 Then
                suffix = Val(Mid$(text, 2))
                If suffix >= n And suffix < 2147483646# Then n = CLng(suffix) + 1
            End If
        End If
    End If
    ws.Range("AI4").Value2 = n
End Sub

Public Function CompactNextLeaveId() As String
    Dim ws As Worksheet, n As Long
    Set ws = CC_Sheet
    CompactPrimeLeaveIds
    n = CLng(ws.Range("AI4").Value2)
    If n >= 2147483646 Then Err.Raise vbObjectError + 281, "CompactNextLeaveId", "Yeni izin numarası üretilemedi."
    CompactNextLeaveId = "N" & Format$(n, "000000")
    ws.Range("AI4").Value2 = n + 1
End Function

Public Function CompactDeleteLeave(ByVal id As String) As String
    Dim result As Variant
    If Len(Trim$(id)) = 0 Then CompactDeleteLeave = "ERR:Silinecek izin kaydını seçin.": Exit Function
    result = CD_DeleteRows(Trim$(id), False)
    If VarType(result) = vbString Then
        CompactDeleteLeave = CStr(result)
    ElseIf CLng(result) = 0 Then
        CompactDeleteLeave = "ERR:Silinecek izin kaydı bulunamadı."
    End If
End Function

Public Function CompactPurgeCancelled() As Variant
    CompactPurgeCancelled = CD_DeleteRows("", True)
End Function

Private Function CD_DeleteRows(ByVal id As String, ByVal cancelledOnly As Boolean) As Variant
    Dim ws As Worksheet, oldLedger As Variant, oldRights As Variant, oldRoster As Variant
    Dim keep() As Variant, lastLeave As Long, lastRight As Long, lastPerson As Long
    Dim rowCount As Long, removed As Long, kept As Long, r As Long, c As Long, remove As Boolean
    Dim result As String, failure As String, wrote As Boolean, snapshotReady As Boolean, oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = CC_Sheet
    lastLeave = CC_Last(18)
    If lastLeave < 2 Then CD_DeleteRows = 0: GoTo Done
    oldLedger = ws.Range("R2:AB" & CStr(lastLeave)).Value2
    rowCount = UBound(oldLedger, 1)
    For r = 1 To rowCount
        If cancelledOnly Then
            remove = Len(CC_Text(oldLedger(r, 1))) > 0 And Not CC_Active(oldLedger(r, 8))
        Else
            remove = CC_Text(oldLedger(r, 1)) = id
        End If
        If remove Then removed = removed + 1
    Next r
    If removed = 0 Then CD_DeleteRows = 0: GoTo Done
    lastRight = Application.Max(1, CC_Last(30))
    lastPerson = CC_Last(2)
    oldRights = ws.Range("AD1:AG" & CStr(lastRight)).Value2
    If lastPerson >= 10 Then oldRoster = ws.Range("E10:G" & CStr(lastPerson)).Value2
    snapshotReady = True
    ' Reserve IDs before deleted/cancelled rows disappear from the ledger.
    CompactPrimeLeaveIds
    If removed < rowCount Then
        ReDim keep(1 To rowCount - removed, 1 To 11)
        For r = 1 To rowCount
            If cancelledOnly Then
                remove = Len(CC_Text(oldLedger(r, 1))) > 0 And Not CC_Active(oldLedger(r, 8))
            Else
                remove = CC_Text(oldLedger(r, 1)) = id
            End If
            If Not remove Then
                kept = kept + 1
                For c = 1 To 11: keep(kept, c) = oldLedger(r, c): Next c
            End If
        Next r
    End If
    wrote = True
    ws.Range("R2:AB" & CStr(lastLeave)).ClearContents
    If kept > 0 Then ws.Range("R2:AB" & CStr(kept + 1)).Value2 = keep
    result = CompactRecalcBalances()
    If Len(result) > 0 Then GoTo Reject
    CompactReportsDirty
    CD_DeleteRows = removed
    GoTo Done
Reject:
    ws.Range("R2:AB" & CStr(lastLeave)).ClearContents
    ws.Range("R2:AB" & CStr(lastLeave)).Value2 = oldLedger
    ws.Range("AD1:AG" & CStr(lastRight)).Value2 = oldRights
    If lastPerson >= 10 Then CompactAssignBlock ws, "E10:G" & CStr(lastPerson), oldRoster
    CD_DeleteRows = result
Done:
    Application.EnableEvents = oldEvents
    Exit Function
Failed:
    failure = Err.Description
    On Error Resume Next
    If wrote And snapshotReady Then
        ws.Range("R2:AB" & CStr(lastLeave)).ClearContents
        ws.Range("R2:AB" & CStr(lastLeave)).Value2 = oldLedger
        ws.Range("AD1:AG" & CStr(lastRight)).Value2 = oldRights
        If lastPerson >= 10 Then CompactAssignBlock ws, "E10:G" & CStr(lastPerson), oldRoster
    End If
    CD_DeleteRows = "ERR:İzin kaydı silinemedi: " & failure
    Application.EnableEvents = oldEvents
End Function

' Reports use bounded arrays and run only when requested.
Private Function CR_Sheet() As Worksheet
    Set CR_Sheet = ThisWorkbook.Worksheets("Raporlar")
End Function

Private Function CR_Date(ByVal value As Variant) As Long
    On Error GoTo Done
    If IsError(value) Or IsEmpty(value) Then Exit Function
    If IsNumeric(value) Then
        If CC_DateOK(value) Then CR_Date = CLng(Fix(CDbl(value)))
    ElseIf IsDate(value) Then
        If Year(CDate(value)) >= 2000 And Year(CDate(value)) <= 2100 Then CR_Date = CLng(DateValue(CDate(value)))
    End If
Done:
End Function

Private Function CR_All(ByVal value As String) As Boolean
    value = CC_Normal(value)
    CR_All = Len(value) = 0 Or value = "tumu" Or value = "tum personel" Or value = "tum birimler" Or value = "tum izinler"
End Function

Private Function CR_PersonMatch(ByVal sicil As String, ByVal personName As String, ByVal query As String) As Boolean
    Dim term As String
    If CR_All(query) Then CR_PersonMatch = True: Exit Function
    term = CC_Normal(query)
    If IsNumeric(term) Then
        CR_PersonMatch = CC_Normal(sicil) = term
    Else
        CR_PersonMatch = InStr(1, CC_Normal(sicil & " - " & personName), term, vbBinaryCompare) > 0
    End If
End Function

Private Function CR_Holidays() As Object
    Dim values As Variant, result As Object, r As Long, serial As Long
    Set result = CC_Dict()
    values = ThisWorkbook.Worksheets("Puantaj").Range("K10:K109").Value2
    For r = 1 To UBound(values, 1)
        serial = CR_Date(values(r, 1))
        If serial > 0 Then result(CStr(serial)) = True
    Next r
    Set CR_Holidays = result
End Function

Private Function CR_WorkDays(ByVal firstDay As Long, ByVal lastDay As Long, ByVal holidays As Object) As Long
    Dim d As Long
    For d = firstDay To lastDay
        If Weekday(CDate(d), vbMonday) <= 5 And Not holidays.Exists(CStr(d)) Then CR_WorkDays = CR_WorkDays + 1
    Next d
End Function

Private Function CR_Number(ByVal value As Variant) As Double
    If Not IsError(value) And IsNumeric(value) Then CR_Number = CDbl(value)
End Function

Private Function CR_SortKey(ByVal sicil As String, ByVal firstDay As Long, ByVal id As String) As String
    If IsNumeric(sicil) Then
        CR_SortKey = "0" & Right$(String$(20, "0") & CStr(Val(sicil)), 20)
    Else
        CR_SortKey = "1" & CC_Normal(sicil)
    End If
    CR_SortKey = CR_SortKey & "|" & Format$(firstDay, "000000") & "|" & id
End Function

Private Sub CR_Sort(ByRef indices() As Long, ByRef keys() As String, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As String, value As Long
    i = low: j = high: pivot = keys(indices((low + high) \ 2))
    Do While i <= j
        Do While keys(indices(i)) < pivot: i = i + 1: Loop
        Do While keys(indices(j)) > pivot: j = j - 1: Loop
        If i <= j Then
            value = indices(i): indices(i) = indices(j): indices(j) = value
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then CR_Sort indices, keys, low, j
    If i < high Then CR_Sort indices, keys, i, high
End Sub

Public Sub CompactReportsDirty()
    Dim ws As Worksheet, oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Done
    Application.EnableEvents = False
    Set ws = CR_Sheet
    ws.Range("N5").Value2 = True
    ws.Range("A14").Value2 = "Veriler veya filtreler değişti. Güncel sonucu görmek için RAPORU GETİR düğmesine basın."
Done:
    Application.EnableEvents = oldEvents
End Sub

Public Sub CompactReportPeriod()
    Dim ws As Worksheet, yr As Variant, mn As Variant, oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Done
    Application.EnableEvents = False
    Set ws = CR_Sheet
    yr = ws.Range("B11").Value2: mn = ws.Range("F11").Value2
    If Not IsNumeric(yr) Or Not IsNumeric(mn) Then GoTo Done
    If CDbl(yr) <> Fix(CDbl(yr)) Or CDbl(mn) <> Fix(CDbl(mn)) Then GoTo Done
    If CLng(yr) < 2000 Or CLng(yr) > 2100 Or CLng(mn) < 0 Or CLng(mn) > 12 Then GoTo Done
    If CLng(mn) = 0 Then
        ws.Range("B9").Value2 = CDbl(DateSerial(CLng(yr), 1, 1))
        ws.Range("F9").Value2 = CDbl(DateSerial(CLng(yr), 12, 31))
    Else
        ws.Range("B9").Value2 = CDbl(DateSerial(CLng(yr), CLng(mn), 1))
        ws.Range("F9").Value2 = CDbl(DateSerial(CLng(yr), CLng(mn) + 1, 0))
    End If
    CompactReportsDirty
Done:
    Application.EnableEvents = oldEvents
End Sub

Public Function CompactReportRefresh() As String
    Dim ws As Worksheet, source As Worksheet, people As Variant, leaves As Variant
    Dim active As Object, selected As Object, counts As Object, holidays As Object, allDates As Object, categoryDates As Object
    Dim headers(1 To 1, 1 To 12) As Variant, output() As Variant, indexes() As Long, keys() As String
    Dim mode As String, queryDepartment As String, queryPerson As String, queryType As String, filterType As String
    Dim firstDay As Long, lastDay As Long, lastPeople As Long, lastLeaves As Long, oldLast As Long, lastOut As Long, lastCol As Long
    Dim r As Long, p As Long, i As Long, c As Long, d As Long, startDay As Long, endDay As Long, a As Long, b As Long, count As Long, category As Long
    Dim pid As String, code As String, key As String, categoryKey As String, note As String, failure As String
    Dim values As Variant, oldEvents As Boolean, oldScreen As Boolean, leftYear As Long, rightYear As Long, wasProtected As Boolean, stage As String, band As String, bandPart As String
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set ws = CR_Sheet: Set source = CC_Sheet
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect "2022-tiski"
    mode = CC_Normal(CC_Text(ws.Range("B5").Value2))
    Select Case mode
        Case "kullanilan izinler": mode = "detail"
        Case "kalan izinler": mode = "remaining"
        Case "personel ozeti": mode = "summary"
        Case Else: failure = "Rapor türünü seçin.": GoTo Invalid
    End Select
    queryDepartment = CC_Text(ws.Range("F5").Value2)
    queryPerson = CC_Text(ws.Range("B7").Value2)
    queryType = CC_Text(ws.Range("F7").Value2)
    If Not CR_All(queryType) Then
        filterType = CompactTypeCode(queryType)
        If Len(filterType) = 0 Then failure = "Geçerli bir izin türü veya Tümü seçin.": GoTo Invalid
    End If
    firstDay = CR_Date(ws.Range("B9").Value2): lastDay = CR_Date(ws.Range("F9").Value2)
    If mode <> "remaining" Then
        If firstDay = 0 Or lastDay = 0 Or lastDay < firstDay Then failure = "Başlangıç ve bitiş tarihlerini kontrol edin.": GoTo Invalid
    End If
    leftYear = CompactLeftYear(): rightYear = CompactRightYear()
    Set active = CC_Dict(): Set selected = CC_Dict(): Set counts = CC_Dict()
    lastPeople = CC_Last(2)
    If lastPeople >= 10 Then
        people = source.Range("A10:I" & CStr(lastPeople)).Value2
        For r = 1 To UBound(people, 1)
            pid = CC_Text(people(r, 9))
            If Len(pid) > 0 And Len(CC_Text(people(r, 2))) > 0 And Len(CC_Text(people(r, 3))) > 0 Then
                If active.Exists(pid) Then failure = "Aynı personel bağlantısı birden çok satırda bulunuyor.": GoTo Invalid
                active(pid) = r
                If CR_All(queryDepartment) Or InStr(1, CC_Normal(CC_Text(people(r, 4))), CC_Normal(queryDepartment), vbBinaryCompare) > 0 Then
                    If CR_PersonMatch(CC_Text(people(r, 2)), CC_Text(people(r, 3)), queryPerson) Then
                        selected(pid) = r
                        counts(pid) = Array(0&, 0&, 0&, 0&, 0&)
                    End If
                End If
            End If
        Next r
    End If
    headers(1, 1) = "Sıra": headers(1, 2) = "Sicil": headers(1, 3) = "Adı Soyadı": headers(1, 4) = "Daire Başkanlığı"
    If mode = "remaining" Then
        headers(1, 5) = CStr(leftYear) & " Kalan": headers(1, 6) = CStr(rightYear) & " Kalan": headers(1, 7) = "Toplam Kalan"
        lastCol = 7: count = selected.count
        If count > 0 Then
            ReDim output(1 To count, 1 To 12)
            i = 0
            For r = 1 To UBound(people, 1)
                pid = CC_Text(people(r, 9))
                If selected.Exists(pid) Then
                    i = i + 1
                    output(i, 1) = i
                    For c = 2 To 4: output(i, c) = people(r, c): Next c
                    output(i, 5) = CR_Number(people(r, 5)): output(i, 6) = CR_Number(people(r, 6))
                    output(i, 7) = CDbl(output(i, 5)) + CDbl(output(i, 6))
                End If
            Next r
        End If
        note = CStr(count) & " personel - Güncel yıllık izin bakiyeleri. Tarih ve izin türü filtresi kalan günleri değiştirmez."
        ws.Range("A13").Value2 = "Kalan günler güncel bakiyedir. Bu rapor geçmiş bir tarih itibarıyla bakiye hesaplamaz. İki yıl, İzin Girişi sayfasındaki başlıklardır."
    Else
        lastLeaves = CC_Last(18)
        Set holidays = CR_Holidays(): Set allDates = CC_Dict(): Set categoryDates = CC_Dict()
        ReDim indexes(1 To Application.Max(1, lastLeaves - 1))
        ReDim keys(1 To Application.Max(1, lastLeaves - 1))
        If lastLeaves >= 2 And selected.count > 0 Then
            leaves = source.Range("R2:AB" & CStr(lastLeaves)).Value2
            For r = 1 To UBound(leaves, 1)
                pid = CC_Text(leaves(r, 2))
                If selected.Exists(pid) And CC_Active(leaves(r, 8)) Then
                    code = CompactTypeCode(CC_Text(leaves(r, 3)))
                    If Len(filterType) = 0 Or code = filterType Then
                        startDay = CR_Date(leaves(r, 4)): endDay = CR_Date(leaves(r, 5))
                        If startDay = 0 Or endDay = 0 Or endDay < startDay Then failure = "İzin kaydının tarihlerini kontrol edin: " & CC_Text(leaves(r, 1)): GoTo Invalid
                        If startDay <= lastDay And endDay >= firstDay Then
                            a = startDay: If a < firstDay Then a = firstDay
                            b = endDay: If b > lastDay Then b = lastDay
                            If mode = "detail" Then
                                count = count + 1: indexes(count) = r
                                p = CLng(selected(pid))
                                keys(r) = CR_SortKey(CC_Text(people(p, 2)), startDay, CC_Text(leaves(r, 1)))
                            Else
                                values = counts(pid): category = 2
                                If code = "YILLIK" Then category = 0
                                If code = "RAPOR" Then category = 1
                                For d = a To b
                                    key = pid & "|" & CStr(d)
                                    categoryKey = key & "|" & CStr(category)
                                    If Not categoryDates.Exists(categoryKey) Then categoryDates(categoryKey) = True: values(category) = CLng(values(category)) + 1
                                    If Not allDates.Exists(key) Then
                                        allDates(key) = True: values(3) = CLng(values(3)) + 1
                                        If Weekday(CDate(d), vbMonday) <= 5 And Not holidays.Exists(CStr(d)) Then values(4) = CLng(values(4)) + 1
                                    End If
                                Next d
                                counts(pid) = values
                            End If
                        End If
                    End If
                End If
            Next r
        End If
        lastCol = 12
        If mode = "detail" Then
            headers(1, 5) = "İzin Türü": headers(1, 6) = "Başlangıç": headers(1, 7) = "Bitiş"
            headers(1, 8) = "Kayıttaki Gün": headers(1, 9) = "Dönemde Takvim Günü": headers(1, 10) = "Dönemde Mesai Günü"
            headers(1, 11) = "Kayıt No": headers(1, 12) = "Açıklama"
            If count > 0 Then
                If count > 1 Then CR_Sort indexes, keys, 1, count
                ReDim output(1 To count, 1 To 12)
                For i = 1 To count
                    r = indexes(i): pid = CC_Text(leaves(r, 2)): p = CLng(selected(pid))
                    startDay = CR_Date(leaves(r, 4)): endDay = CR_Date(leaves(r, 5))
                    a = startDay: If a < firstDay Then a = firstDay
                    b = endDay: If b > lastDay Then b = lastDay
                    output(i, 1) = i
                    For c = 2 To 4: output(i, c) = people(p, c): Next c
                    output(i, 5) = CompactTypeLabel(CC_Text(leaves(r, 3)))
                    output(i, 6) = startDay: output(i, 7) = endDay
                    output(i, 8) = CR_Number(leaves(r, 6)): output(i, 9) = b - a + 1
                    output(i, 10) = CR_WorkDays(a, b, holidays)
                    output(i, 11) = CC_Text(leaves(r, 1)): output(i, 12) = CC_Text(leaves(r, 9))
                Next i
            End If
            note = CStr(count) & " izin / rapor kaydı. Kayıttaki Gün tam kayda aittir; dönem sütunları yalnız seçili aralığın günlerini gösterir."
            ws.Range("A13").Value2 = "Kayıttaki Gün elle girilen tam kayıt günüdür. Dönemde Takvim/Mesai Günleri tarih kesişiminden bulunur; elle girilen gün aylara tahmini dağıtılmaz."
        Else
            headers(1, 5) = "Yıllık Takvim Günü": headers(1, 6) = "Rapor Takvim Günü": headers(1, 7) = "Diğer Takvim Günü"
            headers(1, 8) = "Toplam Takvim Günü": headers(1, 9) = "Puantajdan Düşen Mesai"
            headers(1, 10) = CStr(leftYear) & " Kalan": headers(1, 11) = CStr(rightYear) & " Kalan": headers(1, 12) = "Toplam Kalan"
            count = selected.count
            If count > 0 Then
                ReDim output(1 To count, 1 To 12): i = 0
                For r = 1 To UBound(people, 1)
                    pid = CC_Text(people(r, 9))
                    If selected.Exists(pid) Then
                        i = i + 1: output(i, 1) = i
                        For c = 2 To 4: output(i, c) = people(r, c): Next c
                        values = counts(pid)
                        For c = 0 To 4: output(i, c + 5) = CLng(values(c)): Next c
                        output(i, 10) = CR_Number(people(r, 5)): output(i, 11) = CR_Number(people(r, 6))
                        output(i, 12) = CDbl(output(i, 10)) + CDbl(output(i, 11))
                    End If
                Next r
            End If
            note = CStr(count) & " personel. Çakışan günler toplamda tek sayılır. Kalanlar güncel bakiyedir; dönem sonu bakiyesi değildir."
            ws.Range("A13").Value2 = "Özet tarih aralığındaki takvim ve mesai günlerini sayar; elle düşülen izin hakkı gününden farklı olabilir. Aynı gün toplamda bir kez sayılır."
        End If
    End If
    oldLast = Application.Max(18, CLng(Val(CC_Text(ws.Range("N2").Value2))))
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A17:L" & CStr(oldLast)).ClearContents
    ws.Range("A17:L17").Value2 = headers
    lastOut = 17 + count
    If count > 0 Then ws.Range("A18:L" & CStr(lastOut)).Value2 = output
    If mode = "detail" Then
        ws.Range("F18:G" & CStr(Application.Max(18, lastOut))).NumberFormat = "dd.mm.yyyy"
    Else
        ws.Range("F18:G" & CStr(Application.Max(oldLast, lastOut))).NumberFormat = "General"
    End If
    ws.Range(ws.Cells(17, 1), ws.Cells(Application.Max(18, lastOut), lastCol)).AutoFilter
    ws.Range("N2").Value2 = lastOut: ws.Range("N3").Value2 = mode: ws.Range("N4").Value2 = lastCol: ws.Range("N5").Value2 = False
    If count = 0 Then note = "Seçtiğiniz filtrelere uygun kayıt bulunamadı."
    ws.Range("A14").Value2 = note
    CompactReportRefresh = ""
    GoTo Done
Invalid:
    ws.Range("N5").Value2 = True
    ws.Range("A14").Value2 = "Kontrol gerekli: " & failure & " Önceki raporu güncellemeden kullanmayın."
    CompactReportRefresh = "ERR:" & failure
    GoTo Done
Failed:
    failure = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Range("A14").Value2 = "Rapor hazırlanamadı: " & failure: ws.Range("N5").Value2 = True
    CompactReportRefresh = "ERR:Rapor hazırlanamadı (" & stage & "): " & failure
Done:
    On Error Resume Next
    If wasProtected Then ws.Protect Password:="2022-tiski", UserInterfaceOnly:=True, AllowFiltering:=True
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
End Function

Public Sub CompactReportRun()
    Dim result As String
    result = CompactReportRefresh()
    If Len(result) > 0 Then MsgBox result, vbExclamation, "Rapor"
End Sub

Public Sub CompactReportPrintLayout()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, rightText As String, leftText As String
    Set ws = CR_Sheet
    lastRow = Application.Max(18, CLng(Val(CC_Text(ws.Range("N2").Value2))))
    lastCol = CLng(Val(CC_Text(ws.Range("N4").Value2)))
    If lastCol < 7 Or lastCol > 12 Then lastCol = 12
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(17, 1), ws.Cells(lastRow, lastCol)).address
    ws.PageSetup.PrintTitleRows = "$17:$17"
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.TopMargin = Application.InchesToPoints(0.85)
    ws.PageSetup.HeaderMargin = Application.InchesToPoints(0.18)
    ws.PageSetup.BottomMargin = Application.InchesToPoints(0.45)
    ws.PageSetup.FooterMargin = Application.InchesToPoints(0.16)
    leftText = "Birim: " & CC_Text(ws.Range("F5").Value2) & vbLf & "Personel: " & CC_Text(ws.Range("B7").Value2)
    If CC_Text(ws.Range("N3").Value2) = "remaining" Then
        rightText = "Güncel kalan izinler" & vbLf & "İzin yılları: " & CompactLeftYear() & " / " & CompactRightYear()
    Else
        rightText = Format$(CDate(CR_Date(ws.Range("B9").Value2)), "dd.mm.yyyy") & " - " & Format$(CDate(CR_Date(ws.Range("F9").Value2)), "dd.mm.yyyy") & vbLf & "Tür: " & CC_Text(ws.Range("F7").Value2)
    End If
    ws.PageSetup.LeftHeader = "&8" & Replace(left$(leftText, 220), "&", "&&")
    ws.PageSetup.CenterHeader = "&12&B" & Replace("TİSKİ / " & CC_Text(ws.Range("B5").Value2), "&", "&&") & "&B"
    ws.PageSetup.RightHeader = "&8" & Replace(left$(rightText, 220), "&", "&&")
    ws.PageSetup.LeftFooter = "&8" & Format$(Now, "dd.mm.yyyy HH:nn")
    ws.PageSetup.CenterFooter = "&8İnsan Kaynakları"
    ws.PageSetup.RightFooter = "&8Sayfa &P / &N"
End Sub

Public Sub CompactReportPrint()
    Dim result As String
    result = CompactReportRefresh()
    If Len(result) > 0 Then MsgBox result, vbExclamation, "Rapor": Exit Sub
    CompactReportPrintLayout
    CR_Sheet.PrintPreview
End Sub

Public Sub CompactReportPDF()
    Dim result As String, path As Variant, initial As String
    On Error GoTo Failed
    result = CompactReportRefresh()
    If Len(result) > 0 Then MsgBox result, vbExclamation, "Rapor": Exit Sub
    initial = ThisWorkbook.path
    If Len(initial) > 0 Then initial = initial & Application.PathSeparator
    initial = initial & "Izin_Raporu_" & Format$(Now, "yyyymmdd_hhnn") & ".pdf"
    path = Application.GetSaveAsFilename(InitialFileName:=initial, FileFilter:="PDF Dosyası (*.pdf), *.pdf", title:="Raporu PDF olarak kaydet")
    If VarType(path) = vbBoolean Then Exit Sub
    result = CompactReportPDFFile(CStr(path))
    If Len(result) > 0 Then MsgBox result, vbExclamation, "Rapor": Exit Sub
    MsgBox "Rapor PDF olarak kaydedildi.", vbInformation, "Rapor"
    Exit Sub
Failed:
    MsgBox "PDF oluşturulamadı: " & Err.Description, vbExclamation, "Rapor"
End Sub

Public Function CompactReportPDFFile(ByVal path As String) As String
    Dim result As String
    On Error GoTo Failed
    If Len(Trim$(path)) = 0 Then CompactReportPDFFile = "ERR:PDF dosya yolunu seçin.": Exit Function
    result = CompactReportRefresh()
    If Len(result) > 0 Then CompactReportPDFFile = result: Exit Function
    CompactReportPrintLayout
    CR_Sheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False
    CompactPrepareForSave
    Exit Function
Failed:
    CompactReportPDFFile = "ERR:PDF oluşturulamadı: " & Err.Description
End Function

Public Sub CompactReportLists()
    Dim ws As Worksheet, people As Variant, types As Variant, units As Object, values As Variant
    Dim i As Long, j As Long, n As Long, lastRow As Long, temporary As String, names() As String, output() As Variant
    Dim oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Done
    Application.EnableEvents = False
    Set ws = CR_Sheet: Set units = CC_Dict()
    lastRow = CC_Last(2)
    ws.Range("P2:S" & CStr(Application.Max(320, lastRow))).ClearContents
    ws.Range("P2").Value2 = "Tümü": ws.Range("Q2").Value2 = "Tümü": ws.Range("R2").Value2 = "Tümü"
    If lastRow >= 10 Then
        people = CC_Sheet.Range("A10:I" & CStr(lastRow)).Value2
        ReDim output(1 To UBound(people, 1), 1 To 1)
        For i = 1 To UBound(people, 1)
            If Len(CC_Text(people(i, 9))) > 0 And Len(CC_Text(people(i, 2))) > 0 Then
                n = n + 1: output(n, 1) = CC_Text(people(i, 2)) & " - " & CC_Text(people(i, 3))
                temporary = CC_Text(people(i, 4))
                If Len(temporary) > 0 Then units(temporary) = True
            End If
        Next i
        If n > 0 Then ws.Range("Q3:Q" & CStr(n + 2)).Value2 = output
    End If
    If units.count > 0 Then
        ReDim names(1 To units.count): i = 0
        For Each values In units.keys: i = i + 1: names(i) = CStr(values): Next values
        For i = 1 To UBound(names) - 1
            For j = i + 1 To UBound(names)
                If StrComp(names(i), names(j), vbTextCompare) > 0 Then temporary = names(i): names(i) = names(j): names(j) = temporary
            Next j
        Next i
        ReDim output(1 To units.count, 1 To 1)
        For i = 1 To units.count: output(i, 1) = names(i): Next i
        ws.Range("P3:P" & CStr(units.count + 2)).Value2 = output
    End If
    lastRow = CC_Sheet.Cells(CC_Sheet.Rows.count, 37).End(xlUp).row
    If lastRow >= 2 Then
        types = CC_Sheet.Range("AK2:AK" & CStr(lastRow)).Value2
        ws.Range("R3:R" & CStr(lastRow + 1)).Value2 = types
    End If
    ws.Range("S2").Value2 = "Kalan İzinler": ws.Range("S3").Value2 = "Kullanılan İzinler": ws.Range("S4").Value2 = "Personel Özeti"
    On Error Resume Next
    ThisWorkbook.names("ReportBirimler").Delete: ThisWorkbook.names("ReportPersonel").Delete
    ThisWorkbook.names("ReportTurler").Delete: ThisWorkbook.names("ReportModlar").Delete
    On Error GoTo Done
    ThisWorkbook.names.Add name:="ReportBirimler", RefersTo:="='Raporlar'!$P$2:$P$" & CStr(units.count + 2)
    ThisWorkbook.names.Add name:="ReportPersonel", RefersTo:="='Raporlar'!$Q$2:$Q$" & CStr(Application.Max(2, n + 2))
    ThisWorkbook.names.Add name:="ReportTurler", RefersTo:="='Raporlar'!$R$2:$R$" & CStr(Application.Max(2, lastRow + 1))
    ThisWorkbook.names.Add name:="ReportModlar", RefersTo:="='Raporlar'!$S$2:$S$4"
Done:
    Application.EnableEvents = oldEvents
End Sub

' Preserve filter criteria while assigning complete rectangular matrices.
Public Sub CompactAssignBlock(ByVal ws As Worksheet, ByVal address As String, ByVal values As Variant, Optional ByVal clearAddress As String = "")
    Dim saved As Collection, first As Variant, second As Variant
    Dim i As Long, operation As Long, hasFirst As Boolean, hasSecond As Boolean
    Dim oldEvents As Boolean, failureNumber As Long, failureText As String
    oldEvents = Application.EnableEvents
    Set saved = New Collection
    On Error GoTo Failed
    Application.EnableEvents = False
    If ws.FilterMode And ws.AutoFilterMode Then
        For i = 1 To ws.AutoFilter.filters.count
            If ws.AutoFilter.filters(i).On Then
                first = Empty: second = Empty: operation = 0
                On Error Resume Next
                Err.Clear: first = ws.AutoFilter.filters(i).Criteria1: hasFirst = (Err.Number = 0)
                Err.Clear: second = ws.AutoFilter.filters(i).Criteria2: hasSecond = (Err.Number = 0)
                Err.Clear: operation = ws.AutoFilter.filters(i).Operator
                On Error GoTo Failed
                saved.Add Array(i, operation, hasFirst, first, hasSecond, second)
            End If
        Next i
        ws.ShowAllData
    End If
    If Len(clearAddress) > 0 Then ws.Range(clearAddress).ClearContents
    If Not IsEmpty(values) Then ws.Range(address).Value2 = values
    CB_RestoreFilters ws, saved
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    failureNumber = Err.Number: failureText = Err.Description
    On Error Resume Next
    CB_RestoreFilters ws, saved
    Application.EnableEvents = oldEvents
    Err.Raise failureNumber, "CompactAssignBlock", failureText
End Sub

Private Sub CB_RestoreFilters(ByVal ws As Worksheet, ByVal saved As Collection)
    Dim item As Variant
    If saved Is Nothing Then Exit Sub
    If saved.count = 0 Or Not ws.AutoFilterMode Then Exit Sub
    For Each item In saved
        If CBool(item(2)) And CBool(item(4)) Then
            ws.AutoFilter.Range.AutoFilter Field:=CLng(item(0)), Criteria1:=item(3), Operator:=CLng(item(1)), Criteria2:=item(5)
        ElseIf CBool(item(4)) Then
            ws.AutoFilter.Range.AutoFilter Field:=CLng(item(0)), Operator:=CLng(item(1)), Criteria2:=item(5)
        ElseIf CBool(item(2)) Then
            If CLng(item(1)) = 0 Then
                ws.AutoFilter.Range.AutoFilter Field:=CLng(item(0)), Criteria1:=item(3)
            Else
                ws.AutoFilter.Range.AutoFilter Field:=CLng(item(0)), Criteria1:=item(3), Operator:=CLng(item(1))
            End If
        End If
    Next item
End Sub


' Native Forms search results. The plain worksheet query remains editable.
Public Sub CompactSearchOpen()
    Dim ws As Worksheet, oldEvents As Boolean, oldScreen As Boolean, lastRow As Long
    Dim errorText As String
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    CU_Search False, False
    lastRow = ws.Cells(ws.Rows.count, 46).End(xlUp).row
    CS_ShowList Application.Max(0, lastRow - 1)
    ws.Activate
    ws.Range("C5").Select
    GoTo Done
Failed:
    errorText = Err.Description
Done:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If Len(errorText) > 0 Then MsgBox "Personel araması açılamadı: " & errorText, vbExclamation
End Sub

Public Sub CompactSearchClose()
    Dim ws As Worksheet, oldEvents As Boolean, oldScreen As Boolean, wasProtected As Boolean
    Dim errorNumber As Long, errorText As String
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects
    If wasProtected Then ws.Unprotect Password:="2022-tiski"
    CS_DeleteShapes ws
    GoTo Done
Failed:
    errorNumber = Err.Number: errorText = Err.Description
Done:
    On Error GoTo 0
    If Not ws Is Nothing Then
        If wasProtected Then ws.Protect Password:="2022-tiski", UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If errorNumber <> 0 Then Err.Raise errorNumber, "CompactSearchClose", errorText
End Sub

Public Sub CompactSearchChoose()
    Dim ws As Worksheet, list As Object, selected As Long, pid As String, rosterRow As Long
    Dim oldEvents As Boolean, oldScreen As Boolean, errorText As String
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    Set list = ws.ListBoxes("PS_Results")
    selected = CLng(list.ListIndex)
    If selected < 1 Or selected > list.ListCount Then GoTo Done
    pid = CStr(ws.Cells(selected + 1, 46).Value2)
    If Len(pid) = 0 Then GoTo Done
    rosterRow = CP_PersonRow(pid)
    If rosterRow < 10 Then
        CompactSearchClose
        MsgBox "Personel listesi değişti. Aramayı yeniden yapın.", vbInformation
        GoTo Done
    End If
    ws.Range("C5").Value2 = CStr(ws.Cells(rosterRow, 2).Value2) & " - " & CStr(ws.Cells(rosterRow, 3).Value2)
    CompactSearchClose
    CompactSelectPerson rosterRow
    GoTo Done
Failed:
    errorText = Err.Description
Done:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If Len(errorText) > 0 Then MsgBox "Personel seçilemedi: " & errorText, vbExclamation
End Sub

Private Sub CS_ShowList(ByVal matchCount As Long)
    Dim ws As Worksheet, wasProtected As Boolean, oldEvents As Boolean, oldScreen As Boolean
    Dim frame As Shape, title As Shape, closeButton As Shape, results As Shape, hint As Shape
    Dim left As Double, top As Double, width As Double, listHeight As Double, frameHeight As Double
    Dim errorNumber As Long, errorText As String, list As Object, items As Variant, itemIndex As Long
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects
    If wasProtected Then ws.Unprotect Password:="2022-tiski"
    CS_DeleteShapes ws
    left = ws.Range("A7").left + 3
    top = ws.Range("A7").top + 2
    width = Application.Max(180, ws.Range("A7:H7").width - 6)
    listHeight = Application.Max(1, Application.Min(8, matchCount)) * 18 + 7
    frameHeight = 27 + listHeight + 23
    Set frame = ws.Shapes.AddShape(msoShapeRectangle, left, top, width, frameHeight)
    frame.name = "PS_Back"
    frame.fill.ForeColor.RGB = RGB(255, 255, 255)
    frame.Line.ForeColor.RGB = RGB(8, 125, 140)
    frame.Line.Weight = 1
    frame.Placement = xlFreeFloating
    frame.DrawingObject.PrintObject = False
    Set title = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left + 1, top + 1, width - 2, 25)
    title.name = "PS_Title"
    title.fill.ForeColor.RGB = RGB(8, 125, 140)
    title.Line.Visible = msoFalse
    With title.TextFrame
        .Characters.text = CStr(matchCount) & " PERSONEL  /  Seçmek için satıra tıklayın"
        .Characters.Font.name = "Calibri"
        .Characters.Font.Size = 10
        .Characters.Font.Bold = True
        .Characters.Font.Color = RGB(255, 255, 255)
        .MarginLeft = 8: .MarginRight = 30: .MarginTop = 5
    End With
    title.Placement = xlFreeFloating
    title.DrawingObject.PrintObject = False
    If matchCount > 0 Then
        Set results = ws.Shapes.AddFormControl(xlListBox, left + 5, top + 28, width - 10, listHeight)
        results.name = "PS_Results"
        results.Placement = xlFreeFloating
        results.Locked = False
        results.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!CompactSearchChoose"
        Set list = ws.ListBoxes("PS_Results")
        list.ListFillRange = ""
        items = CS_ResultCache()
        For itemIndex = 1 To matchCount
            list.AddItem CStr(items(itemIndex, 2))
        Next itemIndex
        list.LinkedCell = ""
        list.MultiSelect = xlNone
        list.Display3DShading = False
        list.PrintObject = False
        results.ZOrder msoBringToFront
    Else
        Set hint = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left + 6, top + 29, width - 12, listHeight)
        hint.name = "PS_Empty"
        hint.Line.Visible = msoFalse
        With hint.TextFrame
            .Characters.text = "Eşleşen personel bulunamadı. İsmi veya sicili kontrol edin."
            .Characters.Font.name = "Calibri": .Characters.Font.Size = 10
            .Characters.Font.Color = RGB(23, 60, 75)
            .MarginLeft = 4: .MarginTop = 4
        End With
        hint.Placement = xlFreeFloating: hint.DrawingObject.PrintObject = False
    End If
    Set hint = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left + 5, top + frameHeight - 21, width - 10, 20)
    hint.name = "PS_Help"
    hint.Line.Visible = msoFalse
    With hint.TextFrame
        .Characters.text = "Aramayı değiştirmek için üstte yazın ve Enter'a basın."
        .Characters.Font.name = "Calibri": .Characters.Font.Size = 9
        .Characters.Font.Color = RGB(100, 128, 139)
        .MarginLeft = 3: .MarginTop = 3
    End With
    hint.Placement = xlFreeFloating: hint.DrawingObject.PrintObject = False
    Set closeButton = ws.Shapes.AddShape(msoShapeRectangle, left + width - 26, top + 3, 22, 20)
    closeButton.name = "PS_Close"
    closeButton.fill.ForeColor.RGB = RGB(255, 255, 255)
    closeButton.Line.Visible = msoFalse
    closeButton.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!CompactSearchClose"
    With closeButton.TextFrame
        .Characters.text = "x"
        .Characters.Font.name = "Calibri": .Characters.Font.Size = 12
        .Characters.Font.Bold = True: .Characters.Font.Color = RGB(8, 125, 140)
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    closeButton.Placement = xlFreeFloating: closeButton.DrawingObject.PrintObject = False
    closeButton.ZOrder msoBringToFront
    GoTo Done
Failed:
    errorNumber = Err.Number: errorText = Err.Description
Done:
    On Error GoTo 0
    If Not ws Is Nothing Then
        If wasProtected Then ws.Protect Password:="2022-tiski", UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If errorNumber <> 0 Then Err.Raise errorNumber, "CS_ShowList", errorText
End Sub

Private Sub CS_DeleteShapes(ByVal ws As Worksheet)
    Dim index As Long
    For index = ws.Shapes.count To 1 Step -1
        If left$(ws.Shapes(index).name, 3) = "PS_" Then ws.Shapes(index).Delete
    Next index
End Sub


Private Function CS_ResultCache(Optional ByVal newResults As Variant) As Variant
    Static results As Variant
    If Not IsMissing(newResults) Then results = newResults
    CS_ResultCache = results
End Function


Private Function CC_ImportedValuesUnchanged(ByVal pid As String, ByVal code As String, ByVal firstDay As Double, ByVal lastDay As Double, ByVal days As Double, ByVal editId As String) As Boolean
    Dim values As Variant, r As Long, lastRow As Long
    If Len(editId) = 0 Then Exit Function
    lastRow = CC_Last(18)
    If lastRow < 2 Then Exit Function
    values = CC_Sheet.Range("R2:AB" & CStr(lastRow)).Value2
    For r = 1 To UBound(values, 1)
        If CC_Text(values(r, 1)) = editId Then
            If CC_Text(values(r, 2)) <> pid Then Exit Function
            If CC_Bool(values(r, 7)) Then Exit Function
            If InStr(1, CC_Normal(CC_Text(values(r, 11))), "kaynak", vbBinaryCompare) = 0 Then Exit Function
            If CompactTypeCode(CC_Text(values(r, 3))) <> code Then Exit Function
            If Not CC_DateOK(values(r, 4)) Or Not CC_DateOK(values(r, 5)) Then Exit Function
            If Not IsNumeric(values(r, 6)) Then Exit Function
            CC_ImportedValuesUnchanged = CDbl(values(r, 4)) = firstDay And CDbl(values(r, 5)) = lastDay And CDbl(values(r, 6)) = days
            Exit Function
        End If
    Next r
End Function


' Restore real internal links after bulk value imports; no leave/balance changes.
Public Sub CompactEnsurePersonLinks()
    Dim ws As Worksheet, cell As Range, r As Long, lastRow As Long, needsLink As Boolean
    Dim wasProtected As Boolean, oldEvents As Boolean, errorNumber As Long, errorText As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects
    If wasProtected Then ws.Unprotect Password:="2022-tiski"
    lastRow = CP_LastPerson()
    For r = 10 To lastRow
        Set cell = ws.Cells(r, 8)
        If Len(CC_Text(ws.Cells(r, 2).Value2)) > 0 And Len(CC_Text(ws.Cells(r, 3).Value2)) > 0 And Len(CC_Text(ws.Cells(r, 9).Value2)) > 0 Then
            If CC_Text(cell.Value2) <> "+" Then cell.Value2 = "+"
            needsLink = cell.Hyperlinks.count <> 1
            If Not needsLink Then
                needsLink = Len(cell.Hyperlinks(1).address) <> 0 Or cell.Hyperlinks(1).SubAddress <> "'İzin Girişi'!J9"
            End If
            If needsLink Then
                If cell.Hyperlinks.count > 0 Then cell.Hyperlinks.Delete
                ws.Hyperlinks.Add Anchor:=cell, address:="", SubAddress:="'İzin Girişi'!J9", TextToDisplay:="+"
            End If
        Else
            If cell.Hyperlinks.count > 0 Then cell.Hyperlinks.Delete
            cell.ClearContents
        End If
    Next r
    If lastRow >= 10 Then
        With ws.Range("H10:H" & CStr(lastRow)).Font
            .name = "Calibri": .Size = 17: .Bold = True
            .Color = RGB(8, 127, 140): .Underline = xlUnderlineStyleNone
        End With
    End If
    GoTo Done
Failed:
    errorNumber = Err.Number: errorText = Err.Description
Done:
    On Error GoTo 0
    If Not ws Is Nothing Then
        If wasProtected Then ws.Protect Password:="2022-tiski", UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    Application.EnableEvents = oldEvents
    If errorNumber <> 0 Then Err.Raise errorNumber, "CompactEnsurePersonLinks", errorText
End Sub




Public Sub CompactHistoryInit()
    Dim ws As Worksheet, data As Variant, r As Long, yr As Long, yearRow As Long, typeRow As Long
    Dim years(2000 To 2100) As Boolean, codes As Variant, i As Long, serial As Double, lastLeave As Long
    Dim wasProtected As Boolean, oldEvents As Boolean, errorNumber As Long, errorText As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects
    If wasProtected Then ws.Unprotect Password:="2022-tiski"
    yr = CompactLeftYear(): If yr >= 2000 And yr <= 2100 Then years(yr) = True
    yr = CompactRightYear(): If yr >= 2000 And yr <= 2100 Then years(yr) = True
    If yr < 2100 Then years(yr + 1) = True
    lastLeave = CU_LastRow(ws, 18)
    data = ws.Range("R2:AB" & CStr(lastLeave)).Value2
    For r = 1 To UBound(data, 1)
        For i = 4 To 5
            serial = CU_Date(data(r, i))
            If serial > 0 Then
                yr = Year(CDate(serial))
                If yr >= 2000 And yr <= 2100 Then years(yr) = True
            End If
        Next i
    Next r
    ws.Range("AV2:AW104").ClearContents
    ws.Range("AV2").Value2 = "Tümü": yearRow = 2
    For yr = 2000 To 2100
        If years(yr) Then
            yearRow = yearRow + 1
            ws.Cells(yearRow, 48).Value2 = CStr(yr)
        End If
    Next yr
    ws.Range("AW2").Value2 = "Tümü"
    codes = Split("YILLIK|RAPOR|MAZERET|BABALIK|EVLILIK|VEFAT|ANALIK|REFAKAT|UCRETSIZ|IDARI|GOREVLI|DIGER", "|")
    For i = LBound(codes) To UBound(codes)
        ws.Cells(i + 3, 49).Value2 = CompactTypeLabel(CStr(codes(i)))
    Next i
    typeRow = UBound(codes) + 3
    On Error Resume Next
    ThisWorkbook.names("CompactHistoryYears").Delete
    ThisWorkbook.names("CompactHistoryTypes").Delete
    Err.Clear
    On Error GoTo Failed
    ThisWorkbook.names.Add name:="CompactHistoryYears", RefersTo:="='İzin Girişi'!$AV$2:$AV$" & CStr(yearRow)
    ThisWorkbook.names.Add name:="CompactHistoryTypes", RefersTo:="='İzin Girişi'!$AW$2:$AW$" & CStr(typeRow)
    If Len(CC_Text(ws.Range("J24").Value2)) = 0 Then
        ws.Range("J24").Value2 = CompactRightYear()
    ElseIf Not CH_All(ws.Range("J24").Value2) Then
        If Not IsNumeric(ws.Range("J24").Value2) Then
            ws.Range("J24").Value2 = CompactRightYear()
        ElseIf CDbl(ws.Range("J24").Value2) < 2000 Or CDbl(ws.Range("J24").Value2) > 2100 Or CDbl(ws.Range("J24").Value2) <> Fix(CDbl(ws.Range("J24").Value2)) Then
            ws.Range("J24").Value2 = CompactRightYear()
        End If
    End If
    If Len(CC_Text(ws.Range("L24").Value2)) = 0 Then ws.Range("L24").Value2 = "Tümü"
    If Not CH_All(ws.Range("L24").Value2) Then
        If Len(CompactTypeCode(CC_Text(ws.Range("L24").Value2))) = 0 Then ws.Range("L24").Value2 = "Tümü"
    End If
    With ws.Range("J24").MergeArea.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CompactHistoryYears"
        .IgnoreBlank = True: .InCellDropdown = True: .ShowError = True: .ShowInput = False
        .ErrorTitle = "İzin yılı": .ErrorMessage = "Listeden bir yıl veya Tümü seçin."
    End With
    With ws.Range("L24").MergeArea.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CompactHistoryTypes"
        .IgnoreBlank = True: .InCellDropdown = True: .ShowError = True: .ShowInput = False
        .ErrorTitle = "İzin türü": .ErrorMessage = "Listeden bir izin türü veya Tümü seçin."
    End With
    ws.Range("J24").MergeArea.Locked = False
    ws.Range("L24").MergeArea.Locked = False
    ws.columns("AV:AW").Hidden = True
    ws.Range("AI11").Value2 = CH_Page(ws.Range("AI11").Value2)
Done:
    On Error GoTo 0
    If Not ws Is Nothing Then
        If wasProtected Then ws.Protect Password:="2022-tiski", UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    Application.EnableEvents = oldEvents
    If errorNumber <> 0 Then Err.Raise errorNumber, "CompactHistoryInit", errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume Done
End Sub

Public Sub CompactHistoryReset(Optional ByVal refreshView As Boolean = True)
    Dim oldEvents As Boolean, errorNumber As Long, errorText As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    ThisWorkbook.Worksheets("İzin Girişi").Range("AI11").Value2 = 1
    If refreshView Then CU_ShowSelected
    GoTo Done
Failed:
    errorNumber = Err.Number: errorText = Err.Description
Done:
    Application.EnableEvents = oldEvents
    If errorNumber <> 0 Then Err.Raise errorNumber, "CompactHistoryReset", errorText
End Sub

Public Sub CompactHistoryPrevious()
    CH_MovePage -1
End Sub

Public Sub CompactHistoryNext()
    CH_MovePage 1
End Sub

Private Sub CH_MovePage(ByVal movement As Long)
    Dim ws As Worksheet, oldEvents As Boolean, errorText As String, nextPage As Long
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    nextPage = CH_Page(ws.Range("AI11").Value2) + movement
    If nextPage < 1 Then nextPage = 1
    ws.Range("AI11").Value2 = nextPage
    CU_ShowSelected
    GoTo Done
Failed:
    errorText = Err.Description
Done:
    Application.EnableEvents = oldEvents
    If Len(errorText) > 0 Then MsgBox "İzin listesi gösterilemedi: " & errorText, vbExclamation
End Sub

Private Function CH_Page(ByVal value As Variant) As Long
    CH_Page = 1
    If IsNumeric(value) Then
        If CDbl(value) >= 1 And CDbl(value) < 1000000 Then CH_Page = CLng(Fix(CDbl(value)))
    End If
End Function

Private Function CH_All(ByVal value As Variant) As Boolean
    Dim text As String
    text = CC_Normal(CC_Text(value))
    CH_All = Len(text) = 0 Or text = "tumu" Or text = "tum yillar" Or text = "tum izinler"
End Function

Private Function CH_ViewYear(ByVal value As Variant) As Long
    If CH_All(value) Then Exit Function
    If IsNumeric(value) Then
        If CDbl(value) >= 2000 And CDbl(value) <= 2100 And CDbl(value) = Fix(CDbl(value)) Then
            CH_ViewYear = CLng(value)
            Exit Function
        End If
    End If
    CH_ViewYear = CompactRightYear()
End Function

Private Function CH_Before(ByVal firstIndex As Long, ByVal secondIndex As Long, ByRef data As Variant) As Boolean
    Dim firstDay As Double, secondDay As Double
    firstDay = CU_Date(data(firstIndex, 4)): secondDay = CU_Date(data(secondIndex, 4))
    If firstDay = secondDay Then CH_Before = firstIndex < secondIndex Else CH_Before = firstDay > secondDay
End Function

Private Sub CH_SortIndices(ByRef indices() As Long, ByRef data As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Long, swap As Long
    i = low: j = high: pivot = indices((low + high) \ 2)
    Do While i <= j
        Do While CH_Before(indices(i), pivot, data): i = i + 1: Loop
        Do While CH_Before(pivot, indices(j), data): j = j - 1: Loop
        If i <= j Then
            swap = indices(i): indices(i) = indices(j): indices(j) = swap
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then CH_SortIndices indices, data, low, j
    If i < high Then CH_SortIndices indices, data, i, high
End Sub


' A new year changes the visible year window; historical rights/charges stay replayable.
Public Sub CompactYearOpenUI()
    Dim ws As Worksheet, target As Variant, currentYear As Long, result As String, backupPath As String
    Dim explanation As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    currentYear = CompactRightYear()
    If ThisWorkbook.ReadOnly Then MsgBox "Dosya salt okunur. Önce düzenlenebilir bir kopyasını kaydedin.", vbExclamation: Exit Sub
    target = Application.InputBox("Açılacak yılı girin. Yeni yılın izin haklarını daha sonra boş sütuna elle girebilirsiniz.", "Yeni Yıl Açılışı", currentYear + 1, Type:=1)
    If VarType(target) = vbBoolean Then Exit Sub
    If Not IsNumeric(target) Then Exit Sub
    If CDbl(target) <> Fix(CDbl(target)) Then MsgBox "Yılı tam sayı olarak girin.", vbExclamation: Exit Sub
    If CLng(target) = currentYear Then MsgBox CStr(currentYear) & " yılı zaten açık; hiçbir bilgi değiştirilmedi.", vbInformation: Exit Sub
    If CLng(target) <> currentYear + 1 Or CLng(target) > 2100 Then MsgBox "Yalnız sıradaki yılı açabilirsiniz: " & CStr(currentYear + 1) & ".", vbExclamation: Exit Sub
    explanation = CStr(currentYear) & " yılının kalan günleri sol sütuna taşınacak." & vbCrLf & _
        CStr(CLng(target)) & " izin hakkı sütunu boş açılacak; hakları siz gireceksiniz." & vbCrLf & _
        "Yeni yılın izin listesi temiz başlayacak. Geçmiş izin ve raporlar, eski yıl raporlarından görülebilecek." & vbCrLf & _
        "Personeller korunacak. İşlemden önce bu dosyanın kapanış yedeği alınacak." & vbCrLf & vbCrLf & "Devam edilsin mi?"
    If MsgBox(explanation, vbYesNo + vbQuestion + vbDefaultButton2, "Yeni Yıl Açılışı") <> vbYes Then Exit Sub
    backupPath = CY_YearBackup(currentYear)
    result = CompactOpenNewYear(CLng(target), True)
    If Len(result) > 0 Then
        MsgBox result & vbCrLf & "Kapanış yedeği: " & backupPath, vbExclamation, "Yıl açılamadı"
    Else
        MsgBox CStr(CLng(target)) & " yılı açıldı. Yeni izin haklarını boş sağ sütuna girebilirsiniz." & vbCrLf & _
            "Kapanış yedeği: " & backupPath, vbInformation, "Yeni Yıl Açılışı"
    End If
    Exit Sub
Failed:
    MsgBox "Yıl açılışı tamamlanamadı: " & Err.Description & vbCrLf & "Mevcut bilgiler korunur; yedek alınamazsa yıl değiştirilmez.", vbExclamation, "Yeni Yıl Açılışı"
End Sub

Private Function CY_YearBackup(ByVal closingYear As Long) As String
    Dim directory As String, path As String, baseName As String, suffix As Long
    If Len(ThisWorkbook.path) = 0 Then Err.Raise vbObjectError + 401, "CY_YearBackup", "Önce dosyayı bir klasöre kaydedin."
    directory = ThisWorkbook.path & Application.PathSeparator & "Yil_Gecisi_Yedekleri"
    If Len(Dir$(directory, vbDirectory)) = 0 Then MkDir directory
    baseName = left$(ThisWorkbook.name, InStrRev(ThisWorkbook.name, ".") - 1) & "_Kapanis_" & CStr(closingYear) & "_" & Format$(Now, "yyyymmdd_hhnnss")
    path = directory & Application.PathSeparator & baseName & ".xlsm"
    Do While Len(Dir$(path)) > 0
        suffix = suffix + 1
        path = directory & Application.PathSeparator & baseName & "_" & CStr(suffix) & ".xlsm"
    Loop
    CompactPrepareForSave
    ThisWorkbook.SaveCopyAs path
    If Len(Dir$(path)) = 0 Then Err.Raise vbObjectError + 402, "CY_YearBackup", "Kapanış yedeği oluşturulamadı."
    CY_YearBackup = path
End Function

Public Function CompactOpenNewYear(ByVal targetYear As Long, Optional ByVal carry As Boolean = True) As String
    Dim ws As Worksheet, pu As Worksheet, report As Worksheet, item As Variant, sheet As Worksheet
    Dim oldEvents As Boolean, oldScreen As Boolean, snapshotReady As Boolean, changed As Boolean
    Dim currentYear As Long, lastPeople As Long, lastRights As Long, lastLeaves As Long, lastReport As Long, reportEnd As Long
    Dim rights As Variant, allocations As Variant, roster As Variant, state As Variant, headers As Variant, formats(1 To 2) As String
    Dim puData As Variant, puSummary As Variant, puYear As Variant, puMonth As Variant, reportData As Variant, reportMeta As Variant
    Dim form As Collection, reportFields As Collection, protection As Collection, reportFilters As Collection, puFilters As Collection
    Dim reportFilterRange As String, puFilterRange As String, reportMode As String, values() As Variant, result As String, failure As String
    Dim r As Long, restoreEnd As Long, normalFormat As String
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("İzin Girişi")
    Set pu = ThisWorkbook.Worksheets("Puantaj")
    Set report = ThisWorkbook.Worksheets("Raporlar")
    currentYear = CompactRightYear()
    If targetYear = currentYear Then Exit Function
    If targetYear <> currentYear + 1 Or targetYear > 2100 Then CompactOpenNewYear = "ERR:Yalnız sıradaki yıl açılabilir: " & CStr(currentYear + 1) & ".": Exit Function
    If ThisWorkbook.ReadOnly Then CompactOpenNewYear = "ERR:Dosya salt okunur; önce düzenlenebilir kopyasını kaydedin.": Exit Function
    lastPeople = Application.Max(10, CP_LastPerson())
    lastRights = Application.Max(1, CC_Last(30)): lastLeaves = Application.Max(2, CC_Last(18))
    lastReport = Application.Max(18, CLng(Val(CC_Text(report.Range("N2").Value2))))
    rights = ws.Range("AD1:AG" & CStr(lastRights)).Value2
    allocations = ws.Range("AA2:AA" & CStr(lastLeaves)).Value2
    roster = ws.Range("E10:G" & CStr(lastPeople)).Value2
    state = ws.Range("AI1:AI20").Value2: headers = ws.Range("E9:F9").Value2
    formats(1) = ws.Range("E10").NumberFormat: formats(2) = ws.Range("F10").NumberFormat
    Set form = CY_CaptureCells(ws, Array("A7", "A8", "C5", "J5", "J6", "J9", "M9", "L12", "L14", "J16", "J24", "L24", "N24"))
    puYear = pu.Range("B5").Value2: puMonth = pu.Range("D5").Value2
    puData = pu.Range("A10:G" & CStr(lastPeople)).Value2: puSummary = Array(pu.Range("F5").Value2, pu.Range("H5").Value2)
    reportData = report.Range("A17:L" & CStr(lastReport)).Value2: reportMeta = report.Range("N2:N5").Value2
    reportMode = CStr(report.Range("N3").Value2)
    Set reportFields = CY_CaptureCells(report, Array("B5", "F5", "B7", "F7", "B9", "F9", "B11", "F11", "A14"))
    Set protection = New Collection
    For Each sheet In ThisWorkbook.Worksheets
        protection.Add Array(sheet.name, sheet.ProtectContents Or sheet.ProtectDrawingObjects)
    Next sheet
    Set puFilters = CY_YearFilters(pu, puFilterRange)
    Set reportFilters = CY_YearFilters(report, reportFilterRange)
    snapshotReady = True
    Application.EnableEvents = False: Application.ScreenUpdating = False
    For Each sheet In ThisWorkbook.Worksheets: sheet.Unprotect Password:="2022-tiski": Next sheet
    changed = True
    result = CompactRollYear(targetYear)
    If Len(result) > 0 Then failure = result: GoTo Rollback
    CompactHistoryInit
    If Not carry And CP_LastPerson() >= 10 Then
        ReDim values(1 To lastPeople - 9, 1 To 1)
        For r = 1 To UBound(values, 1): values(r, 1) = 0#: Next r
        CompactAssignBlock ws, "E10:E" & CStr(lastPeople), values
        result = CompactEditRemaining(ws.Range("E10:E" & CStr(lastPeople)))
        If Len(result) > 0 Then failure = result: GoTo Rollback
    End If
    ' Zero new-year rights are intentionally left undefined until manually entered.
    normalFormat = formats(2)
    If normalFormat = "0.##;-0.##;;@" Or normalFormat = "General;General;;@" Then normalFormat = "General"
    ws.Range("E10:E" & CStr(lastPeople)).NumberFormat = normalFormat
    ws.Range("F10:F" & CStr(lastPeople)).NumberFormat = "General;General;;@"
    ws.Range("F10:F" & CStr(lastPeople)).ClearContents
    CompactClear
    CompactSearchClear
    CompactHistoryReset
    ws.Range("J24").Value2 = targetYear: ws.Range("L24").Value2 = "Tümü": ws.Range("AI11").Value2 = 1
    CU_ShowSelected
    ws.Range("A8").Value2 = CStr(targetYear) & " yılı açıldı · Önceki yılın kalan günleri solda; yeni yılın haklarını boş sağ sütuna girin."
    pu.Range("B5").Value2 = targetYear: pu.Range("D5").Value2 = 1
    CompactRefreshAttendance
    report.Range("B5").Value2 = "Kalan İzinler": report.Range("F5").Value2 = "Tümü"
    report.Range("B7").Value2 = "Tümü": report.Range("F7").Value2 = "Tümü"
    report.Range("B11").Value2 = targetYear: report.Range("F11").Value2 = 0
    report.Range("B9").Value2 = CDbl(DateSerial(targetYear, 1, 1))
    report.Range("F9").Value2 = CDbl(DateSerial(targetYear, 12, 31))
    CompactPrimeRoster
    CompactReportsDirty
    result = CompactReportRefresh()
    If Len(result) > 0 Then failure = result: GoTo Rollback
    ws.Activate
    ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
    ws.Range("F10").Select
    GoTo Done
Failed:
    failure = "ERR:Yıl açılamadı: " & Err.Description
    If snapshotReady And changed Then Resume Rollback
    Resume Done
Rollback:
    On Error Resume Next
    restoreEnd = Application.Max(lastRights, CC_Last(30))
    ws.Range("AD1:AG" & CStr(restoreEnd)).ClearContents
    ws.Range("AD1:AG" & CStr(lastRights)).Value2 = rights
    ws.Range("AA2:AA" & CStr(lastLeaves)).Value2 = allocations
    ws.Range("AI1:AI20").Value2 = state: ws.Range("E9:F9").Value2 = headers
    CompactAssignBlock ws, "E10:G" & CStr(lastPeople), roster
    ws.Range("E10:E" & CStr(lastPeople)).NumberFormat = formats(1)
    ws.Range("F10:F" & CStr(lastPeople)).NumberFormat = formats(2)
    CY_RestoreCells ws, form
    pu.Range("B5").Value2 = puYear: pu.Range("D5").Value2 = puMonth
    CompactAssignBlock pu, "A10:G" & CStr(lastPeople), puData
    pu.Range("F5").Value2 = puSummary(0): pu.Range("H5").Value2 = puSummary(1)
    reportEnd = Application.Max(lastReport, CLng(Val(CC_Text(report.Range("N2").Value2))))
    If report.FilterMode Then report.ShowAllData
    If report.AutoFilterMode Then report.AutoFilterMode = False
    report.Range("A17:L" & CStr(reportEnd)).ClearContents
    report.Range("A17:L" & CStr(lastReport)).Value2 = reportData
    report.Range("N2:N5").Value2 = reportMeta
    If reportMode = "detail" Then report.Range("F18:G" & CStr(lastReport)).NumberFormat = "dd.mm.yyyy" Else report.Range("F18:G" & CStr(lastReport)).NumberFormat = "General"
    CY_RestoreCells report, reportFields
    CY_RestoreYearFilters report, reportFilterRange, reportFilters
    CY_RestoreYearFilters pu, puFilterRange, puFilters
    CompactPrimeRoster
    CompactHistoryInit
    CU_ShowSelected
Done:
    On Error Resume Next
    If Not protection Is Nothing Then
        For Each item In protection
            If CBool(item(1)) Then ThisWorkbook.Worksheets(CStr(item(0))).Protect Password:="2022-tiski", UserInterfaceOnly:=True, AllowFiltering:=True
        Next item
    End If
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    CompactOpenNewYear = failure
End Function

Private Function CY_CaptureCells(ByVal ws As Worksheet, ByVal addresses As Variant) As Collection
    Dim result As Collection, address As Variant
    Set result = New Collection
    For Each address In addresses: result.Add Array(CStr(address), ws.Range(CStr(address)).Value2): Next address
    Set CY_CaptureCells = result
End Function

Private Sub CY_RestoreCells(ByVal ws As Worksheet, ByVal values As Collection)
    Dim item As Variant
    For Each item In values: ws.Range(CStr(item(0))).Value2 = item(1): Next item
End Sub

Private Function CY_YearFilters(ByVal ws As Worksheet, ByRef filterRange As String) As Collection
    Dim result As Collection, i As Long, first As Variant, second As Variant, operation As Long, hasFirst As Boolean, hasSecond As Boolean
    Set result = New Collection
    If ws.AutoFilterMode Then
        filterRange = ws.AutoFilter.Range.address
        For i = 1 To ws.AutoFilter.filters.count
            If ws.AutoFilter.filters(i).On Then
                first = Empty: second = Empty: operation = 0
                On Error Resume Next
                Err.Clear: first = ws.AutoFilter.filters(i).Criteria1: hasFirst = Err.Number = 0
                Err.Clear: second = ws.AutoFilter.filters(i).Criteria2: hasSecond = Err.Number = 0
                Err.Clear: operation = ws.AutoFilter.filters(i).Operator
                On Error GoTo 0
                result.Add Array(i, operation, hasFirst, first, hasSecond, second)
            End If
        Next i
    End If
    Set CY_YearFilters = result
End Function

Private Sub CY_RestoreYearFilters(ByVal ws As Worksheet, ByVal filterRange As String, ByVal filters As Collection)
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If Len(filterRange) > 0 Then
        ws.Range(filterRange).AutoFilter
        CB_RestoreFilters ws, filters
    End If
End Sub

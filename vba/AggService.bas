Attribute VB_Name = "AggService"
'===============================================================================
' AggService - validation and orchestration between the sheets and Access
' (Universe: server/src/services + validation/*). Rules:
'   * Access is the system of record; the Aggregates sheet is a working copy.
'   * Push is all-or-nothing: one transaction, rolled back on the first failure
'     (Universe: db/withTransaction.js).
'   * Updates and deletes are optimistic-locked on row_version: if someone else
'     saved the row since your pull, your push is refused for that row.
'   * Every write is recorded in audit_log.
'===============================================================================
Option Explicit

Private Const A_NONE As Long = 0
Private Const A_INS As Long = 1
Private Const A_UPD As Long = 2
Private Const A_DEL As Long = 3
Private Const A_SAME As Long = 4

' ============================== small helpers =================================
Private Function Sh(ByVal nm As String) As Worksheet
    Set Sh = ThisWorkbook.Worksheets(nm)
End Function

Public Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function

Private Function IsBlankV(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    IsBlankV = (Len(Txt(v)) = 0)
End Function

Private Function NullIfBlank(ByVal v As Variant) As Variant
    If Len(Txt(v)) = 0 Then NullIfBlank = Null Else NullIfBlank = Txt(v)
End Function

Private Function NumText(ByVal v As Variant) As String
    If IsBlankV(v) Then
        NumText = "0"
    ElseIf IsNumeric(v) And VarType(v) <> vbBoolean Then
        NumText = Trim$(Str$(CDbl(v)))
    Else
        NumText = Txt(v)
    End If
End Function

Private Function GridValues(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal nCols As Long) As Variant
    GridValues = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastRow, nCols)).Value
End Function

Private Function SheetRow(ByVal i As Long) As Long
    SheetRow = FIRST_ROW + i - 1
End Function

' ============================ aggregate fingerprint ============================
' Normalised text of the editable columns. Pull stores it in the hidden column T;
' push compares it to decide whether a row changed.
Public Function Fingerprint(ByVal vals As Variant, ByVal i As Long) As String
    Dim parts(COL_CEDANT To COL_NOTES) As String, c As Long, v As Variant
    For c = COL_CEDANT To COL_NOTES
        v = vals(i, c)
        Select Case c
            Case COL_ZONE, COL_TREATY, COL_COB
                parts(c) = Txt(v)
                If Len(parts(c)) = 0 Then parts(c) = UNSPECIFIED
            Case COL_UWYEAR
                If IsNumeric(v) And Not IsBlankV(v) Then parts(c) = CStr(CLng(v)) Else parts(c) = Txt(v)
            Case COL_CCY
                parts(c) = UCase$(Txt(v))
            Case COL_ASAT
                If IsDate(v) And Not IsBlankV(v) Then parts(c) = Format$(CDate(v), "yyyy-mm-dd") Else parts(c) = Txt(v)
            Case COL_EQ, COL_WS, COL_FLOOD, COL_SRCC, COL_OTHERS
                parts(c) = NumText(v)
            Case Else
                parts(c) = Txt(v)
        End Select
    Next c
    Fingerprint = Join(parts, "|")
End Function

Private Function RowHasContent(ByVal vals As Variant, ByVal i As Long) As Boolean
    Dim c As Long
    For c = COL_ID To COL_DELETE
        If IsError(vals(i, c)) Then RowHasContent = True: Exit Function
        If Not IsBlankV(vals(i, c)) Then RowHasContent = True: Exit Function
    Next c
End Function

' ============================== validation ====================================
' Validates one Aggregates row and builds the record array for AggRepo.
' Returns "" when valid, otherwise the error text.
Private Function ValidateAggRow(ByVal vals As Variant, ByVal i As Long, ByVal ced As Object, _
        ByVal ctry As Object, ByVal tt As Object, ByVal cob As Object, ByVal ccy As Object, _
        ByRef rec As Variant) As String
    Dim c As Long, s As String, v As Variant, amt As Variant, cols As Variant, k As Long
    Dim r(0 To 14) As Variant

    For c = COL_CEDANT To COL_NOTES
        If IsError(vals(i, c)) Then
            ValidateAggRow = "column " & ColLetter(c) & " contains an error value": Exit Function
        End If
    Next c

    s = Txt(vals(i, COL_CEDANT))
    If Len(s) = 0 Then ValidateAggRow = "Cedant is required": Exit Function
    If Not ced.Exists(s) Then ValidateAggRow = "Cedant '" & s & "' is not in the database (add it on Cedants, then Push masters)": Exit Function
    r(AggRepo.R_CEDANT) = ced(s)

    s = Txt(vals(i, COL_COUNTRY))
    If Len(s) = 0 Then ValidateAggRow = "Country is required": Exit Function
    If Not ctry.Exists(s) Then ValidateAggRow = "Country '" & s & "' is not in the database": Exit Function
    r(AggRepo.R_COUNTRY) = ctry(s)

    s = Txt(vals(i, COL_ZONE))
    If Len(s) = 0 Then s = UNSPECIFIED
    If Len(s) > 64 Then ValidateAggRow = "Zone ID is longer than 64 characters": Exit Function
    r(AggRepo.R_ZONE) = s
    s = Txt(vals(i, COL_ZONENAME))
    If Len(s) > 150 Then ValidateAggRow = "Zone name is longer than 150 characters": Exit Function
    r(AggRepo.R_ZONENAME) = NullIfBlank(s)

    v = vals(i, COL_UWYEAR)
    If IsBlankV(v) Then ValidateAggRow = "UW year is required": Exit Function
    If Not IsNumeric(v) Then ValidateAggRow = "UW year must be a year, e.g. 2026": Exit Function
    If CDbl(v) <> Int(CDbl(v)) Or CDbl(v) < MIN_UW_YEAR Or CDbl(v) > MAX_UW_YEAR Then
        ValidateAggRow = "UW year must be a whole year between " & MIN_UW_YEAR & " and " & MAX_UW_YEAR: Exit Function
    End If
    r(AggRepo.R_UWYEAR) = CLng(v)

    s = Txt(vals(i, COL_TREATY))
    If Len(s) = 0 Then s = UNSPECIFIED
    If Not tt.Exists(s) Then ValidateAggRow = "Treaty type '" & s & "' is not in the database": Exit Function
    r(AggRepo.R_TREATY) = tt(s)

    s = Txt(vals(i, COL_COB))
    If Len(s) = 0 Then s = UNSPECIFIED
    If Not cob.Exists(s) Then ValidateAggRow = "Class of business '" & s & "' is not in the database": Exit Function
    r(AggRepo.R_COB) = cob(s)

    s = UCase$(Txt(vals(i, COL_CCY)))
    If Len(s) = 0 Then ValidateAggRow = "Currency is required": Exit Function
    If Not ccy.Exists(s) Then ValidateAggRow = "Currency '" & s & "' is not in the database (add it on Rates)": Exit Function
    r(AggRepo.R_CCY) = s

    v = vals(i, COL_ASAT)
    If IsBlankV(v) Then
        r(AggRepo.R_ASAT) = Null
    ElseIf IsDate(v) Then
        r(AggRepo.R_ASAT) = CDate(v)
    Else
        ValidateAggRow = "As-at date is not a valid date": Exit Function
    End If

    cols = Array(COL_EQ, COL_WS, COL_FLOOD, COL_SRCC, COL_OTHERS)
    For k = 0 To 4
        amt = vals(i, cols(k))
        If IsBlankV(amt) Then
            r(AggRepo.R_EQ + k) = 0#
        ElseIf IsNumeric(amt) And VarType(amt) <> vbBoolean Then
            If CDbl(amt) < 0 Then ValidateAggRow = PerilName(k) & " cannot be negative": Exit Function
            r(AggRepo.R_EQ + k) = CDbl(amt)
        Else
            ValidateAggRow = PerilName(k) & " must be a number": Exit Function
        End If
    Next k

    s = Txt(vals(i, COL_NOTES))
    r(AggRepo.R_NOTES) = NullIfBlank(Left$(s, 255))
    rec = r
End Function

Private Function PerilName(ByVal k As Long) As String
    PerilName = Choose(k + 1, "EQ", "WS", "Flood", "SRCC", "Others")
End Function

Private Function ColLetter(ByVal c As Long) As String
    ColLetter = Split(Sh(SH_AGG).Cells(1, c).Address(True, False), "$")(0)
End Function

' Natural key (matches ux_cedant_aggregate_key; Access text compare is case-insensitive)
Private Function AggKey(ByVal rec As Variant) As String
    AggKey = rec(AggRepo.R_CEDANT) & "|" & rec(AggRepo.R_COUNTRY) & "|" & UCase$(rec(AggRepo.R_ZONE)) & "|" & _
             rec(AggRepo.R_UWYEAR) & "|" & rec(AggRepo.R_TREATY) & "|" & rec(AggRepo.R_COB) & "|" & rec(AggRepo.R_CCY)
End Function

' ================================ aggregates ===================================
Public Function PullAggregates() As String
    Dim cn As Object, rows As Variant, n As Long
    On Error GoTo EH
    Set cn = AggDb.Connect()
    rows = AggRepo.SelectAggregates(cn)
    AggDb.CloseQuietly cn
    n = WriteAggregates(rows)
    SetConfigValue "cfg_LastPull", Now
    PullAggregates = n & " aggregate row(s) loaded from Access."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    AggDb.CloseQuietly cn
    Err.Raise vbObjectError + 1300, APP_NAME, d
End Function

Private Function WriteAggregates(ByVal rows As Variant) As Long
    Dim ws As Worksheet, nr As Long, r As Long, c As Long, out() As Variant, vals As Variant, fp() As Variant
    Set ws = Sh(SH_AGG)
    ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(LAST_ROW_AGG, AGG_INPUT_COLS)).ClearContents
    If Not IsArray(rows) Then Exit Function
    nr = UBound(rows, 1)
    If nr > LAST_ROW_AGG - FIRST_ROW + 1 Then
        Err.Raise vbObjectError + 1301, APP_NAME, "Access holds " & nr & " aggregate rows but the sheet holds " & _
            (LAST_ROW_AGG - FIRST_ROW + 1) & ". Report from the Access views (vw_agg_*) or extend the sheet."
    End If
    ReDim out(1 To nr, 1 To AGG_INPUT_COLS)
    For r = 1 To nr
        out(r, COL_STATUS) = "OK"
        For c = 1 To 16                       ' agg_id .. agg_notes  ->  B .. Q
            out(r, COL_ID + c - 1) = rows(r, c)
        Next c
        out(r, COL_VERSION) = rows(r, 17)
    Next r
    ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + nr - 1, AGG_INPUT_COLS)).Value = out

    ' Fingerprint what the sheet now holds (after Excel's own type conversion)
    vals = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + nr - 1, AGG_INPUT_COLS)).Value
    ReDim fp(1 To nr, 1 To 1)
    For r = 1 To nr
        fp(r, 1) = Fingerprint(vals, r)
    Next r
    ws.Range(ws.Cells(FIRST_ROW, COL_FP), ws.Cells(FIRST_ROW + nr - 1, COL_FP)).Value = fp
    WriteAggregates = nr
End Function

Public Function PushAggregates() As String
    Dim ws As Worksheet, vals As Variant, n As Long, i As Long, pass As Long
    Dim cn As Object, inTx As Boolean, curRow As Long
    Dim ced As Object, ctry As Object, tt As Object, cob As Object, ccy As Object, keys As Object
    Dim stat() As Variant, act() As Long, recs() As Variant, rec As Variant
    Dim msg As String, k As String, errCount As Long, isDel As Boolean, hasId As Boolean
    Dim nIns As Long, nUpd As Long, nDel As Long, nSame As Long, newId As Long
    Dim aggId As Long, ver As Long

    Set ws = Sh(SH_AGG)
    n = LAST_ROW_AGG - FIRST_ROW + 1
    vals = GridValues(ws, LAST_ROW_AGG, AGG_INPUT_COLS)
    ReDim stat(1 To n, 1 To 1)
    ReDim act(1 To n)
    ReDim recs(1 To n)

    On Error GoTo EH
    Set cn = AggDb.Connect()
    Set ced = AggRepo.CedantMap(cn)
    Set ctry = AggRepo.CountryMap(cn)
    Set tt = AggRepo.TreatyTypeMap(cn)
    Set cob = AggRepo.CobMap(cn)
    Set ccy = AggRepo.CurrencyMap(cn)
    Set keys = CreateObject("Scripting.Dictionary")

    ' ---- 1. validate every row, decide the action -----------------------------
    For i = 1 To n
        act(i) = A_NONE
        If RowHasContent(vals, i) Then
            isDel = (UCase$(Txt(vals(i, COL_DELETE))) = "X")
            hasId = Not IsBlankV(vals(i, COL_ID))
            If hasId Then
                If Not IsNumeric(vals(i, COL_ID)) Or Not IsNumeric(vals(i, COL_VERSION)) Or IsBlankV(vals(i, COL_VERSION)) Then
                    stat(i, 1) = "ERROR: ID / version (columns B, S) were edited - Pull again"
                    errCount = errCount + 1
                    GoTo NextRow
                End If
            End If
            If isDel Then
                If hasId Then
                    act(i) = A_DEL
                Else
                    stat(i, 1) = "SKIPPED: marked for delete but never saved - clear the row"
                End If
                GoTo NextRow
            End If
            msg = ValidateAggRow(vals, i, ced, ctry, tt, cob, ccy, rec)
            If Len(msg) > 0 Then
                stat(i, 1) = "ERROR: " & msg
                errCount = errCount + 1
                GoTo NextRow
            End If
            k = AggKey(rec)
            If keys.Exists(k) Then
                stat(i, 1) = "ERROR: same cedant/country/zone/year/treaty/class/currency as row " & keys(k)
                errCount = errCount + 1
                GoTo NextRow
            End If
            keys.Add k, SheetRow(i)
            recs(i) = rec
            If Not hasId Then
                act(i) = A_INS
            ElseIf Fingerprint(vals, i) <> Txt(vals(i, COL_FP)) Then
                act(i) = A_UPD
            Else
                act(i) = A_SAME
                stat(i, 1) = "OK"
            End If
        End If
NextRow:
    Next i

    If errCount > 0 Then
        For i = 1 To n
            If act(i) = A_INS Or act(i) = A_UPD Or act(i) = A_DEL Then stat(i, 1) = "NOT SAVED - fix the errors first"
        Next i
        AggDb.CloseQuietly cn
        WriteStatus ws, stat
        PushAggregates = errCount & " row(s) have errors - nothing was saved." & vbCrLf & _
            "Filter column A (Status) on 'ERROR' to see them."
        Exit Function
    End If

    ' ---- 2. write: deletes, then updates, then inserts, in one transaction -----
    cn.BeginTrans
    inTx = True
    For pass = 1 To 3
        For i = 1 To n
            curRow = i
            Select Case act(i)
                Case A_DEL
                    If pass = 1 Then
                        aggId = CLng(vals(i, COL_ID)): ver = CLng(vals(i, COL_VERSION))
                        If AggRepo.DeleteAggregate(cn, aggId, ver) = 0 Then
                            Err.Raise vbObjectError + 1310, APP_NAME, "CONFLICT: changed or deleted in Access since your last pull - Pull again"
                        End If
                        AggRepo.WriteAudit cn, "cedant_aggregate", aggId, "AGG_DELETED", Txt(vals(i, COL_FP))
                        nDel = nDel + 1
                    End If
                Case A_UPD
                    If pass = 2 Then
                        aggId = CLng(vals(i, COL_ID)): ver = CLng(vals(i, COL_VERSION))
                        If AggRepo.UpdateAggregate(cn, aggId, ver, recs(i)) = 0 Then
                            Err.Raise vbObjectError + 1311, APP_NAME, "CONFLICT: changed or deleted in Access since your last pull - Pull again"
                        End If
                        AggRepo.WriteAudit cn, "cedant_aggregate", aggId, "AGG_UPDATED", Fingerprint(vals, i), _
                            "was: " & Txt(vals(i, COL_FP))
                        nUpd = nUpd + 1
                    End If
                Case A_INS
                    If pass = 3 Then
                        newId = AggRepo.InsertAggregate(cn, recs(i))
                        AggRepo.WriteAudit cn, "cedant_aggregate", newId, "AGG_CREATED", Fingerprint(vals, i)
                        nIns = nIns + 1
                    End If
                Case A_SAME
                    If pass = 1 Then nSame = nSame + 1
            End Select
        Next i
    Next pass
    curRow = 0
    cn.CommitTrans
    inTx = False
    AggDb.CloseQuietly cn

    ' ---- 3. reload so the sheet carries fresh IDs, versions and fingerprints ----
    PullAggregates
    SetConfigValue "cfg_LastPush", Now
    PushAggregates = "Saved to Access: " & nIns & " added, " & nUpd & " updated, " & nDel & " deleted, " & _
        nSame & " unchanged."
    Exit Function

EH:
    Dim d As String: d = Err.Description
    If inTx Then AggDb.RollbackQuietly cn
    AggDb.CloseQuietly cn
    If curRow > 0 Then
        For i = 1 To n
            If act(i) = A_INS Or act(i) = A_UPD Or act(i) = A_DEL Then stat(i, 1) = "NOT SAVED (rolled back)"
        Next i
        stat(curRow, 1) = "ERROR: " & d
        WriteStatus ws, stat
        Err.Raise vbObjectError + 1312, APP_NAME, "Nothing was saved. Row " & SheetRow(curRow) & ": " & d
    End If
    Err.Raise vbObjectError + 1313, APP_NAME, d
End Function

Private Sub WriteStatus(ByVal ws As Worksheet, ByVal stat As Variant)
    ws.Range(ws.Cells(FIRST_ROW, COL_STATUS), ws.Cells(LAST_ROW_AGG, COL_STATUS)).Value = stat
End Sub

' ================================ master data ==================================
' Push order: countries -> currencies/rates -> treaty types/classes -> cedants.
' Validation runs over every sheet first; if anything fails nothing is written.
Public Function PushMasters() As String
    Dim cn As Object, inTx As Boolean, errs As String
    Dim vC As Variant, vR As Variant, vL As Variant, vK As Variant
    Dim sC() As Variant, sR() As Variant, sK() As Variant
    Dim nCtry As Long, nCcy As Long, nRate As Long, nList As Long, nCed As Long

    vC = GridValues(Sh(SH_COUNTRIES), LAST_ROW_CTRY, CTRY_COLS)
    vR = GridValues(Sh(SH_RATES), LAST_ROW_RATE, RATE_COLS)
    vL = GridValues(Sh(SH_LISTS), LAST_ROW_LIST, 5)
    vK = GridValues(Sh(SH_CEDANTS), LAST_ROW_CED, CED_COLS)
    ReDim sC(1 To UBound(vC, 1), 1 To 1)
    ReDim sR(1 To UBound(vR, 1), 1 To 1)
    ReDim sK(1 To UBound(vK, 1), 1 To 1)

    errs = ValidateCountries(vC, sC) & ValidateRates(vR, sR) & ValidateLists(vL) & ValidateCedants(vK, vC, sK)
    If Len(errs) > 0 Then
        WriteCol Sh(SH_COUNTRIES), 6, sC
        WriteCol Sh(SH_RATES), 5, sR
        WriteCol Sh(SH_CEDANTS), 5, sK
        PushMasters = "Nothing was saved - fix these first:" & vbCrLf & Left$(errs, 1500)
        Exit Function
    End If

    On Error GoTo EH
    Set cn = AggDb.Connect()
    cn.BeginTrans
    inTx = True
    nCtry = WriteCountries(cn, vC)
    WriteRates cn, vR, nCcy, nRate
    nList = WriteLists(cn, vL)
    nCed = WriteCedants(cn, vK)
    cn.CommitTrans
    inTx = False
    AggDb.CloseQuietly cn
    PullMasters
    PushMasters = "Saved to Access: " & nCtry & " countries, " & nCcy & " currencies, " & nRate & _
        " FX rate change(s), " & nList & " treaty types/classes, " & nCed & " cedants."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    If inTx Then AggDb.RollbackQuietly cn
    AggDb.CloseQuietly cn
    Err.Raise vbObjectError + 1400, APP_NAME, "Nothing was saved: " & d
End Function

Private Sub WriteCol(ByVal ws As Worksheet, ByVal col As Long, ByVal arr As Variant)
    ws.Range(ws.Cells(FIRST_ROW, col), ws.Cells(FIRST_ROW + UBound(arr, 1) - 1, col)).Value = arr
End Sub

Private Function RowBlank(ByVal v As Variant, ByVal i As Long, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim c As Long
    For c = c1 To c2
        If IsError(v(i, c)) Then Exit Function
        If Not IsBlankV(v(i, c)) Then Exit Function
    Next c
    RowBlank = True
End Function

Private Function ErrLine(ByVal sheetName As String, ByVal i As Long, ByVal m As String) As String
    ErrLine = "  " & sheetName & " row " & SheetRow(i) & ": " & m & vbCrLf
End Function

' Countries: A name | B ISO2 | C ISO3 | D region | E id | F status
Private Function ValidateCountries(ByVal v As Variant, ByRef st() As Variant) As String
    Dim i As Long, e As String, m As String, seenC As Object, seenN As Object, cd As String, nm As String
    Set seenC = CreateObject("Scripting.Dictionary"): seenC.CompareMode = vbTextCompare
    Set seenN = CreateObject("Scripting.Dictionary"): seenN.CompareMode = vbTextCompare
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            nm = Txt(v(i, 1)): cd = UCase$(Txt(v(i, 2))): m = ""
            If Len(nm) = 0 Then
                m = "country name is required"
            ElseIf Len(cd) <> 2 Then
                m = "ISO2 code must be 2 letters"
            ElseIf Len(Txt(v(i, 3))) > 0 And Len(Txt(v(i, 3))) <> 3 Then
                m = "ISO3 code must be 3 letters"
            ElseIf seenC.Exists(cd) Then
                m = "ISO2 " & cd & " repeats row " & seenC(cd)
            ElseIf seenN.Exists(nm) Then
                m = "name repeats row " & seenN(nm)
            End If
            If Len(m) > 0 Then
                st(i, 1) = "ERROR: " & m: e = e & ErrLine(SH_COUNTRIES, i, m)
            Else
                seenC(cd) = SheetRow(i): seenN(nm) = SheetRow(i): st(i, 1) = v(i, 6)
            End If
        End If
    Next i
    ValidateCountries = e
End Function

Private Function WriteCountries(ByVal cn As Object, ByVal v As Variant) As Long
    Dim i As Long, byCode As Object, byName As Object, cd As String, nm As String, id As Variant
    Set byCode = AggRepo.CountryCodeMap(cn)
    Set byName = AggRepo.CountryMap(cn)
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            nm = Txt(v(i, 1)): cd = UCase$(Txt(v(i, 2))): id = Empty
            If byCode.Exists(cd) Then
                id = byCode(cd)
            ElseIf byName.Exists(nm) Then
                id = byName(nm)
            End If
            AggRepo.UpsertCountry cn, id, nm, cd, NullIfBlank(UCase$(Txt(v(i, 3)))), NullIfBlank(v(i, 4))
            WriteCountries = WriteCountries + 1
        End If
    Next i
End Function

' Rates: A code | B name | C USD per 1 unit | D effective date | E status
Private Function ValidateRates(ByVal v As Variant, ByRef st() As Variant) As String
    Dim i As Long, e As String, m As String, seen As Object, cd As String
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            cd = UCase$(Txt(v(i, 1))): m = ""
            If IsError(v(i, 3)) Or IsError(v(i, 4)) Then
                m = "rate or date cell contains an error value"
            ElseIf Len(cd) <> 3 Then
                m = "currency code must be 3 letters"
            ElseIf seen.Exists(cd) Then
                m = cd & " repeats row " & seen(cd)
            ElseIf Not IsBlankV(v(i, 3)) Then
                If Not IsNumeric(v(i, 3)) Then
                    m = "rate must be a number (USD per 1 " & cd & ")"
                ElseIf CDbl(v(i, 3)) <= 0 Then
                    m = "rate must be greater than zero"
                ElseIf IsBlankV(v(i, 4)) Or Not IsDate(v(i, 4)) Then
                    m = "effective date is required when a rate is entered"
                End If
            End If
            If Len(m) > 0 Then
                st(i, 1) = "ERROR: " & m: e = e & ErrLine(SH_RATES, i, m)
            Else
                seen(cd) = SheetRow(i): st(i, 1) = v(i, 5)
            End If
        End If
    Next i
    ValidateRates = e
End Function

Private Sub WriteRates(ByVal cn As Object, ByVal v As Variant, ByRef nCcy As Long, ByRef nRate As Long)
    Dim i As Long, have As Object, cd As String, nm As String
    Set have = AggRepo.CurrencyMap(cn)
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            cd = UCase$(Txt(v(i, 1))): nm = Txt(v(i, 2))
            If Len(nm) = 0 Then nm = cd
            AggRepo.UpsertCurrency cn, have.Exists(cd), cd, Left$(nm, 60)
            nCcy = nCcy + 1
            If Not IsBlankV(v(i, 3)) Then
                If AggRepo.UpsertRate(cn, cd, CDbl(v(i, 3)), DateValue(CDate(v(i, 4)))) Then nRate = nRate + 1
            End If
        End If
    Next i
End Sub

' Lists: A treaty type | B category | D class of business | E code
Private Function ValidateLists(ByVal v As Variant) As String
    Dim i As Long, e As String, seenT As Object, seenB As Object, nm As String
    Set seenT = CreateObject("Scripting.Dictionary"): seenT.CompareMode = vbTextCompare
    Set seenB = CreateObject("Scripting.Dictionary"): seenB.CompareMode = vbTextCompare
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 2) Then
            nm = Txt(v(i, 1))
            If Len(nm) = 0 Then
                e = e & ErrLine(SH_LISTS, i, "treaty type name is required")
            ElseIf seenT.Exists(nm) Then
                e = e & ErrLine(SH_LISTS, i, "treaty type '" & nm & "' repeats")
            Else
                seenT(nm) = True
            End If
        End If
        If Not RowBlank(v, i, 4, 5) Then
            nm = Txt(v(i, 4))
            If Len(nm) = 0 Then
                e = e & ErrLine(SH_LISTS, i, "class of business name is required")
            ElseIf seenB.Exists(nm) Then
                e = e & ErrLine(SH_LISTS, i, "class of business '" & nm & "' repeats")
            Else
                seenB(nm) = True
            End If
        End If
    Next i
    ValidateLists = e
End Function

Private Function WriteLists(ByVal cn As Object, ByVal v As Variant) As Long
    Dim i As Long, tt As Object, cob As Object, nm As String
    Set tt = AggRepo.TreatyTypeMap(cn)
    Set cob = AggRepo.CobMap(cn)
    For i = 1 To UBound(v, 1)
        nm = Txt(v(i, 1))
        If Len(nm) > 0 Then
            AggRepo.UpsertTreatyType cn, tt.Exists(nm), Left$(nm, 80), NullIfBlank(v(i, 2))
            WriteLists = WriteLists + 1
        End If
        nm = Txt(v(i, 4))
        If Len(nm) > 0 Then
            AggRepo.UpsertCob cn, cob.Exists(nm), Left$(nm, 80), NullIfBlank(v(i, 5))
            WriteLists = WriteLists + 1
        End If
    Next i
End Function

' Cedants: A name | B home country | C active Y/N | D id | E status
Private Function ValidateCedants(ByVal v As Variant, ByVal vCountries As Variant, ByRef st() As Variant) As String
    Dim i As Long, e As String, m As String, seen As Object, ctry As Object, nm As String, hc As String, act As String
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    Set ctry = CreateObject("Scripting.Dictionary"): ctry.CompareMode = vbTextCompare
    For i = 1 To UBound(vCountries, 1)
        If Len(Txt(vCountries(i, 1))) > 0 Then ctry(Txt(vCountries(i, 1))) = True
    Next i
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            nm = Txt(v(i, 1)): hc = Txt(v(i, 2)): act = UCase$(Txt(v(i, 3))): m = ""
            If Len(nm) = 0 Then
                m = "cedant name is required"
            ElseIf Len(nm) > 150 Then
                m = "cedant name is longer than 150 characters"
            ElseIf seen.Exists(nm) Then
                m = "name repeats row " & seen(nm)
            ElseIf Len(hc) > 0 And Not ctry.Exists(hc) Then
                m = "home country '" & hc & "' is not on the Countries sheet"
            ElseIf Len(act) > 0 And act <> "Y" And act <> "N" Then
                m = "Active must be Y or N"
            ElseIf Not IsBlankV(v(i, 4)) And Not IsNumeric(v(i, 4)) Then
                m = "ID column was edited - Pull masters again"
            End If
            If Len(m) > 0 Then
                st(i, 1) = "ERROR: " & m: e = e & ErrLine(SH_CEDANTS, i, m)
            Else
                seen(nm) = SheetRow(i): st(i, 1) = v(i, 5)
            End If
        End If
    Next i
    ValidateCedants = e
End Function

Private Function WriteCedants(ByVal cn As Object, ByVal v As Variant) As Long
    Dim i As Long, byName As Object, ctry As Object, nm As String, hc As Variant, active As Boolean
    Dim newId As Long, cid As Long
    Set byName = AggRepo.CedantMap(cn)
    Set ctry = AggRepo.CountryMap(cn)          ' includes countries written earlier in this transaction
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            nm = Txt(v(i, 1))
            hc = Null
            If Len(Txt(v(i, 2))) > 0 Then hc = ctry(Txt(v(i, 2)))
            active = (UCase$(Txt(v(i, 3))) <> "N")
            If Not IsBlankV(v(i, 4)) Then
                cid = CLng(v(i, 4))
                If AggRepo.UpdateCedant(cn, cid, nm, hc, active) = 0 Then
                    Err.Raise vbObjectError + 1410, APP_NAME, "Cedants row " & SheetRow(i) & ": ID " & cid & _
                        " is not in the database - Pull masters again"
                End If
                AggRepo.WriteAudit cn, "cedant", cid, "CEDANT_SAVED", nm
            ElseIf byName.Exists(nm) Then
                cid = CLng(byName(nm))
                AggRepo.UpdateCedant cn, cid, nm, hc, active
                AggRepo.WriteAudit cn, "cedant", cid, "CEDANT_SAVED", nm
            Else
                newId = AggRepo.InsertCedant(cn, nm, hc, active)
                AggRepo.WriteAudit cn, "cedant", newId, "CEDANT_CREATED", nm
            End If
            WriteCedants = WriteCedants + 1
        End If
    Next i
End Function

Public Function PullMasters() As String
    Dim cn As Object, rows As Variant, out() As Variant, r As Long, nK As Long
    On Error GoTo EH
    Set cn = AggDb.Connect()

    ' Countries
    rows = AggRepo.SelectCountries(cn)
    ClearGrid Sh(SH_COUNTRIES), LAST_ROW_CTRY, 1, CTRY_COLS
    If IsArray(rows) Then
        ReDim out(1 To UBound(rows, 1), 1 To CTRY_COLS)
        For r = 1 To UBound(rows, 1)
            out(r, 1) = rows(r, 1): out(r, 2) = rows(r, 2): out(r, 3) = rows(r, 3)
            out(r, 4) = rows(r, 4): out(r, 5) = rows(r, 5): out(r, 6) = "OK"
        Next r
        PutGrid Sh(SH_COUNTRIES), out
    End If

    ' Currencies with their latest rate
    rows = AggRepo.SelectRates(cn)
    ClearGrid Sh(SH_RATES), LAST_ROW_RATE, 1, RATE_COLS
    If IsArray(rows) Then
        ReDim out(1 To UBound(rows, 1), 1 To RATE_COLS)
        For r = 1 To UBound(rows, 1)
            out(r, 1) = rows(r, 1): out(r, 2) = rows(r, 2): out(r, 3) = rows(r, 3)
            out(r, 4) = rows(r, 4)
            If IsEmpty(rows(r, 3)) Then out(r, 5) = "NO RATE - USD totals exclude this currency" Else out(r, 5) = "OK"
        Next r
        PutGrid Sh(SH_RATES), out
    End If

    ' Treaty types (A:B) and classes (D:E)
    ClearGrid Sh(SH_LISTS), LAST_ROW_LIST, 1, 2
    ClearGrid Sh(SH_LISTS), LAST_ROW_LIST, 4, 5
    rows = AggRepo.SelectTreatyTypes(cn)
    If IsArray(rows) Then PutGrid Sh(SH_LISTS), rows, 1
    rows = AggRepo.SelectCobs(cn)
    If IsArray(rows) Then PutGrid Sh(SH_LISTS), rows, 4

    ' Cedants
    rows = AggRepo.SelectCedants(cn)
    ClearGrid Sh(SH_CEDANTS), LAST_ROW_CED, 1, CED_COLS
    If IsArray(rows) Then
        nK = UBound(rows, 1)
        ReDim out(1 To nK, 1 To CED_COLS)
        For r = 1 To nK
            out(r, 1) = rows(r, 1): out(r, 2) = rows(r, 2)
            If CBool(rows(r, 3)) Then out(r, 3) = "Y" Else out(r, 3) = "N"
            out(r, 4) = rows(r, 4): out(r, 5) = "OK"
        Next r
        PutGrid Sh(SH_CEDANTS), out
    End If

    AggDb.CloseQuietly cn
    SetConfigValue "cfg_LastPull", Now
    PullMasters = nK & " cedant(s) and reference lists loaded from Access."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    AggDb.CloseQuietly cn
    Err.Raise vbObjectError + 1420, APP_NAME, d
End Function

Private Sub ClearGrid(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal c1 As Long, ByVal c2 As Long)
    ws.Range(ws.Cells(FIRST_ROW, c1), ws.Cells(lastRow, c2)).ClearContents
End Sub

Private Sub PutGrid(ByVal ws As Worksheet, ByVal arr As Variant, Optional ByVal firstCol As Long = 1)
    ws.Range(ws.Cells(FIRST_ROW, firstCol), _
             ws.Cells(FIRST_ROW + UBound(arr, 1) - 1, firstCol + UBound(arr, 2) - 1)).Value = arr
End Sub

' ================================ add cedant ===================================
' Adds one cedant straight to Access and appends it to the Cedants sheet, so it
' is in the dropdowns at once. Unsaved edits elsewhere on the sheet are kept.
Public Function AddCedant(ByVal nm As String, ByVal countryText As String) As String
    Dim cn As Object, inTx As Boolean, hc As Variant, hcName As Variant, newId As Long
    Dim byName As Object, byCode As Object, ws As Worksheet, r As Long, freeRow As Long

    nm = Trim$(nm): countryText = Trim$(countryText)
    If Len(nm) = 0 Then Err.Raise vbObjectError + 1700, APP_NAME, "Cedant name is required."
    If Len(nm) > 150 Then Err.Raise vbObjectError + 1701, APP_NAME, "Cedant name is longer than 150 characters."

    Set ws = Sh(SH_CEDANTS)
    For r = FIRST_ROW To LAST_ROW_CED
        If Len(Txt(ws.Cells(r, 1).Value)) = 0 And Len(Txt(ws.Cells(r, 4).Value)) = 0 Then freeRow = r: Exit For
    Next r
    If freeRow = 0 Then Err.Raise vbObjectError + 1702, APP_NAME, "The Cedants sheet is full (" & _
        (LAST_ROW_CED - FIRST_ROW + 1) & " rows)."

    On Error GoTo EH
    Set cn = AggDb.Connect()
    If AggRepo.CedantMap(cn).Exists(nm) Then
        Err.Raise vbObjectError + 1703, APP_NAME, "A cedant called '" & nm & "' already exists."
    End If
    hc = Null: hcName = Empty
    If Len(countryText) > 0 Then
        Set byName = AggRepo.CountryMap(cn)
        Set byCode = AggRepo.CountryCodeMap(cn)
        If byName.Exists(countryText) Then
            hc = byName(countryText)
        ElseIf byCode.Exists(countryText) Then
            hc = byCode(countryText)
        Else
            Err.Raise vbObjectError + 1704, APP_NAME, "Country '" & countryText & "' is not in the database." & _
                vbCrLf & "Use a name or ISO2 code from the Countries sheet, or add the country there and Push masters."
        End If
        hcName = AggDb.Scalar(cn, "SELECT country_name FROM ref_country WHERE country_id = ?", CLng(hc))
    End If
    cn.BeginTrans
    inTx = True
    newId = AggRepo.InsertCedant(cn, nm, hc, True)
    AggRepo.WriteAudit cn, "cedant", newId, "CEDANT_CREATED", nm
    cn.CommitTrans
    inTx = False
    AggDb.CloseQuietly cn

    ws.Cells(freeRow, 1).Value = nm
    If Not IsEmpty(hcName) And Not IsNull(hcName) Then ws.Cells(freeRow, 2).Value = hcName
    ws.Cells(freeRow, 3).Value = "Y"
    ws.Cells(freeRow, 4).Value = newId
    ws.Cells(freeRow, 5).Value = "OK"
    AddCedant = "Added cedant '" & nm & "' (ID " & newId & ")." & vbCrLf & _
        "It is now in the Cedant dropdown on the Aggregates sheet."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    If inTx Then AggDb.RollbackQuietly cn
    AggDb.CloseQuietly cn
    Err.Raise vbObjectError + 1705, APP_NAME, d
End Function

' ================================== setup ======================================
' Create the .accdb if needed, bring the schema up to date and, for a brand-new
' database, seed it from the reference sheets shipped in this workbook.
Public Function SetupDatabase() As String
    Dim path As String, existed As Boolean, cn As Object, m As String, seedMsg As String
    path = DbPath()
    existed = (Len(Dir$(path)) > 0)
    If Not existed Then AggDb.CreateDatabaseFile path
    On Error GoTo EH
    Set cn = AggDb.Connect()
    m = AggMigrate.RunMigrations(cn)
    m = m & vbCrLf & AggSeed.EnsureReferenceData(cn)
    AggDb.CloseQuietly cn
    On Error GoTo 0
    If Not existed Then
        seedMsg = PushMasters()
        If Left$(seedMsg, 7) = "Nothing" Then
            SetupDatabase = "Database created at" & vbCrLf & path & vbCrLf & m & vbCrLf & vbCrLf & seedMsg
            Exit Function
        End If
        ' A new database has no aggregates yet: push whatever is on the sheet.
        seedMsg = seedMsg & vbCrLf & PushAggregates()
    Else
        PullMasters
        PullAggregates
        seedMsg = "Existing data loaded from Access."
    End If
    SetupDatabase = IIf(existed, "Connected to", "Created") & vbCrLf & path & vbCrLf & vbCrLf & m & vbCrLf & seedMsg
    Exit Function
EH:
    Dim d As String: d = Err.Description
    AggDb.CloseQuietly cn
    Err.Raise vbObjectError + 1500, APP_NAME, d
End Function

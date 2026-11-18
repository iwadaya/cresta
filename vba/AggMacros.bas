Attribute VB_Name = "AggMacros"
'===============================================================================
' AggMacros - the macros behind the buttons. Thin wrappers: switch Excel into a
' fast mode, call AggService, report the result. (Universe: client screens
' calling the API.)
'===============================================================================
Option Explicit

Private mCalc As Long

Private Sub BeginWork()
    mCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = -4135          ' xlCalculationManual
    Application.StatusBar = APP_NAME & ": working..."
End Sub

Private Sub EndWork()
    If mCalc = 0 Then mCalc = -4105          ' xlCalculationAutomatic
    Application.Calculation = mCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Private Sub Fail(ByVal what As String)
    Dim d As String: d = Err.Description
    EndWork
    MsgBox what & " failed:" & vbCrLf & vbCrLf & d, vbExclamation, APP_NAME
End Sub

Public Sub Setup_Database()
    Dim p As String, msg As String
    On Error GoTo EH
    p = DbPath()
    If Len(Dir$(p)) = 0 Then
        If MsgBox("Create a new Access database at:" & vbCrLf & p & vbCrLf & vbCrLf & _
                  "The reference lists, cedants and any aggregate rows in this workbook will be loaded into it.", _
                  vbOKCancel + vbQuestion, APP_NAME) <> vbOK Then Exit Sub
    End If
    BeginWork
    msg = AggService.SetupDatabase()
    EndWork
    MsgBox msg, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Setup"
End Sub

Public Sub Pull_All()
    Dim m1 As String, m2 As String
    On Error GoTo EH
    BeginWork
    m1 = AggService.PullMasters()
    m2 = AggService.PullAggregates()
    EndWork
    MsgBox m1 & vbCrLf & m2, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Pull"
End Sub

Public Sub Pull_Aggregates()
    Dim m As String
    On Error GoTo EH
    If HasPendingEdits() Then
        If MsgBox("The Aggregates sheet has rows that are not saved to Access." & vbCrLf & _
                  "Pulling will discard them. Continue?", vbOKCancel + vbExclamation, APP_NAME) <> vbOK Then Exit Sub
    End If
    BeginWork
    m = AggService.PullAggregates()
    EndWork
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Pull aggregates"
End Sub

Public Sub Push_Aggregates()
    Dim m As String
    On Error GoTo EH
    BeginWork
    m = AggService.PushAggregates()
    EndWork
    If InStr(m, "nothing was saved") > 0 Then
        MsgBox m, vbExclamation, APP_NAME
    Else
        MsgBox m, vbInformation, APP_NAME
    End If
    Exit Sub
EH:
    Fail "Push aggregates"
End Sub

Public Sub Push_Masters()
    Dim m As String
    On Error GoTo EH
    BeginWork
    m = AggService.PushMasters()
    EndWork
    If Left$(m, 7) = "Nothing" Then
        MsgBox m, vbExclamation, APP_NAME
    Else
        MsgBox m, vbInformation, APP_NAME
    End If
    Exit Sub
EH:
    Fail "Push masters"
End Sub

Public Sub Pull_Masters()
    Dim m As String
    On Error GoTo EH
    BeginWork
    m = AggService.PullMasters()
    EndWork
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Pull masters"
End Sub

Public Sub Add_Cedant()
    Dim nm As Variant, ctry As Variant, m As String
    On Error GoTo EH
    nm = Application.InputBox("Cedant name:", APP_NAME & " - add cedant", Type:=2)
    If VarType(nm) = vbBoolean Then Exit Sub            ' Cancel
    If Len(Trim$(CStr(nm))) = 0 Then Exit Sub
    ctry = Application.InputBox("Home country of '" & Trim$(CStr(nm)) & "'" & vbCrLf & _
        "Country name or ISO2 code, e.g. Zambia or ZM. Leave blank if none.", _
        APP_NAME & " - add cedant", Type:=2)
    If VarType(ctry) = vbBoolean Then Exit Sub
    BeginWork
    m = AggService.AddCedant(CStr(nm), CStr(ctry))
    EndWork
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Add cedant"
End Sub

Public Sub Open_Database()
    On Error GoTo EH
    Dim p As String: p = DbPath()
    If Len(Dir$(p)) = 0 Then
        MsgBox "No database yet at" & vbCrLf & p & vbCrLf & "Run Setup first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    ThisWorkbook.FollowHyperlink p
    Exit Sub
EH:
    Fail "Open database"
End Sub

' Rows added or edited since the last pull (unsaved work on the sheet).
Private Function HasPendingEdits() As Boolean
    Dim ws As Worksheet, vals As Variant, i As Long, hasData As Boolean, c As Long
    Set ws = ThisWorkbook.Worksheets(SH_AGG)
    vals = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(LAST_ROW_AGG, AGG_INPUT_COLS)).Value
    For i = 1 To UBound(vals, 1)
        hasData = False
        For c = COL_CEDANT To COL_DELETE
            If Not IsError(vals(i, c)) Then
                If Len(AggService.Txt(vals(i, c))) > 0 Then hasData = True: Exit For
            Else
                hasData = True: Exit For
            End If
        Next c
        If hasData Then
            If Len(AggService.Txt(vals(i, COL_ID))) = 0 Then HasPendingEdits = True: Exit Function
            If UCase$(AggService.Txt(vals(i, COL_DELETE))) = "X" Then HasPendingEdits = True: Exit Function
            If AggService.Fingerprint(vals, i) <> AggService.Txt(vals(i, COL_FP)) Then HasPendingEdits = True: Exit Function
        End If
    Next i
End Function

' ------------------------------------------------------------------------------
' Adds the buttons to the sheets. Run once after importing the modules
' (Build-Workbook.ps1 runs it for you). Safe to run again.
' ------------------------------------------------------------------------------
Public Sub InstallButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_START)
    ClearButtons ws
    AddButton ws, "C13", "1. Setup / create database", "Setup_Database"
    AddButton ws, "C15", "Pull everything from Access", "Pull_All"
    AddButton ws, "C17", "Push reference lists + cedants", "Push_Masters"
    AddButton ws, "C19", "Push aggregates", "Push_Aggregates"
    AddButton ws, "C21", "Open database in Access", "Open_Database"
    AddButton ws, "C23", "Add a cedant", "Add_Cedant", 1

    Set ws = ThisWorkbook.Worksheets(SH_AGG)
    ClearButtons ws
    AddButton ws, "L1", "Push to Access", "Push_Aggregates", 1, 140
    AddButton ws, "O1", "Pull from Access", "Pull_Aggregates", 1, 140
    AddButton ws, "R1", "Add cedant", "Add_Cedant", 1, 110

    Dim nm As Variant
    For Each nm In Array(SH_CEDANTS, SH_COUNTRIES, SH_RATES, SH_LISTS)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        ClearButtons ws
        AddButton ws, "H1", "Push lists + cedants", "Push_Masters", 1, 150
        AddButton ws, "K1", "Pull from Access", "Pull_Masters", 1, 140
    Next nm
    AddButton ThisWorkbook.Worksheets(SH_CEDANTS), "N1", "Add cedant", "Add_Cedant", 1, 110
    ThisWorkbook.Worksheets(SH_START).Activate
End Sub

Private Sub ClearButtons(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal anchor As String, ByVal caption As String, _
                      ByVal macroName As String, Optional ByVal rowsHigh As Long = 2, _
                      Optional ByVal widthPt As Double = 210)
    Dim r As Range, b As Object
    Set r = ws.Range(anchor)
    Set b = ws.Buttons.Add(r.Left + 2, r.Top + 2, widthPt, r.Resize(rowsHigh, 1).Height * 0.9 + 12)
    b.Name = "btn_" & macroName & "_" & Replace(anchor, "$", "")
    b.Caption = caption
    b.OnAction = macroName
    b.Font.Bold = True
End Sub

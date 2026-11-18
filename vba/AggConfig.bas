Attribute VB_Name = "AggConfig"
'===============================================================================
' AggConfig - constants and configuration (mirrors Universe server/src/config)
'
' Layering (same shape as Universe):
'   AggMacros  -> button entry points            (Universe: client screens)
'   AggService -> validation + orchestration     (Universe: server/src/services)
'   AggRepo    -> all SQL against Access         (Universe: modules/*/repositories)
'   AggDb      -> connection / command helpers   (Universe: server/src/db/pool.js)
'   AggMigrate -> numbered schema migrations     (Universe: db/migrations + runMigrations.js)
'   AggConfig  -> constants                      (Universe: config/env.js)
'===============================================================================
Option Explicit

Public Const APP_NAME As String = "Aggregate Tracker"
Public Const DB_FILENAME As String = "AggregateTracker.accdb"

' --- Sheet names --------------------------------------------------------------
Public Const SH_START As String = "Start"
Public Const SH_AGG As String = "Aggregates"
Public Const SH_CEDANTS As String = "Cedants"
Public Const SH_COUNTRIES As String = "Countries"
Public Const SH_RATES As String = "Rates"
Public Const SH_LISTS As String = "Lists"

' --- Grid geometry ------------------------------------------------------------
Public Const FIRST_ROW As Long = 6           ' first data row on every grid sheet
Public Const LAST_ROW_AGG As Long = 5005     ' 5,000 aggregate rows
Public Const LAST_ROW_CED As Long = 505      ' 500 cedants
Public Const LAST_ROW_CTRY As Long = 205     ' 200 countries
Public Const LAST_ROW_RATE As Long = 105     ' 100 currencies
Public Const LAST_ROW_LIST As Long = 105     ' 100 treaty types / classes

' --- Aggregates sheet columns (A..T are inputs / sync state; U.. are formulas) --
Public Const COL_STATUS As Long = 1
Public Const COL_ID As Long = 2
Public Const COL_CEDANT As Long = 3
Public Const COL_COUNTRY As Long = 4
Public Const COL_ZONE As Long = 5
Public Const COL_ZONENAME As Long = 6
Public Const COL_UWYEAR As Long = 7
Public Const COL_TREATY As Long = 8
Public Const COL_COB As Long = 9
Public Const COL_CCY As Long = 10
Public Const COL_ASAT As Long = 11
Public Const COL_EQ As Long = 12
Public Const COL_WS As Long = 13
Public Const COL_FLOOD As Long = 14
Public Const COL_SRCC As Long = 15
Public Const COL_OTHERS As Long = 16
Public Const COL_NOTES As Long = 17
Public Const COL_DELETE As Long = 18
Public Const COL_VERSION As Long = 19
Public Const COL_FP As Long = 20             ' fingerprint of values as last pulled
Public Const AGG_INPUT_COLS As Long = 20

' --- Master sheet columns -----------------------------------------------------
' Cedants:   A name | B home country | C active Y/N | D id | E status
' Countries: A name | B ISO2 | C ISO3 | D region | E id | F status
' Rates:     A code | B name | C USD per 1 unit | D effective date | E status
' Lists:     A treaty type | B category | D class of business | E code
Public Const CED_COLS As Long = 5
Public Const CTRY_COLS As Long = 6
Public Const RATE_COLS As Long = 5

' --- Defaults -----------------------------------------------------------------
Public Const UNSPECIFIED As String = "ALL"   ' blank zone / treaty type / class
Public Const MIN_UW_YEAR As Long = 1990
Public Const MAX_UW_YEAR As Long = 2100

' Path to the Access database. Start!C7 (named cfg_DbPath) overrides the default
' of "same folder as this workbook".
Public Function DbPath() As String
    Dim p As String
    On Error Resume Next
    p = Trim$(CStr(ThisWorkbook.Names("cfg_DbPath").RefersToRange.Value))
    On Error GoTo 0
    If Len(p) = 0 Then
        If Len(ThisWorkbook.Path) = 0 Then
            Err.Raise vbObjectError + 1000, APP_NAME, _
                "Save the workbook first, or enter a database path in Start!C7."
        End If
        p = ThisWorkbook.Path & Application.PathSeparator & DB_FILENAME
    End If
    If LCase$(Left$(p, 4)) = "http" Then
        Err.Raise vbObjectError + 1001, APP_NAME, _
            "The workbook is opened from OneDrive/SharePoint (" & p & ")." & vbCrLf & _
            "Access needs a local or network-drive path. Enter one in Start!C7, e.g. C:\Data\" & DB_FILENAME
    End If
    DbPath = p
End Function

' Who is making the change (Universe: services/auditActor.js)
Public Function Actor() As String
    Dim u As String
    u = Environ$("USERNAME")
    If Len(u) = 0 Then u = Application.UserName
    If Len(u) = 0 Then u = "EXCEL"
    Actor = Left$(u, 100)
End Function

Public Sub SetConfigValue(ByVal nm As String, ByVal v As Variant)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.Value = v
End Sub

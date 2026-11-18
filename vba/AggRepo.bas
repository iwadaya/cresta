Attribute VB_Name = "AggRepo"
'===============================================================================
' AggRepo - every SQL statement against the Access database lives here
' (Universe: modules/*/repositories). No validation and no sheet access.
'===============================================================================
Option Explicit

' Field positions in an aggregate record array (see AggService.BuildRecord)
Public Const R_CEDANT As Long = 0
Public Const R_COUNTRY As Long = 1
Public Const R_ZONE As Long = 2
Public Const R_ZONENAME As Long = 3
Public Const R_UWYEAR As Long = 4
Public Const R_TREATY As Long = 5
Public Const R_COB As Long = 6
Public Const R_CCY As Long = 7
Public Const R_ASAT As Long = 8
Public Const R_EQ As Long = 9
Public Const R_WS As Long = 10
Public Const R_FLOOD As Long = 11
Public Const R_SRCC As Long = 12
Public Const R_OTHERS As Long = 13
Public Const R_NOTES As Long = 14
Public Const R_WRITTEN As Long = 15
Public Const R_SIGNED As Long = 16

' --- Lookups -----------------------------------------------------------------
' name (case-insensitive) -> id
Public Function NameIdMap(ByVal cn As Object, ByVal tableName As String, _
                          ByVal nameCol As String, ByVal idCol As String) As Object
    Dim d As Object, rows As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    rows = AggDb.QueryRows(cn, "SELECT " & nameCol & ", " & idCol & " FROM " & tableName)
    If IsArray(rows) Then
        For i = 1 To UBound(rows, 1)
            d(Trim$(CStr(rows(i, 1)))) = rows(i, 2)
        Next i
    End If
    Set NameIdMap = d
End Function

Public Function CedantMap(ByVal cn As Object) As Object
    Set CedantMap = NameIdMap(cn, "cedant", "cedant_name", "cedant_id")
End Function

Public Function CountryMap(ByVal cn As Object) As Object
    Set CountryMap = NameIdMap(cn, "ref_country", "country_name", "country_id")
End Function

Public Function CountryCodeMap(ByVal cn As Object) As Object
    Set CountryCodeMap = NameIdMap(cn, "ref_country", "country_code", "country_id")
End Function

Public Function TreatyTypeMap(ByVal cn As Object) As Object
    Set TreatyTypeMap = NameIdMap(cn, "ref_treaty_type", "treaty_type_name", "treaty_type_id")
End Function

Public Function CobMap(ByVal cn As Object) As Object
    Set CobMap = NameIdMap(cn, "ref_class_of_business", "cob_name", "cob_id")
End Function

Public Function CurrencyMap(ByVal cn As Object) As Object
    Set CurrencyMap = NameIdMap(cn, "ref_currency", "currency_code", "currency_name")
End Function

' --- Aggregates ----------------------------------------------------------------
' Column order matches the Aggregates sheet from COL_ID to COL_NOTES, then row_version.
Public Function SelectAggregates(ByVal cn As Object) As Variant
    Dim s As String
    s = "SELECT d.agg_id, d.cedant_name, d.country_name, d.zone_id, d.zone_name, d.uw_year, d.treaty_type_name, "
    s = s & "d.cob_name, d.currency_code, d.as_at_date, a.written_line, a.signed_line, "
    s = s & "d.eq_agg, d.ws_agg, d.flood_agg, d.srcc_agg, d.others_agg, d.agg_notes, d.row_version "
    s = s & "FROM vw_aggregate_detail AS d INNER JOIN cedant_aggregate AS a ON d.agg_id = a.agg_id "
    s = s & "ORDER BY d.cedant_name, d.country_name, d.uw_year, d.zone_id, d.treaty_type_name, d.cob_name, d.currency_code"
    SelectAggregates = AggDb.QueryRows(cn, s)
End Function

Public Function InsertAggregate(ByVal cn As Object, ByVal rec As Variant) As Long
    Dim s As String, ts As Date, who As String
    ts = Now: who = Actor()
    s = "INSERT INTO cedant_aggregate (cedant_id, country_id, zone_id, zone_name, uw_year, treaty_type_id, "
    s = s & "cob_id, currency_code, as_at_date, eq_agg, ws_agg, flood_agg, srcc_agg, others_agg, agg_notes, "
    s = s & "written_line, signed_line, row_version, created_at, created_by, updated_at, updated_by) "
    s = s & "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    AggDb.ExecParams cn, s, rec(R_CEDANT), rec(R_COUNTRY), rec(R_ZONE), rec(R_ZONENAME), rec(R_UWYEAR), _
        rec(R_TREATY), rec(R_COB), rec(R_CCY), rec(R_ASAT), rec(R_EQ), rec(R_WS), rec(R_FLOOD), _
        rec(R_SRCC), rec(R_OTHERS), rec(R_NOTES), rec(R_WRITTEN), rec(R_SIGNED), 1&, ts, who, ts, who
    InsertAggregate = AggDb.LastId(cn)
End Function

' Optimistic lock (Universe: db/optimisticLock.js): only updates if nobody has
' changed the row since it was pulled. Returns rows affected (0 = conflict).
Public Function UpdateAggregate(ByVal cn As Object, ByVal aggId As Long, ByVal version As Long, _
                                ByVal rec As Variant) As Long
    Dim s As String
    s = "UPDATE cedant_aggregate SET cedant_id = ?, country_id = ?, zone_id = ?, zone_name = ?, "
    s = s & "uw_year = ?, treaty_type_id = ?, cob_id = ?, currency_code = ?, as_at_date = ?, "
    s = s & "eq_agg = ?, ws_agg = ?, flood_agg = ?, srcc_agg = ?, others_agg = ?, agg_notes = ?, "
    s = s & "written_line = ?, signed_line = ?, "
    s = s & "row_version = row_version + 1, updated_at = ?, updated_by = ? "
    s = s & "WHERE agg_id = ? AND row_version = ?"
    UpdateAggregate = AggDb.ExecParams(cn, s, rec(R_CEDANT), rec(R_COUNTRY), rec(R_ZONE), rec(R_ZONENAME), _
        rec(R_UWYEAR), rec(R_TREATY), rec(R_COB), rec(R_CCY), rec(R_ASAT), rec(R_EQ), rec(R_WS), _
        rec(R_FLOOD), rec(R_SRCC), rec(R_OTHERS), rec(R_NOTES), rec(R_WRITTEN), rec(R_SIGNED), _
        Now, Actor(), aggId, version)
End Function

Public Function DeleteAggregate(ByVal cn As Object, ByVal aggId As Long, ByVal version As Long) As Long
    DeleteAggregate = AggDb.ExecParams(cn, "DELETE FROM cedant_aggregate WHERE agg_id = ? AND row_version = ?", _
        aggId, version)
End Function

' --- Audit (Universe: services/audit.js -> audit_log) ---------------------------
Public Sub WriteAudit(ByVal cn As Object, ByVal entityType As String, ByVal entityId As Long, _
                      ByVal eventType As String, ByVal payload As String, Optional ByVal comment As String = "")
    Dim c As Variant
    If Len(comment) = 0 Then c = Null Else c = Left$(comment, 255)
    AggDb.ExecParams cn, "INSERT INTO audit_log (entity_type, entity_id, event_type, actor, payload, audit_comment, created_at) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?)", entityType, entityId, eventType, Actor(), payload, c, Now
End Sub

' --- Cedants -------------------------------------------------------------------
Public Function SelectCedants(ByVal cn As Object) As Variant
    SelectCedants = AggDb.QueryRows(cn, _
        "SELECT c.cedant_name, k.country_name, c.is_active, c.cedant_id " & _
        "FROM cedant AS c LEFT JOIN ref_country AS k ON c.home_country_id = k.country_id " & _
        "ORDER BY c.cedant_name")
End Function

Public Function InsertCedant(ByVal cn As Object, ByVal nm As String, ByVal homeCountryId As Variant, _
                             ByVal active As Boolean) As Long
    Dim ts As Date: ts = Now
    AggDb.ExecParams cn, "INSERT INTO cedant (cedant_name, home_country_id, is_active, created_at, created_by, updated_at, updated_by) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?)", nm, homeCountryId, active, ts, Actor(), ts, Actor()
    InsertCedant = AggDb.LastId(cn)
End Function

Public Function UpdateCedant(ByVal cn As Object, ByVal cedantId As Long, ByVal nm As String, _
                             ByVal homeCountryId As Variant, ByVal active As Boolean) As Long
    UpdateCedant = AggDb.ExecParams(cn, "UPDATE cedant SET cedant_name = ?, home_country_id = ?, is_active = ?, " & _
        "updated_at = ?, updated_by = ? WHERE cedant_id = ?", nm, homeCountryId, active, Now, Actor(), cedantId)
End Function

' --- Countries -----------------------------------------------------------------
Public Function SelectCountries(ByVal cn As Object) As Variant
    SelectCountries = AggDb.QueryRows(cn, _
        "SELECT country_name, country_code, country_iso3, region, country_id FROM ref_country ORDER BY country_name")
End Function

Public Sub UpsertCountry(ByVal cn As Object, ByVal existingId As Variant, ByVal nm As String, _
                         ByVal iso2 As String, ByVal iso3 As Variant, ByVal region As Variant)
    If IsEmpty(existingId) Or IsNull(existingId) Then
        AggDb.ExecParams cn, "INSERT INTO ref_country (country_code, country_iso3, country_name, region) VALUES (?, ?, ?, ?)", _
            iso2, iso3, nm, region
    Else
        AggDb.ExecParams cn, "UPDATE ref_country SET country_code = ?, country_iso3 = ?, country_name = ?, region = ? " & _
            "WHERE country_id = ?", iso2, iso3, nm, region, CLng(existingId)
    End If
End Sub

' --- Currencies and FX ---------------------------------------------------------
Public Function SelectRates(ByVal cn As Object) As Variant
    SelectRates = AggDb.QueryRows(cn, _
        "SELECT c.currency_code, c.currency_name, f.rate_to_usd, f.effective_date " & _
        "FROM ref_currency AS c LEFT JOIN vw_fx_latest AS f ON c.currency_code = f.currency_code " & _
        "ORDER BY c.currency_code")
End Function

Public Sub UpsertCurrency(ByVal cn As Object, ByVal exists As Boolean, ByVal code As String, ByVal nm As String)
    If exists Then
        AggDb.ExecParams cn, "UPDATE ref_currency SET currency_name = ? WHERE currency_code = ?", nm, code
    Else
        AggDb.ExecParams cn, "INSERT INTO ref_currency (currency_code, currency_name) VALUES (?, ?)", code, nm
    End If
End Sub

' One rate per currency per effective date: same date = correction, new date = new history row.
' Returns True when the database changed.
Public Function UpsertRate(ByVal cn As Object, ByVal code As String, ByVal rateToUsd As Double, _
                           ByVal effective As Date) As Boolean
    Dim cur As Variant
    cur = AggDb.Scalar(cn, "SELECT rate_to_usd FROM ref_exchange_rate WHERE currency_code = ? AND effective_date = ?", _
        code, effective)
    If IsNull(cur) Then
        AggDb.ExecParams cn, "INSERT INTO ref_exchange_rate (currency_code, rate_to_usd, effective_date, created_at, created_by) " & _
            "VALUES (?, ?, ?, ?, ?)", code, rateToUsd, effective, Now, Actor()
        UpsertRate = True
    ElseIf Abs(CDbl(cur) - rateToUsd) > 0.0000000001 Then
        AggDb.ExecParams cn, "UPDATE ref_exchange_rate SET rate_to_usd = ? WHERE currency_code = ? AND effective_date = ?", _
            rateToUsd, code, effective
        UpsertRate = True
    End If
End Function

' --- Treaty types and classes of business ---------------------------------------
Public Function SelectTreatyTypes(ByVal cn As Object) As Variant
    SelectTreatyTypes = AggDb.QueryRows(cn, _
        "SELECT treaty_type_name, treaty_category FROM ref_treaty_type ORDER BY treaty_category, treaty_type_name")
End Function

Public Function SelectCobs(ByVal cn As Object) As Variant
    SelectCobs = AggDb.QueryRows(cn, "SELECT cob_name, cob_code FROM ref_class_of_business ORDER BY cob_name")
End Function

Public Sub UpsertTreatyType(ByVal cn As Object, ByVal exists As Boolean, ByVal nm As String, ByVal category As Variant)
    If exists Then
        AggDb.ExecParams cn, "UPDATE ref_treaty_type SET treaty_category = ? WHERE treaty_type_name = ?", category, nm
    Else
        AggDb.ExecParams cn, "INSERT INTO ref_treaty_type (treaty_type_name, treaty_category) VALUES (?, ?)", nm, category
    End If
End Sub

Public Sub UpsertCob(ByVal cn As Object, ByVal exists As Boolean, ByVal nm As String, ByVal code As Variant)
    If exists Then
        AggDb.ExecParams cn, "UPDATE ref_class_of_business SET cob_code = ? WHERE cob_name = ?", code, nm
    Else
        AggDb.ExecParams cn, "INSERT INTO ref_class_of_business (cob_name, cob_code) VALUES (?, ?)", nm, code
    End If
End Sub

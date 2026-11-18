Attribute VB_Name = "AggMigrate"
'===============================================================================
' AggMigrate - numbered, forward-only schema migrations, recorded in
' schema_migrations (Universe: server/src/db/migrations/NNN_*.sql +
' startup/runMigrations.js). To change the schema later, ADD a new migration
' number - never edit one that has already been applied.
'===============================================================================
Option Explicit

Public Const MIG_COUNT As Long = 6

Private Function MigrationName(ByVal v As Long) As String
    Select Case v
        Case 1: MigrationName = "001_foundation"
        Case 2: MigrationName = "002_reference_data"
        Case 3: MigrationName = "003_cedant_aggregates"
        Case 4: MigrationName = "004_reporting_views"
        Case 5: MigrationName = "005_seed_unspecified"
        Case 6: MigrationName = "006_written_signed_lines"
    End Select
End Function

' Returns the DDL statements of one migration.
Private Function MigrationSql(ByVal v As Long) As Collection
    Dim c As New Collection, s As String
    Select Case v

    Case 1 ' --- foundation: migration ledger + audit log (Universe 001_foundation.sql)
        s = "CREATE TABLE schema_migrations ("
        s = s & "mig_version LONG NOT NULL, mig_name VARCHAR(100) NOT NULL, "
        s = s & "applied_at DATETIME NOT NULL, applied_by VARCHAR(100), "
        s = s & "CONSTRAINT pk_schema_migrations PRIMARY KEY (mig_version))"
        c.Add s
        s = "CREATE TABLE audit_log ("
        s = s & "event_id COUNTER NOT NULL, entity_type VARCHAR(50) NOT NULL, entity_id LONG NOT NULL, "
        s = s & "event_type VARCHAR(50) NOT NULL, actor VARCHAR(100) NOT NULL, payload LONGTEXT, "
        s = s & "audit_comment VARCHAR(255), created_at DATETIME NOT NULL, "
        s = s & "CONSTRAINT pk_audit_log PRIMARY KEY (event_id))"
        c.Add s
        c.Add "CREATE INDEX ix_audit_log_entity ON audit_log (entity_type, entity_id, created_at)"

    Case 2 ' --- reference data (Universe: country, currency, ref_exchange_rate, treaty_type, class_of_business)
        s = "CREATE TABLE ref_country ("
        s = s & "country_id COUNTER NOT NULL, country_code VARCHAR(2) NOT NULL, country_iso3 VARCHAR(3), "
        s = s & "country_name VARCHAR(100) NOT NULL, region VARCHAR(50), "
        s = s & "CONSTRAINT pk_ref_country PRIMARY KEY (country_id))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_ref_country_code ON ref_country (country_code)"
        c.Add "CREATE UNIQUE INDEX ux_ref_country_name ON ref_country (country_name)"

        s = "CREATE TABLE ref_currency ("
        s = s & "currency_code VARCHAR(3) NOT NULL, currency_name VARCHAR(60) NOT NULL, "
        s = s & "CONSTRAINT pk_ref_currency PRIMARY KEY (currency_code))"
        c.Add s

        ' rate_to_usd = USD per 1 unit of the currency (same convention as Universe ref_exchange_rate)
        s = "CREATE TABLE ref_exchange_rate ("
        s = s & "rate_id COUNTER NOT NULL, currency_code VARCHAR(3) NOT NULL, rate_to_usd DOUBLE NOT NULL, "
        s = s & "effective_date DATETIME NOT NULL, created_at DATETIME NOT NULL, created_by VARCHAR(100), "
        s = s & "CONSTRAINT pk_ref_exchange_rate PRIMARY KEY (rate_id), "
        s = s & "CONSTRAINT fk_fx_currency FOREIGN KEY (currency_code) REFERENCES ref_currency (currency_code))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_fx_ccy_date ON ref_exchange_rate (currency_code, effective_date)"

        s = "CREATE TABLE ref_treaty_type ("
        s = s & "treaty_type_id COUNTER NOT NULL, treaty_type_name VARCHAR(80) NOT NULL, treaty_category VARCHAR(30), "
        s = s & "CONSTRAINT pk_ref_treaty_type PRIMARY KEY (treaty_type_id))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_treaty_type_name ON ref_treaty_type (treaty_type_name)"

        s = "CREATE TABLE ref_class_of_business ("
        s = s & "cob_id COUNTER NOT NULL, cob_name VARCHAR(80) NOT NULL, cob_code VARCHAR(10), "
        s = s & "CONSTRAINT pk_ref_cob PRIMARY KEY (cob_id))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_cob_name ON ref_class_of_business (cob_name)"

    Case 3 ' --- cedants + aggregate fact table (Universe: company + contract_cresta_data)
        s = "CREATE TABLE cedant ("
        s = s & "cedant_id COUNTER NOT NULL, cedant_name VARCHAR(150) NOT NULL, home_country_id LONG, "
        s = s & "is_active BIT NOT NULL, created_at DATETIME NOT NULL, created_by VARCHAR(100), "
        s = s & "updated_at DATETIME NOT NULL, updated_by VARCHAR(100), "
        s = s & "CONSTRAINT pk_cedant PRIMARY KEY (cedant_id), "
        s = s & "CONSTRAINT fk_cedant_country FOREIGN KEY (home_country_id) REFERENCES ref_country (country_id))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_cedant_name ON cedant (cedant_name)"

        ' One row = one cedant's aggregate exposure in one country / CRESTA zone /
        ' UW year / treaty type / class, in original currency, split by peril.
        ' Perils are NOT additive (the same sum insured can be exposed to EQ and WS).
        s = "CREATE TABLE cedant_aggregate ("
        s = s & "agg_id COUNTER NOT NULL, cedant_id LONG NOT NULL, country_id LONG NOT NULL, "
        s = s & "zone_id VARCHAR(64) NOT NULL, zone_name VARCHAR(150), uw_year LONG NOT NULL, "
        s = s & "treaty_type_id LONG NOT NULL, cob_id LONG NOT NULL, currency_code VARCHAR(3) NOT NULL, "
        s = s & "as_at_date DATETIME, "
        s = s & "eq_agg DOUBLE NOT NULL, ws_agg DOUBLE NOT NULL, flood_agg DOUBLE NOT NULL, "
        s = s & "srcc_agg DOUBLE NOT NULL, others_agg DOUBLE NOT NULL, agg_notes VARCHAR(255), "
        s = s & "row_version LONG NOT NULL, "
        s = s & "created_at DATETIME NOT NULL, created_by VARCHAR(100), "
        s = s & "updated_at DATETIME NOT NULL, updated_by VARCHAR(100), "
        s = s & "CONSTRAINT pk_cedant_aggregate PRIMARY KEY (agg_id), "
        s = s & "CONSTRAINT fk_agg_cedant FOREIGN KEY (cedant_id) REFERENCES cedant (cedant_id), "
        s = s & "CONSTRAINT fk_agg_country FOREIGN KEY (country_id) REFERENCES ref_country (country_id), "
        s = s & "CONSTRAINT fk_agg_treaty FOREIGN KEY (treaty_type_id) REFERENCES ref_treaty_type (treaty_type_id), "
        s = s & "CONSTRAINT fk_agg_cob FOREIGN KEY (cob_id) REFERENCES ref_class_of_business (cob_id), "
        s = s & "CONSTRAINT fk_agg_currency FOREIGN KEY (currency_code) REFERENCES ref_currency (currency_code))"
        c.Add s
        ' Natural key: one row per cedant/country/zone/year/treaty/class/currency
        s = "CREATE UNIQUE INDEX ux_cedant_aggregate_key ON cedant_aggregate "
        s = s & "(cedant_id, country_id, zone_id, uw_year, treaty_type_id, cob_id, currency_code)"
        c.Add s
        c.Add "CREATE INDEX ix_cedant_aggregate_year ON cedant_aggregate (uw_year)"

    Case 4 ' --- reporting views (saved queries you can open in Access)
        c.Add "CREATE VIEW vw_fx_latest_date AS SELECT currency_code, MAX(effective_date) AS max_date " & _
              "FROM ref_exchange_rate GROUP BY currency_code"

        s = "CREATE VIEW vw_fx_latest AS SELECT r.currency_code, r.rate_to_usd, r.effective_date "
        s = s & "FROM ref_exchange_rate AS r INNER JOIN vw_fx_latest_date AS m "
        s = s & "ON (r.currency_code = m.currency_code) AND (r.effective_date = m.max_date)"
        c.Add s

        s = "CREATE VIEW vw_aggregate_detail AS SELECT a.agg_id, c.cedant_name, k.country_name, a.zone_id, "
        s = s & "a.zone_name, a.uw_year, t.treaty_type_name, b.cob_name, a.currency_code, a.as_at_date, "
        s = s & "a.eq_agg, a.ws_agg, a.flood_agg, a.srcc_agg, a.others_agg, a.agg_notes, a.row_version, "
        s = s & "k.country_code, k.region, a.updated_at, a.updated_by "
        s = s & "FROM (((cedant_aggregate AS a "
        s = s & "INNER JOIN cedant AS c ON a.cedant_id = c.cedant_id) "
        s = s & "INNER JOIN ref_country AS k ON a.country_id = k.country_id) "
        s = s & "INNER JOIN ref_treaty_type AS t ON a.treaty_type_id = t.treaty_type_id) "
        s = s & "INNER JOIN ref_class_of_business AS b ON a.cob_id = b.cob_id"
        c.Add s

        s = "CREATE VIEW vw_aggregate_usd AS SELECT d.agg_id, d.cedant_name, d.country_name, d.country_code, "
        s = s & "d.region, d.zone_id, d.zone_name, d.uw_year, d.treaty_type_name, d.cob_name, d.currency_code, "
        s = s & "f.rate_to_usd, d.eq_agg * f.rate_to_usd AS eq_usd, d.ws_agg * f.rate_to_usd AS ws_usd, "
        s = s & "d.flood_agg * f.rate_to_usd AS flood_usd, d.srcc_agg * f.rate_to_usd AS srcc_usd, "
        s = s & "d.others_agg * f.rate_to_usd AS others_usd "
        s = s & "FROM vw_aggregate_detail AS d LEFT JOIN vw_fx_latest AS f ON d.currency_code = f.currency_code"
        c.Add s

        s = "CREATE VIEW vw_agg_by_cedant AS SELECT cedant_name, uw_year, COUNT(*) AS row_count, "
        s = s & "SUM(eq_usd) AS total_eq_usd, SUM(ws_usd) AS total_ws_usd, SUM(flood_usd) AS total_flood_usd, "
        s = s & "SUM(srcc_usd) AS total_srcc_usd, SUM(others_usd) AS total_others_usd "
        s = s & "FROM vw_aggregate_usd GROUP BY cedant_name, uw_year"
        c.Add s

        s = "CREATE VIEW vw_agg_by_country AS SELECT country_name, region, uw_year, COUNT(*) AS row_count, "
        s = s & "SUM(eq_usd) AS total_eq_usd, SUM(ws_usd) AS total_ws_usd, SUM(flood_usd) AS total_flood_usd, "
        s = s & "SUM(srcc_usd) AS total_srcc_usd, SUM(others_usd) AS total_others_usd "
        s = s & "FROM vw_aggregate_usd GROUP BY country_name, region, uw_year"
        c.Add s

        s = "CREATE VIEW vw_agg_by_cedant_country AS SELECT cedant_name, country_name, uw_year, COUNT(*) AS row_count, "
        s = s & "SUM(eq_usd) AS total_eq_usd, SUM(ws_usd) AS total_ws_usd, SUM(flood_usd) AS total_flood_usd, "
        s = s & "SUM(srcc_usd) AS total_srcc_usd, SUM(others_usd) AS total_others_usd "
        s = s & "FROM vw_aggregate_usd GROUP BY cedant_name, country_name, uw_year"
        c.Add s

        ' Rows whose currency has no rate - they drop out of every USD total
        s = "CREATE VIEW vw_agg_missing_fx AS SELECT agg_id, cedant_name, country_name, uw_year, currency_code "
        s = s & "FROM vw_aggregate_usd WHERE rate_to_usd IS NULL"
        c.Add s

    Case 5
        ' data seed, handled in ApplySeed

    Case 6 ' --- our written / signed line per aggregate row + share reporting views
        ' Fractions (7.5% = 0.075); NULL = not known yet. Signed <= written is checked in AggService.
        c.Add "ALTER TABLE cedant_aggregate ADD COLUMN written_line DOUBLE"
        c.Add "ALTER TABLE cedant_aggregate ADD COLUMN signed_line DOUBLE"

        s = "CREATE VIEW vw_aggregate_share_usd AS SELECT u.agg_id, u.cedant_name, u.country_name, u.country_code, "
        s = s & "u.region, u.zone_id, u.uw_year, u.treaty_type_name, u.cob_name, u.currency_code, u.rate_to_usd, "
        s = s & "a.written_line, a.signed_line, "
        s = s & "u.eq_usd, u.ws_usd, u.flood_usd, u.srcc_usd, u.others_usd, "
        s = s & "u.eq_usd * IIf(a.written_line Is Null, 0, a.written_line) AS eq_written_usd, "
        s = s & "u.ws_usd * IIf(a.written_line Is Null, 0, a.written_line) AS ws_written_usd, "
        s = s & "u.flood_usd * IIf(a.written_line Is Null, 0, a.written_line) AS flood_written_usd, "
        s = s & "u.srcc_usd * IIf(a.written_line Is Null, 0, a.written_line) AS srcc_written_usd, "
        s = s & "u.others_usd * IIf(a.written_line Is Null, 0, a.written_line) AS others_written_usd, "
        s = s & "u.eq_usd * IIf(a.signed_line Is Null, 0, a.signed_line) AS eq_signed_usd, "
        s = s & "u.ws_usd * IIf(a.signed_line Is Null, 0, a.signed_line) AS ws_signed_usd, "
        s = s & "u.flood_usd * IIf(a.signed_line Is Null, 0, a.signed_line) AS flood_signed_usd, "
        s = s & "u.srcc_usd * IIf(a.signed_line Is Null, 0, a.signed_line) AS srcc_signed_usd, "
        s = s & "u.others_usd * IIf(a.signed_line Is Null, 0, a.signed_line) AS others_signed_usd "
        s = s & "FROM vw_aggregate_usd AS u INNER JOIN cedant_aggregate AS a ON u.agg_id = a.agg_id"
        c.Add s

        c.Add ShareGroupView("vw_agg_share_by_cedant", "cedant_name")
        c.Add ShareGroupView("vw_agg_share_by_country", "country_name, region")
        c.Add ShareGroupView("vw_agg_share_by_cedant_country", "cedant_name, country_name")
    End Select
    Set MigrationSql = c
End Function

' Grouped view over vw_aggregate_share_usd: 100%, written and signed USD per peril.
Private Function ShareGroupView(ByVal viewName As String, ByVal groupCols As String) As String
    Dim s As String, p As Variant
    s = "CREATE VIEW " & viewName & " AS SELECT " & groupCols & ", uw_year, COUNT(*) AS row_count"
    For Each p In Array("eq", "ws", "flood", "srcc", "others")
        s = s & ", SUM(" & p & "_usd) AS " & p & "_100_usd"
        s = s & ", SUM(" & p & "_written_usd) AS " & p & "_written_usd"
        s = s & ", SUM(" & p & "_signed_usd) AS " & p & "_signed_usd"
    Next p
    ShareGroupView = s & " FROM vw_aggregate_share_usd GROUP BY " & groupCols & ", uw_year"
End Function

' "ALL" rows so blank treaty type / class on the sheet always resolves to an id.
Private Sub ApplySeed(ByVal cn As Object)
    If IsNull(AggDb.Scalar(cn, "SELECT treaty_type_id FROM ref_treaty_type WHERE treaty_type_name = ?", UNSPECIFIED)) Then
        AggDb.ExecParams cn, "INSERT INTO ref_treaty_type (treaty_type_name, treaty_category) VALUES (?, ?)", _
            UNSPECIFIED, "UNSPECIFIED"
    End If
    If IsNull(AggDb.Scalar(cn, "SELECT cob_id FROM ref_class_of_business WHERE cob_name = ?", UNSPECIFIED)) Then
        AggDb.ExecParams cn, "INSERT INTO ref_class_of_business (cob_name, cob_code) VALUES (?, ?)", _
            UNSPECIFIED, UNSPECIFIED
    End If
    If IsNull(AggDb.Scalar(cn, "SELECT currency_code FROM ref_currency WHERE currency_code = ?", "USD")) Then
        AggDb.ExecParams cn, "INSERT INTO ref_currency (currency_code, currency_name) VALUES (?, ?)", _
            "USD", "US Dollar"
    End If
End Sub

' Name of the object a CREATE TABLE / CREATE VIEW statement creates, else "".
Private Function CreatedObjectName(ByVal sql As String) As String
    Dim u As String, rest As String, p As Long
    u = UCase$(Trim$(sql))
    If Left$(u, 13) = "CREATE TABLE " Then
        rest = Trim$(Mid$(Trim$(sql), 14))
    ElseIf Left$(u, 12) = "CREATE VIEW " Then
        rest = Trim$(Mid$(Trim$(sql), 13))
    Else
        Exit Function
    End If
    p = InStr(rest, " ")
    If p = 0 Then p = InStr(rest, "(")
    If p > 0 Then rest = Left$(rest, p - 1)
    CreatedObjectName = Replace(rest, "(", "")
End Function

Private Function AppliedVersions(ByVal cn As Object) As Object
    Dim d As Object, rows As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    If AggDb.ObjectExists(cn, "schema_migrations") Then
        rows = AggDb.QueryRows(cn, "SELECT mig_version FROM schema_migrations")
        If IsArray(rows) Then
            For i = 1 To UBound(rows, 1)
                d(CLng(rows(i, 1))) = True
            Next i
        End If
    End If
    Set AppliedVersions = d
End Function

' Apply every pending migration in order. Returns a one-line summary.
' Re-runnable: objects that already exist are skipped, so a migration that
' stopped half-way can simply be run again.
Public Function RunMigrations(ByVal cn As Object) As String
    Dim applied As Object, v As Long, stmts As Collection, i As Long
    Dim sql As String, objName As String, ran As Long
    Set applied = AppliedVersions(cn)
    For v = 1 To MIG_COUNT
        If Not applied.Exists(v) Then
            Set stmts = MigrationSql(v)
            For i = 1 To stmts.Count
                sql = stmts(i)
                objName = CreatedObjectName(sql)
                If Len(objName) > 0 Then
                    If Not AggDb.ObjectExists(cn, objName) Then ExecDdl cn, sql, v
                Else
                    ExecIndexDdl cn, sql, v
                End If
            Next i
            If v = 5 Then ApplySeed cn
            AggDb.ExecParams cn, "INSERT INTO schema_migrations (mig_version, mig_name, applied_at, applied_by) VALUES (?, ?, ?, ?)", _
                v, MigrationName(v), Now, Actor()
            ran = ran + 1
        End If
    Next v
    If ran = 0 Then
        RunMigrations = "Schema up to date (" & MIG_COUNT & " migrations)."
    Else
        RunMigrations = "Applied " & ran & " migration(s); schema at " & MigrationName(MIG_COUNT) & "."
    End If
End Function

Private Sub ExecDdl(ByVal cn As Object, ByVal sql As String, ByVal v As Long)
    On Error GoTo EH
    AggDb.ExecSql cn, sql
    Exit Sub
EH:
    Err.Raise Err.Number, APP_NAME, "Migration " & MigrationName(v) & " failed:" & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & sql
End Sub

' CREATE INDEX: ignore "already exists" so re-runs are safe.
Private Sub ExecIndexDdl(ByVal cn As Object, ByVal sql As String, ByVal v As Long)
    Dim d As String, n As Long
    On Error Resume Next
    AggDb.ExecSql cn, sql
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    If n <> 0 Then
        If InStr(1, d, "already", vbTextCompare) = 0 Then
            Err.Raise n, APP_NAME, "Migration " & MigrationName(v) & " failed:" & vbCrLf & d & vbCrLf & vbCrLf & sql
        End If
    End If
End Sub

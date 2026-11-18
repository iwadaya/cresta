"""Generate Aggregate_Tracker.xlsx (clean) and Aggregate_Tracker_DEMO.xlsx.

Layout contract with the VBA modules (AggConfig.bas):
  header row 5, data from row 6 on every grid sheet;
  Aggregates A..V = inputs + sync state (L/M = written/signed line),
  W = FX, X..AB = 100% USD, AC..AG = written share USD, AH..AL = signed share USD.
"""
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, GradientFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, DataBarRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

# ---- Universe palette (shared/universeXlsxTheme.js) ---------------------------
TEAL, GREEN, HEADER, INK, INK_SOFT = "0D5C75", "2F8F6B", "1E7A5A", "16323B", "5B7079"
TOTAL_BG, TOTAL_INK, GRID, SUB_BG = "CDE7DD", "0D4A3A", "D8E2E6", "F4F9FA"
CALC_BG, INPUT_BLUE = "EEF3F5", "0000FF"
FONT = "Arial"
MONEY = '#,##0;[Red]-#,##0;"-"'
RATE = '#,##0.000000'
DATE = 'dd-mmm-yyyy'
PCT = '0.00%'
STAMP = 'dd-mmm-yyyy hh:mm'

FIRST, LAST_AGG, LAST_CED, LAST_CTRY, LAST_RATE, LAST_LIST = 6, 5005, 505, 205, 105, 105

# Aggregates calculated columns (must match vba/AggConfig.bas: inputs end at V)
USD_100 = ["X", "Y", "Z", "AA", "AB"]
USD_WRITTEN = ["AC", "AD", "AE", "AF", "AG"]
USD_SIGNED = ["AH", "AI", "AJ", "AK", "AL"]
AG_NCOL = 38
USD_COLS = {"100": USD_100, "written": USD_WRITTEN, "signed": USD_SIGNED}
# (sheet name, description, column block)
BASES = [("100% Share", "100% of the cedant's aggregate", "100"),
         ("Written Share", "our written share (100% x written line)", "written"),
         ("Signed Share", "our signed share (100% x signed line)", "signed")]
DEFAULT_FACTOR = 0.0035     # 0.35% - given by the user
BASIS_LIST = ["100% share", "Written share", "Signed share", "Written x factor", "Signed x factor"]

thin = Side(style="thin", color=GRID)
BOX = Border(top=thin, left=thin, bottom=thin, right=thin)


def f(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", fgColor=c)


def title_bar(ws, ncols, title, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(1, 1, f"AGGREGATE TRACKER    {title}")
    c.font = Font(name=FONT, size=15, bold=True, color="FFFFFF")
    c.alignment = Alignment(vertical="center", indent=1)
    c.fill = GradientFill(degree=0, stop=(TEAL, GREEN))
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 14
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
    s = ws.cell(3, 1, subtitle)
    s.font = f(9, italic=True, color=INK_SOFT)
    s.fill = fill(SUB_BG)
    s.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[3].height = 16


def header(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row, start_col + i, lab)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = fill(HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 30


def widths(ws, ws_widths):
    for col, w in ws_widths.items():
        ws.column_dimensions[col].width = w


def style_range(ws, ref, font=None, number_format=None, bg=None, align=None, border=True):
    for row in ws[ref]:
        for c in row:
            if font: c.font = font
            if number_format: c.number_format = number_format
            if bg: c.fill = fill(bg)
            if align: c.alignment = align
            if border: c.border = BOX


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


# ---- Reference data (from Universe startup/ensureReferenceData.js, plus extra
#      African markets and currencies the Universe seed does not carry) -----------
REGION = {}
for codes, reg in [
    ("AE SA KW BH OM QA JO LB", "Middle East"),
    ("EG MA TN DZ ZA NG KE ZM ZW GH TZ UG ET RW MZ BW NA MW SN CI CM AO CD MU", "Africa"),
    ("GB FR DE IT ES CH NL BE TR", "Europe"),
    ("IN PK LK BD MY SG JP CN KR TH ID PH", "Asia"),
    ("AU NZ", "Oceania"), ("US CA MX BR CL CO AR", "Americas")]:
    for c in codes.split():
        REGION[c] = reg

COUNTRIES = [
    ('United Arab Emirates','AE','ARE'),('Saudi Arabia','SA','SAU'),('Kuwait','KW','KWT'),('Bahrain','BH','BHR'),
    ('Oman','OM','OMN'),('Qatar','QA','QAT'),('Jordan','JO','JOR'),('Lebanon','LB','LBN'),('Egypt','EG','EGY'),
    ('Morocco','MA','MAR'),('Tunisia','TN','TUN'),('Algeria','DZ','DZA'),('South Africa','ZA','ZAF'),
    ('Nigeria','NG','NGA'),('Kenya','KE','KEN'),('United Kingdom','GB','GBR'),('France','FR','FRA'),
    ('Germany','DE','DEU'),('Italy','IT','ITA'),('Spain','ES','ESP'),('Switzerland','CH','CHE'),
    ('Netherlands','NL','NLD'),('Belgium','BE','BEL'),('Turkey','TR','TUR'),('India','IN','IND'),
    ('Pakistan','PK','PAK'),('Sri Lanka','LK','LKA'),('Bangladesh','BD','BGD'),('Malaysia','MY','MYS'),
    ('Singapore','SG','SGP'),('Japan','JP','JPN'),('China','CN','CHN'),('South Korea','KR','KOR'),
    ('Thailand','TH','THA'),('Indonesia','ID','IDN'),('Philippines','PH','PHL'),('Australia','AU','AUS'),
    ('New Zealand','NZ','NZL'),('United States','US','USA'),('Canada','CA','CAN'),('Mexico','MX','MEX'),
    ('Brazil','BR','BRA'),('Chile','CL','CHL'),('Colombia','CO','COL'),('Argentina','AR','ARG'),
    # additions for Africa coverage
    ('Zambia','ZM','ZMB'),('Zimbabwe','ZW','ZWE'),('Ghana','GH','GHA'),('Tanzania','TZ','TZA'),
    ('Uganda','UG','UGA'),('Ethiopia','ET','ETH'),('Rwanda','RW','RWA'),('Mozambique','MZ','MOZ'),
    ('Botswana','BW','BWA'),('Namibia','NA','NAM'),('Malawi','MW','MWI'),('Senegal','SN','SEN'),
    ("Cote d'Ivoire",'CI','CIV'),('Cameroon','CM','CMR'),('Angola','AO','AGO'),
    ('DR Congo','CD','COD'),('Mauritius','MU','MUS'),
]
COUNTRIES.sort(key=lambda r: r[0])

CURRENCIES = [
    ('USD','US Dollar'),('EUR','Euro'),('GBP','British Pound'),('AED','UAE Dirham'),
    ('SAR','Saudi Riyal'),('KWD','Kuwaiti Dinar'),('BHD','Bahraini Dinar'),
    ('OMR','Omani Rial'),('QAR','Qatari Riyal'),('JOD','Jordanian Dinar'),
    ('EGP','Egyptian Pound'),('MAD','Moroccan Dirham'),('ZAR','South African Rand'),
    ('NGN','Nigerian Naira'),('KES','Kenyan Shilling'),('INR','Indian Rupee'),
    ('PKR','Pakistani Rupee'),('LKR','Sri Lankan Rupee'),('MYR','Malaysian Ringgit'),
    ('SGD','Singapore Dollar'),('JPY','Japanese Yen'),('CNY','Chinese Yuan'),
    ('KRW','South Korean Won'),('THB','Thai Baht'),('IDR','Indonesian Rupiah'),
    ('AUD','Australian Dollar'),('NZD','New Zealand Dollar'),('CAD','Canadian Dollar'),
    ('CHF','Swiss Franc'),('TRY','Turkish Lira'),('BRL','Brazilian Real'),('MXN','Mexican Peso'),
    # additions for Africa coverage
    ('ZMW','Zambian Kwacha'),('ZWG','Zimbabwe Gold'),('GHS','Ghanaian Cedi'),('TZS','Tanzanian Shilling'),
    ('UGX','Ugandan Shilling'),('ETB','Ethiopian Birr'),('RWF','Rwandan Franc'),('MZN','Mozambican Metical'),
    ('BWP','Botswana Pula'),('NAD','Namibian Dollar'),('MWK','Malawian Kwacha'),('XOF','West African CFA Franc'),
    ('XAF','Central African CFA Franc'),('AOA','Angolan Kwanza'),('CDF','Congolese Franc'),('MUR','Mauritian Rupee'),
]
CURRENCIES.sort()
# Only USD and fixed USD pegs are pre-filled (USD per 1 unit). Every floating
# currency is left blank for the user to enter - no rate is guessed.
PEGS = {'USD': (1.0, 'Base currency'), 'SAR': (1/3.75, 'Peg 3.75 SAR/USD'),
        'AED': (1/3.6725, 'Peg 3.6725 AED/USD'), 'BHD': (1/0.376, 'Peg 0.376 BHD/USD'),
        'QAR': (1/3.64, 'Peg 3.64 QAR/USD'), 'JOD': (1/0.709, 'Peg 0.709 JOD/USD'),
        'OMR': (1/0.3845, 'Peg 0.3845 OMR/USD')}
PEG_DATE = date(2026, 1, 1)

TREATY_TYPES = [('ALL','UNSPECIFIED'),
    ('Quota Share','PROPORTIONAL'),('Quota Share & Surplus','PROPORTIONAL'),
    ('First Surplus','PROPORTIONAL'),('Second Surplus','PROPORTIONAL'),
    ('Third Surplus','PROPORTIONAL'),('Fac Oblig','PROPORTIONAL'),
    ('Risk XL','NON_PROPORTIONAL'),('CAT XL','NON_PROPORTIONAL'),
    ('Risk & CAT XL','NON_PROPORTIONAL'),('Stop Loss','NON_PROPORTIONAL'),
    ('Aggregate XL','NON_PROPORTIONAL')]
CLASSES = [('ALL','ALL'),
    ('Property','PROP'),('Motor','MOT'),('Marine','MAR'),('Engineering','ENG'),
    ('Liability','LIA'),('Medical','MED'),('Aviation','AVI'),('Energy','ENE'),
    ('Agriculture','AGR'),('Credit & Surety','CS'),('Miscellaneous','MISC'),
    ('Life','LIFE'),('Group Life','GL'),('Workers Compensation','WC')]
PERILS = [('EQ','Earthquake'),('WS','Windstorm / tropical cyclone'),('Flood','Flood (riverine, flash, surge)'),
          ('SRCC','Strikes, riots & civil commotion'),('Others','Any other peril tracked')]
MATRIX_COUNTRIES = ['Saudi Arabia','United Arab Emirates','Kuwait','Bahrain','Oman','Qatar','Jordan','Egypt',
                    'Morocco','South Africa','Nigeria','Kenya','Zambia','Zimbabwe','Ghana','Tanzania','Uganda',
                    'Ethiopia','Pakistan','India']

# Universe cedant seed (ensureReferenceData.js + seeds/002_reference_data.sql,
# merged; ADNIC / Oman Insurance / Kuwait Insurance dropped as duplicates).
# Must match vba/AggSeed.bas.
UNIVERSE_CEDANTS = [
    ('Tawuniya','Saudi Arabia'),('MEDGULF','Saudi Arabia'),('Al Rajhi Takaful','Saudi Arabia'),
    ('Bupa Arabia','Saudi Arabia'),('Walaa Insurance','Saudi Arabia'),('SALAMA','Saudi Arabia'),
    ('Malath Insurance','Saudi Arabia'),('Gulf Union Insurance','Saudi Arabia'),
    ('Orient Insurance','United Arab Emirates'),('Oman Insurance Co','United Arab Emirates'),
    ('Abu Dhabi National Insurance','United Arab Emirates'),('Sukoon Insurance','United Arab Emirates'),
    ('Dubai Insurance','United Arab Emirates'),('Salama Islamic Insurance','United Arab Emirates'),
    ('GIG Bahrain','Bahrain'),('Solidarity Bahrain','Bahrain'),
    ('Kuwait Insurance Co','Kuwait'),('Gulf Insurance Group','Kuwait'),('Warba Insurance','Kuwait'),
    ('Misr Insurance','Egypt'),('GIG Egypt','Egypt'),('Allianz Egypt','Egypt'),
    ('Dhofar Insurance','Oman'),('National Life & General','Oman'),
    ('Aviva','United Kingdom'),('AXA UK','United Kingdom'),('RSA Insurance','United Kingdom'),('Zurich UK','United Kingdom'),
]

# ---- Demo data (fictional cedants) ---------------------------------------------
DEMO_CEDANTS = [('Demo - Alpha Insurance','Saudi Arabia'),('Demo - Baobab General','Zambia'),
                ('Demo - Crescent Takaful','United Arab Emirates'),('Demo - Delta Re-Assurance','Kenya'),
                ('Demo - Gulf Pearl Insurance','Bahrain')]
# cedant, country, zone, zone name, year, treaty, class, ccy, written line, signed line, EQ, WS, Flood, SRCC, Others
DEMO_ROWS = [
    ('Demo - Alpha Insurance','Saudi Arabia','SA-01','Riyadh',2026,'Quota Share','Property','SAR',0.1,0.08,450e6,0,120e6,80e6,0),
    ('Demo - Alpha Insurance','Saudi Arabia','SA-02','Makkah',2026,'Quota Share','Property','SAR',0.1,0.08,380e6,0,210e6,60e6,0),
    ('Demo - Alpha Insurance','Saudi Arabia','SA-03','Eastern Province',2026,'CAT XL','Property','SAR',0.15,0.125,300e6,90e6,150e6,40e6,0),
    ('Demo - Alpha Insurance','United Arab Emirates','','',2026,'Quota Share','Engineering','AED',0.05,0.05,75e6,40e6,25e6,10e6,5e6),
    ('Demo - Alpha Insurance','Saudi Arabia','SA-01','Riyadh',2025,'Quota Share','Property','SAR',0.1,0.1,400e6,0,100e6,70e6,0),
    ('Demo - Baobab General','Zambia','','',2026,'Quota Share','Property','USD',0.2,0.2,12e6,0,8e6,6e6,1e6),
    ('Demo - Baobab General','Zimbabwe','','',2026,'Quota Share','Property','USD',0.2,0.15,9e6,0,5e6,7e6,0),
    ('Demo - Baobab General','Zambia','','',2025,'First Surplus','Property','USD',0.25,0.25,10e6,0,7e6,5e6,0),
    ('Demo - Crescent Takaful','United Arab Emirates','AE-DXB','Dubai',2026,'Quota Share','Property','AED',0.075,0.075,520e6,310e6,190e6,45e6,0),
    ('Demo - Crescent Takaful','United Arab Emirates','AE-AUH','Abu Dhabi',2026,'Quota Share','Property','AED',0.075,0.075,340e6,200e6,110e6,30e6,0),
    ('Demo - Crescent Takaful','Oman','','',2026,'Risk & CAT XL','Property','OMR',0.1,None,18e6,22e6,15e6,3e6,0),
    ('Demo - Crescent Takaful','Qatar','','',2026,'Quota Share','Energy','QAR',0.05,0.04,95e6,60e6,20e6,12e6,0),
    ('Demo - Delta Re-Assurance','Kenya','KE-NBO','Nairobi',2026,'Quota Share','Property','KES',0.15,0.12,2.5e9,0,1.2e9,0.9e9,0),
    ('Demo - Delta Re-Assurance','Tanzania','','',2026,'Quota Share','Property','USD',0.15,0.15,14e6,3e6,9e6,4e6,0),
    ('Demo - Delta Re-Assurance','Uganda','','',2026,'Quota Share','Agriculture','USD',0.1,0.1,0,0,6e6,0,4e6),
    ('Demo - Gulf Pearl Insurance','Bahrain','','',2026,'Quota Share','Property','BHD',0.125,0.1,60e6,15e6,25e6,20e6,0),
    ('Demo - Gulf Pearl Insurance','Saudi Arabia','SA-03','Eastern Province',2026,'Quota Share','Property','SAR',0.125,0.1,150e6,30e6,90e6,25e6,0),
    ('Demo - Gulf Pearl Insurance','Kuwait','','',2026,'Fac Oblig','Property','USD',0.2,0.2,35e6,8e6,12e6,15e6,0),
    ('Demo - Gulf Pearl Insurance','Jordan','','',2025,'Quota Share','Property','JOD',0.1,0.1,40e6,0,6e6,9e6,0),
]


def build(path, demo=False):
    wb = Workbook()
    st = wb.active
    st.title = "Start"
    dash = wb.create_sheet("Dashboard")
    basis_ws = [wb.create_sheet(nm) for nm, _, _ in BASES]
    fac = wb.create_sheet("Factor Applied")
    mx = wb.create_sheet("Matrix")
    ag = wb.create_sheet("Aggregates")
    ce = wb.create_sheet("Cedants")
    co = wb.create_sheet("Countries")
    ra = wb.create_sheet("Rates")
    li = wb.create_sheet("Lists")
    calc = wb.create_sheet("Calc")
    report_tabs = {"Start", "Dashboard", "Factor Applied", "Matrix"} | {nm for nm, _, _ in BASES}
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = HEADER if ws.title in report_tabs else TEAL

    # ---------------- defined names ---------------------------------------------
    rng = lambda col: f"Aggregates!${col}${FIRST}:${col}${LAST_AGG}"
    for nm, col in [("agg_Status", "A"), ("agg_Cedant", "C"), ("agg_Country", "D"), ("agg_UWYear", "G"),
                    ("agg_Ccy", "J"), ("agg_Written", "L"), ("agg_Signed", "M"), ("agg_Delete", "T"),
                    ("agg_FX", "W")]:
        name(wb, nm, rng(col))
    # 15 USD columns: 100% (X:AB), written share (AC:AG), signed share (AH:AL); 5 perils each
    name(wb, "agg_USD_block", f"Aggregates!$X${FIRST}:$AL${LAST_AGG}")
    name(wb, "lst_Cedants", f"OFFSET(Cedants!$A${FIRST},0,0,MAX(1,COUNTA(Cedants!$A${FIRST}:$A${LAST_CED})),1)")
    name(wb, "lst_Countries", f"OFFSET(Countries!$A${FIRST},0,0,MAX(1,COUNTA(Countries!$A${FIRST}:$A${LAST_CTRY})),1)")
    name(wb, "lst_Ccy", f"OFFSET(Rates!$A${FIRST},0,0,MAX(1,COUNTA(Rates!$A${FIRST}:$A${LAST_RATE})),1)")
    name(wb, "lst_Treaty", f"OFFSET(Lists!$A${FIRST},0,0,MAX(1,COUNTA(Lists!$A${FIRST}:$A${LAST_LIST})),1)")
    name(wb, "lst_Cob", f"OFFSET(Lists!$D${FIRST},0,0,MAX(1,COUNTA(Lists!$D${FIRST}:$D${LAST_LIST})),1)")
    name(wb, "lst_Perils", "Lists!$G$6:$G$10")
    name(wb, "lst_UWYears", "Lists!$J$6:$J$27")
    name(wb, "lst_Basis", "Lists!$M$6:$M$10")
    name(wb, "lst_ViewBy", "Lists!$O$6:$O$7")
    name(wb, "cfg_DbPath", "Start!$C$7")
    name(wb, "cfg_LastPull", "Start!$C$9")
    name(wb, "cfg_LastPush", "Start!$C$10")
    name(wb, "cfg_Factor", "'Factor Applied'!$F$4")
    name(wb, "sel_UWYear", "Dashboard!$C$5")
    name(wb, "crit_UWYear", 'IF(Dashboard!$C$5="All",">0",Dashboard!$C$5)')
    name(wb, "sel_MxBasis", "MATCH(Matrix!$C$4,Lists!$M$6:$M$10,0)")
    name(wb, "sel_MxCol", "CHOOSE(MATCH(Matrix!$C$4,Lists!$M$6:$M$10,0),0,5,10,5,10)+MATCH(Matrix!$C$5,Lists!$G$6:$G$10,0)")
    name(wb, "sel_MxMult", "IF(MATCH(Matrix!$C$4,Lists!$M$6:$M$10,0)>3,'Factor Applied'!$F$4,1)")

    # ================= Aggregates ===============================================
    ncol = AG_NCOL
    title_bar(ag, ncol, "Aggregates by cedant / country",
              "Working copy of the Access table cedant_aggregate. Enter in blue columns, then Push to Access. "
              "Amounts are 100% in original currency, one column per peril (perils are not additive). "
              "Written and signed lines turn them into our share. Type X in 'Delete' to remove a saved row.")
    labels = ["Status", "ID", "Cedant", "Country", "CRESTA zone ID", "Zone name", "UW year", "Treaty type",
              "Class of business", "Currency", "As-at date", "Written line %", "Signed line %",
              "EQ (100%)", "WS (100%)", "Flood (100%)", "SRCC (100%)", "Others (100%)", "Notes",
              "Delete (X)", "row_version", "_fp", "FX (USD per 1)"]
    labels += [f"{p} 100% USD" for p, _ in PERILS]
    labels += [f"{p} written USD" for p, _ in PERILS]
    labels += [f"{p} signed USD" for p, _ in PERILS]
    header(ag, 5, labels)
    for c in range(23, ncol + 1):
        ag.cell(5, c).fill = fill(TEAL)
    for c in (12, 13):
        ag.cell(5, c).fill = fill("2F6F8F")
    ag.cell(4, 1, "Sync state").font = f(8, True, INK_SOFT)
    ag.cell(4, 3, "Inputs (blue) - blank zone / treaty type / class = ALL").font = f(8, True, INK_SOFT)
    ag.cell(4, 12, "Our lines on this treaty").font = f(8, True, INK_SOFT)
    ag.cell(4, 23, "Calculated - do not type here").font = f(8, True, INK_SOFT)
    w = {"A": 26, "B": 7, "C": 28, "D": 22, "E": 12, "F": 18, "G": 8, "H": 18, "I": 18, "J": 9,
         "K": 12, "L": 10, "M": 10, "N": 15, "O": 15, "P": 15, "Q": 15, "R": 15, "S": 24, "T": 8,
         "U": 8, "V": 8, "W": 12}
    for c in range(24, ncol + 1):
        w[L(c)] = 14
    widths(ag, w)
    ag.column_dimensions["U"].hidden = True
    ag.column_dimensions["V"].hidden = True
    blue = f(10, color=INPUT_BLUE)
    style_range(ag, f"A{FIRST}:A{LAST_AGG}", font=f(9, color=INK_SOFT), border=False)
    style_range(ag, f"B{FIRST}:B{LAST_AGG}", font=f(9, color=INK_SOFT), number_format="0", border=False)
    style_range(ag, f"C{FIRST}:T{LAST_AGG}", font=blue, border=False)
    style_range(ag, f"E{FIRST}:E{LAST_AGG}", number_format="@", border=False)
    style_range(ag, f"G{FIRST}:G{LAST_AGG}", number_format="0", border=False)
    style_range(ag, f"K{FIRST}:K{LAST_AGG}", number_format=DATE, border=False)
    style_range(ag, f"L{FIRST}:M{LAST_AGG}", number_format=PCT, border=False)
    style_range(ag, f"N{FIRST}:R{LAST_AGG}", number_format=MONEY, border=False)
    style_range(ag, f"T{FIRST}:T{LAST_AGG}", align=Alignment(horizontal="center"), border=False)
    c100, cwr, csg = USD_100, USD_WRITTEN, USD_SIGNED
    for r in range(FIRST, LAST_AGG + 1):
        ag[f"W{r}"] = (f'=IF(OR($C{r}="",$J{r}="",UPPER($T{r})="X"),"",IFERROR(IF(VLOOKUP($J{r},Rates!$A${FIRST}:$C${LAST_RATE},3,FALSE)>0,'
                       f'VLOOKUP($J{r},Rates!$A${FIRST}:$C${LAST_RATE},3,FALSE),""),""))')
        for k, src in enumerate("NOPQR"):
            ag[f"{c100[k]}{r}"] = f'=IF($W{r}="",0,N({src}{r})*$W{r})'
            ag[f"{cwr[k]}{r}"] = f'={c100[k]}{r}*N($L{r})'
            ag[f"{csg[k]}{r}"] = f'={c100[k]}{r}*N($M{r})'
    style_range(ag, f"W{FIRST}:W{LAST_AGG}", font=f(9), number_format=RATE, bg=CALC_BG, border=False)
    style_range(ag, f"X{FIRST}:{L(ncol)}{LAST_AGG}", font=f(10), number_format=MONEY, bg=CALC_BG, border=False)
    ag.freeze_panes = "E6"
    ag.auto_filter.ref = f"A5:{L(ncol)}{LAST_AGG}"
    data_ref = f"A{FIRST}:A{LAST_AGG}"
    ag.conditional_formatting.add(data_ref, FormulaRule(formula=[f'LEFT($A{FIRST},5)="ERROR"'],
                                  fill=fill("F8D7DA"), font=Font(name=FONT, color="9C0006", bold=True)))
    ag.conditional_formatting.add(data_ref, FormulaRule(formula=[f'OR(LEFT($A{FIRST},3)="NOT",LEFT($A{FIRST},4)="SKIP")'],
                                  fill=fill("FFF3CD"), font=Font(name=FONT, color="7A5B00")))
    ag.conditional_formatting.add(data_ref, FormulaRule(formula=[f'$A{FIRST}="OK"'], font=Font(name=FONT, color="2F8F6B")))
    ag.conditional_formatting.add(f"C{FIRST}:{L(ncol)}{LAST_AGG}", FormulaRule(formula=[f'UPPER($T{FIRST})="X"'],
                                  font=Font(name=FONT, strike=True, color="999999")))
    ag.conditional_formatting.add(f"W{FIRST}:W{LAST_AGG}", FormulaRule(formula=[f'AND($C{FIRST}<>"",$J{FIRST}<>"",$W{FIRST}="",UPPER($T{FIRST})<>"X")'],
                                  fill=fill("FFF3CD")))
    # missing line (amber) and signed > written (red)
    ag.conditional_formatting.add(f"L{FIRST}:M{LAST_AGG}", FormulaRule(
        formula=[f'AND($C{FIRST}<>"",UPPER($T{FIRST})<>"X",ISNUMBER($L{FIRST}),ISNUMBER($M{FIRST}),$M{FIRST}>$L{FIRST})'],
        fill=fill("F8D7DA"), font=Font(name=FONT, color="9C0006", bold=True)))
    ag.conditional_formatting.add(f"L{FIRST}:M{LAST_AGG}", FormulaRule(
        formula=[f'AND($C{FIRST}<>"",UPPER($T{FIRST})<>"X",L{FIRST}="")'], fill=fill("FFF3CD")))

    def dv(ws, ref, **kw):
        d = DataValidation(allow_blank=True, showErrorMessage=True, **kw)
        ws.add_data_validation(d)
        d.add(ref)
        return d
    dv(ag, f"C{FIRST}:C{LAST_AGG}", type="list", formula1="lst_Cedants", errorTitle="Cedant",
       error="Pick a cedant from the Cedants sheet (add new cedants there first).")
    dv(ag, f"D{FIRST}:D{LAST_AGG}", type="list", formula1="lst_Countries", errorTitle="Country",
       error="Pick a country from the Countries sheet.")
    dv(ag, f"G{FIRST}:G{LAST_AGG}", type="whole", operator="between", formula1="1990", formula2="2100",
       errorTitle="UW year", error="Enter an underwriting year, e.g. 2026.")
    dv(ag, f"H{FIRST}:H{LAST_AGG}", type="list", formula1="lst_Treaty")
    dv(ag, f"I{FIRST}:I{LAST_AGG}", type="list", formula1="lst_Cob")
    dv(ag, f"J{FIRST}:J{LAST_AGG}", type="list", formula1="lst_Ccy", errorTitle="Currency",
       error="Pick a currency from the Rates sheet.")
    dv(ag, f"K{FIRST}:K{LAST_AGG}", type="date", operator="greaterThan", formula1="32874",
       errorTitle="As-at date", error="Enter a date.")
    dv(ag, f"L{FIRST}:M{LAST_AGG}", type="decimal", operator="between", formula1="0", formula2="1",
       errorTitle="Line %", error="Enter our line as a percentage between 0% and 100%, e.g. 7.5%.")
    dv(ag, f"N{FIRST}:R{LAST_AGG}", type="decimal", operator="greaterThanOrEqual", formula1="0",
       errorTitle="Aggregate", error="Aggregates are amounts of zero or more.")
    dv(ag, f"T{FIRST}:T{LAST_AGG}", type="list", formula1='"X"')

    if demo:
        for i, row in enumerate(DEMO_ROWS):
            r = FIRST + i
            ced, ctry, zone, zname, yr, tt, cob, ccy, wl, sl, *amts = row
            vals = {"C": ced, "D": ctry, "E": zone or None, "F": zname or None, "G": yr, "H": tt, "I": cob,
                    "J": ccy, "K": date(2026, 9, 30), "L": wl, "M": sl}
            for col, v in vals.items():
                ag[f"{col}{r}"] = v
            for col, v in zip("NOPQR", amts):
                ag[f"{col}{r}"] = v
            ag[f"A{r}"] = "NEW - not yet pushed"

    # ================= Cedants ==================================================
    title_bar(ce, 6, "Cedants", "Pre-loaded with Universe's cedant list. Add more with the Add cedant button, or type a row "
              "and Push lists + cedants. Rename in place (the ID keeps the link); set Active = N instead of deleting.")
    header(ce, 5, ["Cedant name", "Home country", "Active (Y/N)", "ID", "Status"])
    widths(ce, {"A": 36, "B": 24, "C": 11, "D": 7, "E": 40, "F": 4})
    style_range(ce, f"A{FIRST}:C{LAST_CED}", font=blue, border=False)
    style_range(ce, f"D{FIRST}:E{LAST_CED}", font=f(9, color=INK_SOFT), border=False)
    style_range(ce, f"C{FIRST}:C{LAST_CED}", align=Alignment(horizontal="center"), border=False)
    dv(ce, f"B{FIRST}:B{LAST_CED}", type="list", formula1="lst_Countries")
    dv(ce, f"C{FIRST}:C{LAST_CED}", type="list", formula1='"Y,N"')
    ce.freeze_panes = "A6"
    ced_seed = sorted(UNIVERSE_CEDANTS + (DEMO_CEDANTS if demo else []), key=lambda x: x[0].lower())
    for i, (nm, hc) in enumerate(ced_seed):
        ce.cell(FIRST + i, 1, nm); ce.cell(FIRST + i, 2, hc); ce.cell(FIRST + i, 3, "Y")

    # ================= Countries ================================================
    title_bar(co, 7, "Countries", "Reference list (Universe seed plus additional African markets). Add rows to extend.")
    header(co, 5, ["Country", "ISO2", "ISO3", "Region", "ID", "Status"])
    widths(co, {"A": 28, "B": 7, "C": 7, "D": 14, "E": 7, "F": 30, "G": 4})
    style_range(co, f"A{FIRST}:D{LAST_CTRY}", font=blue, border=False)
    style_range(co, f"E{FIRST}:F{LAST_CTRY}", font=f(9, color=INK_SOFT), border=False)
    for i, (nm, a2, a3) in enumerate(COUNTRIES):
        r = FIRST + i
        co.cell(r, 1, nm); co.cell(r, 2, a2); co.cell(r, 3, a3); co.cell(r, 4, REGION[a2])
    co.freeze_panes = "A6"

    # ================= Rates ====================================================
    title_bar(ra, 6, "Currencies and FX rates",
              "Rate = USD per 1 unit of currency. Only USD and fixed pegs are pre-filled; enter the rest. "
              "A new effective date adds history in Access; the latest rate per currency is used.")
    header(ra, 5, ["Currency", "Name", "USD per 1 unit", "Effective date", "Status"])
    widths(ra, {"A": 10, "B": 26, "C": 16, "D": 14, "E": 44, "F": 4})
    style_range(ra, f"A{FIRST}:D{LAST_RATE}", font=blue, border=False)
    style_range(ra, f"C{FIRST}:C{LAST_RATE}", number_format=RATE, border=False)
    style_range(ra, f"D{FIRST}:D{LAST_RATE}", number_format=DATE, border=False)
    style_range(ra, f"E{FIRST}:E{LAST_RATE}", font=f(9, color=INK_SOFT), border=False)
    for i, (code, nm) in enumerate(CURRENCIES):
        r = FIRST + i
        ra.cell(r, 1, code); ra.cell(r, 2, nm)
        if code in PEGS:
            rate, note = PEGS[code]
            ra.cell(r, 3, round(rate, 6)); ra.cell(r, 4, PEG_DATE)
            ra.cell(r, 5, f"{note} - verify before use")
        else:
            ra.cell(r, 5, "NO RATE - enter USD per 1 unit + date")
    ra.conditional_formatting.add(f"E{FIRST}:E{LAST_RATE}", FormulaRule(formula=[f'LEFT($E{FIRST},2)="NO"'],
                                  fill=fill("FFF3CD")))
    ra.conditional_formatting.add(f"E{FIRST}:E{LAST_RATE}", FormulaRule(formula=[f'LEFT($E{FIRST},5)="ERROR"'],
                                  fill=fill("F8D7DA")))
    ra.freeze_panes = "A6"

    # ================= Lists ====================================================
    title_bar(li, 11, "Lists", "Treaty types and classes of business are Universe's predefined lists; Setup loads them into "
              "Access and the Aggregates dropdowns only accept these. Add a row + Push to extend. Perils and UW years are fixed.")
    header(li, 5, ["Treaty type", "Category"], 1)
    header(li, 5, ["Class of business", "Code"], 4)
    header(li, 5, ["Peril", "Meaning"], 7)
    header(li, 5, ["UW year filter"], 10)
    widths(li, {"A": 24, "B": 20, "C": 3, "D": 24, "E": 9, "F": 3, "G": 9, "H": 32, "I": 3, "J": 14, "K": 3, "L": 3, "M": 18, "N": 3, "O": 10})
    style_range(li, f"A{FIRST}:B{LAST_LIST}", font=blue, border=False)
    style_range(li, f"D{FIRST}:E{LAST_LIST}", font=blue, border=False)
    for i, (a, b) in enumerate(TREATY_TYPES):
        li.cell(FIRST + i, 1, a); li.cell(FIRST + i, 2, b)
    for i, (a, b) in enumerate(CLASSES):
        li.cell(FIRST + i, 4, a); li.cell(FIRST + i, 5, b)
    for i, (a, b) in enumerate(PERILS):
        li.cell(FIRST + i, 7, a).font = f(10, True); li.cell(FIRST + i, 8, b).font = f(10)
    li.cell(FIRST, 10, "All").font = f(10, True)
    for i, y in enumerate(range(2015, 2036)):
        li.cell(FIRST + 1 + i, 10, y).font = f(10)
    li.cell(4, 7, "Fixed - matches the five peril columns").font = f(8, True, INK_SOFT)
    header(li, 5, ["Matrix basis"], 13)
    header(li, 5, ["View by"], 15)
    for i, b in enumerate(BASIS_LIST):
        li.cell(FIRST + i, 13, b).font = f(10)
    for i, b in enumerate(["Cedant", "Country"]):
        li.cell(FIRST + i, 15, b).font = f(10)
    li.freeze_panes = "A6"

    # ================= Basis sheets: 100% / written / signed ======================
    N_ROWS = LAST_CED - FIRST + 1                     # 500 (covers 200 countries too)
    TOP, BOT = FIRST + 1, FIRST + N_ROWS              # row 6 = total

    def list_cols(ws, r, s, vw="$B$4"):
        """A name, B home country / ISO2, C active / region - switched by the View-by cell."""
        ws[f"A{r}"] = (f'=IF({vw}="Country",IF(Countries!A{s}="","",Countries!A{s}),'
                       f'IF(Cedants!A{s}="","",Cedants!A{s}))')
        ws[f"B{r}"] = (f'=IF($A{r}="","",IF({vw}="Country",Countries!B{s}&"",'
                       f'IF(Cedants!B{s}="","",Cedants!B{s})))')
        ws[f"C{r}"] = (f'=IF($A{r}="","",IF({vw}="Country",Countries!D{s}&"",'
                       f'IF(Cedants!C{s}="","Y",Cedants!C{s})))')
        ws[f"D{r}"] = (f'=IF($A{r}="","",IF({vw}="Country",COUNTIFS(agg_Country,$A{r},agg_UWYear,crit_UWYear),'
                       f'COUNTIFS(agg_Cedant,$A{r},agg_UWYear,crit_UWYear)))')

    def sumifs(col, r, mult="", vw="$B$4"):
        src = rng(col)
        return (f'=IF($A{r}="","",{mult}IF({vw}="Country",SUMIFS({src},agg_Country,$A{r},agg_UWYear,crit_UWYear),'
                f'SUMIFS({src},agg_Cedant,$A{r},agg_UWYear,crit_UWYear)))')

    def view_controls(ws, extra=""):
        ws["A4"] = "View by"; ws["A4"].font = f(10, True); ws["A4"].alignment = Alignment(horizontal="right")
        ws["B4"] = "Cedant"
        ws["B4"].font = f(11, True, INPUT_BLUE); ws["B4"].fill = fill("FFFF00"); ws["B4"].border = BOX
        ws["B4"].alignment = Alignment(horizontal="center")
        dv(ws, "B4", type="list", formula1="lst_ViewBy")
        ws["D4"] = '="UW year: "&sel_UWYear&"   (change on Dashboard)"' + extra
        ws["D4"].font = f(9, True, TEAL)

    def head_labels(ws, amount_labels, tail):
        header(ws, 5, ["", "", "", "Rows"] + amount_labels + tail)
        ws["A5"] = '=IF($B$4="Country","Country","Cedant")'
        ws["B5"] = '=IF($B$4="Country","ISO2","Home country")'
        ws["C5"] = '=IF($B$4="Country","Region","Active")'

    def finish(ws, last_col, amount_cols, peak_cols, bar_col, bar_color):
        style_range(ws, f"A{FIRST}:{last_col}{FIRST}", font=f(10, True, TOTAL_INK), bg=TOTAL_BG)
        ws[f"A{FIRST}"] = "TOTAL"
        ws[f"D{FIRST}"] = f"=SUM(D{TOP}:D{BOT})"
        ws[f"D{FIRST}"].number_format = "0"
        style_range(ws, f"A{TOP}:C{BOT}", font=f(10), border=False)
        style_range(ws, f"D{TOP}:D{BOT}", font=f(10), number_format='0;-0;"-"', border=False)
        for c in amount_cols + peak_cols:
            style_range(ws, f"{c}{TOP}:{c}{BOT}", font=f(10), number_format=MONEY, border=False)
            ws[f"{c}{FIRST}"].number_format = MONEY
        ws.conditional_formatting.add(f"{bar_col}{TOP}:{bar_col}{BOT}", DataBarRule(
            start_type="num", start_value=0, end_type="max", color=bar_color, showValue=True))
        ws.conditional_formatting.add(f"A{TOP}:{last_col}{BOT}", FormulaRule(formula=["MOD(ROW(),2)=0"], fill=fill("F4F9FA")))
        ws.freeze_panes = f"B{TOP}"

    for ws, (title, desc, start) in zip(basis_ws, BASES):
        cols = USD_COLS[start]
        title_bar(ws, 12, f"{title} (USD)",
                  f"Aggregates at {desc}, in USD at the latest rate on Rates. Switch View by between Cedant and Country. "
                  f"Peak = largest single-peril aggregate (perils are not added together).")
        view_controls(ws)
        head_labels(ws, [p for p, _ in PERILS], ["Peak peril agg", "Peak peril", "Rank"])
        widths(ws, {"A": 34, "B": 20, "C": 13, "D": 7, "E": 16, "F": 16, "G": 16, "H": 16, "I": 16,
                    "J": 17, "K": 10, "L": 7})
        for col in "EFGHI":
            ws[f"{col}{FIRST}"] = f"=SUM({col}{TOP}:{col}{BOT})"
        ws[f"J{FIRST}"] = f"=MAX(E{FIRST}:I{FIRST})"
        ws[f"K{FIRST}"] = f'=IF(N(J{FIRST})=0,"",INDEX(lst_Perils,MATCH(J{FIRST},E{FIRST}:I{FIRST},0)))'
        for r in range(TOP, BOT + 1):
            s = r - 1
            list_cols(ws, r, s)
            for k, col in enumerate("EFGHI"):
                ws[f"{col}{r}"] = sumifs(cols[k], r)
            ws[f"J{r}"] = f'=IF($A{r}="","",MAX(E{r}:I{r}))'
            ws[f"K{r}"] = f'=IF(N(J{r})=0,"",INDEX(lst_Perils,MATCH(J{r},E{r}:I{r},0)))'
            ws[f"L{r}"] = f'=IF(N(J{r})=0,"",RANK(J{r},$J${TOP}:$J${BOT}))'
        finish(ws, "L", list("EFGHI"), ["J"], "J", "7FB8A4")
        style_range(ws, f"K{TOP}:L{BOT}", font=f(10), align=Alignment(horizontal="center"), border=False)
        ws[f"K{FIRST}"].alignment = Alignment(horizontal="center")

    # ================= Factor Applied ==============================================
    title_bar(fac, 17, "Factor applied to written and signed share (USD)",
              "Each written-share and signed-share aggregate multiplied by the factor in F4 (default 0.35%). "
              "Same View by and UW year filter as the other share sheets.")
    view_controls(fac)
    fac["D4"] = "Factor"; fac["D4"].font = f(10, True); fac["D4"].alignment = Alignment(horizontal="right")
    fac["F4"] = DEFAULT_FACTOR
    fac["F4"].number_format = "0.00%"
    fac["F4"].font = f(11, True, INPUT_BLUE); fac["F4"].fill = fill("FFFF00"); fac["F4"].border = BOX
    fac["F4"].alignment = Alignment(horizontal="center")
    fac["F4"].comment = Comment("Factor applied to written and signed share aggregates. "
                                "Set by the user: 0.35%. Change it here; every figure on this sheet, "
                                "the Dashboard and the Matrix follows.", "Aggregate Tracker")
    dv(fac, "F4", type="decimal", operator="between", formula1="0", formula2="1",
       errorTitle="Factor", error="Enter a percentage between 0% and 100%.")
    fac["H4"] = '="UW year: "&sel_UWYear'; fac["H4"].font = f(9, True, TEAL)
    head_labels(fac, [f"Written {p}" for p, _ in PERILS] + [f"Signed {p}" for p, _ in PERILS],
                ["Peak written", "Peak signed", "Rank (signed)"])
    fw = {"A": 34, "B": 20, "C": 13, "D": 7, "Q": 9}
    for c in range(5, 17):
        fw[L(c)] = 14
    widths(fac, fw)
    for c in range(5, 10):
        fac.cell(5, c).fill = fill("2F6F8F")
    for c in range(10, 15):
        fac.cell(5, c).fill = fill(HEADER)
    wcols, scols = list("EFGHI"), list("JKLMN")
    for col in wcols + scols:
        fac[f"{col}{FIRST}"] = f"=SUM({col}{TOP}:{col}{BOT})"
    fac[f"O{FIRST}"] = f"=MAX(E{FIRST}:I{FIRST})"
    fac[f"P{FIRST}"] = f"=MAX(J{FIRST}:N{FIRST})"
    for r in range(TOP, BOT + 1):
        s = r - 1
        list_cols(fac, r, s)
        for k in range(5):
            fac[f"{wcols[k]}{r}"] = sumifs(USD_WRITTEN[k], r, mult="cfg_Factor*")
            fac[f"{scols[k]}{r}"] = sumifs(USD_SIGNED[k], r, mult="cfg_Factor*")
        fac[f"O{r}"] = f'=IF($A{r}="","",MAX(E{r}:I{r}))'
        fac[f"P{r}"] = f'=IF($A{r}="","",MAX(J{r}:N{r}))'
        fac[f"Q{r}"] = f'=IF(N(P{r})=0,"",RANK(P{r},$P${TOP}:$P${BOT}))'
    finish(fac, "Q", wcols + scols, ["O", "P"], "P", "7FB8A4")
    style_range(fac, f"Q{TOP}:Q{BOT}", font=f(10), align=Alignment(horizontal="center"), border=False)

    # ================= Calc (hidden): signed share per cedant / country ============
    calc.sheet_state = "hidden"
    calc["A1"] = "Helper for the Dashboard 'largest' columns - signed share USD by cedant and by country."
    header(calc, 5, ["Cedant"] + [p for p, _ in PERILS] + ["", "Country"] + [p for p, _ in PERILS])
    for r in range(FIRST, LAST_CED + 1):
        calc[f"A{r}"] = f'=IF(Cedants!A{r}="","",Cedants!A{r})'
        for k, col in enumerate("BCDEF"):
            calc[f"{col}{r}"] = f'=IF($A{r}="",0,SUMIFS({rng(USD_SIGNED[k])},agg_Cedant,$A{r},agg_UWYear,crit_UWYear))'
    for r in range(FIRST, LAST_CTRY + 1):
        calc[f"H{r}"] = f'=IF(Countries!A{r}="","",Countries!A{r})'
        for k, col in enumerate("IJKLM"):
            calc[f"{col}{r}"] = f'=IF($H{r}="",0,SUMIFS({rng(USD_SIGNED[k])},agg_Country,$H{r},agg_UWYear,crit_UWYear))'

    # ================= Matrix ===================================================
    MX_HDR, MX_TOP, MX_BOT = 7, 8, 207
    title_bar(mx, 23, "Cedant x country matrix (USD)",
              "Pick a basis and a peril, and change any country in row 7 (dropdowns). 'Other countries' = everything not shown.")
    mx["B4"] = "Basis"; mx["B4"].font = f(10, True)
    mx["C4"] = "Signed share"
    mx["B5"] = "Peril"; mx["B5"].font = f(10, True)
    mx["C5"] = "EQ"
    for c in ("C4", "C5"):
        mx[c].font = f(11, True, INPUT_BLUE); mx[c].fill = fill("FFFF00"); mx[c].border = BOX
    dv(mx, "C4", type="list", formula1="lst_Basis")
    dv(mx, "C5", type="list", formula1="lst_Perils")
    mx["E4"] = '=C4&" - "&C5&IF(sel_MxBasis>3,"  (factor "&TEXT(cfg_Factor,"0.00%")&")","")'
    mx["E4"].font = f(10, True, TEAL)
    mx["E5"] = '="UW year: "&sel_UWYear&"   (change on Dashboard)"'; mx["E5"].font = f(9, True, TEAL)
    labels = ["Cedant"] + MATRIX_COUNTRIES + ["Other countries", "Total"]
    header(mx, MX_HDR, labels)
    for i in range(20):
        mx.cell(MX_HDR, 2 + i).fill = fill(TEAL)
    dv(mx, f"B{MX_HDR}:U{MX_HDR}", type="list", formula1="lst_Countries")
    mx.column_dimensions["A"].width = 32
    mx.column_dimensions["C"].width = 16
    for c in range(2, 24):
        if c != 3:
            mx.column_dimensions[L(c)].width = 13
    mx.row_dimensions[MX_HDR].height = 42
    for r in range(MX_TOP, MX_BOT + 1):
        s = r - MX_TOP + FIRST
        mx[f"A{r}"] = f'=IF(Cedants!A{s}="","",Cedants!A{s})'
        for c in range(2, 22):
            cl = L(c)
            mx[f"{cl}{r}"] = (f'=IF(OR($A{r}="",{cl}${MX_HDR}=""),"",sel_MxMult*SUMIFS(INDEX(agg_USD_block,0,sel_MxCol),'
                              f'agg_Cedant,$A{r},agg_Country,{cl}${MX_HDR},agg_UWYear,crit_UWYear))')
        mx[f"W{r}"] = f'=IF($A{r}="","",sel_MxMult*SUMIFS(INDEX(agg_USD_block,0,sel_MxCol),agg_Cedant,$A{r},agg_UWYear,crit_UWYear))'
        mx[f"V{r}"] = f'=IF($A{r}="","",W{r}-SUM(B{r}:U{r}))'
    TOT = MX_BOT + 1
    mx.cell(TOT, 1, "TOTAL")
    for c in range(2, 24):
        cl = L(c)
        mx[f"{cl}{TOT}"] = f"=SUM({cl}{MX_TOP}:{cl}{MX_BOT})"
    style_range(mx, f"A{MX_TOP}:A{MX_BOT}", font=f(10), border=False)
    style_range(mx, f"B{MX_TOP}:W{MX_BOT}", font=f(10), number_format=MONEY, border=False)
    style_range(mx, f"V{MX_TOP}:V{MX_BOT}", font=f(10, color=INK_SOFT), number_format=MONEY, bg=CALC_BG, border=False)
    style_range(mx, f"W{MX_TOP}:W{MX_BOT}", font=f(10, True), number_format=MONEY, border=False)
    style_range(mx, f"A{TOT}:W{TOT}", font=f(10, True, TOTAL_INK), number_format=MONEY, bg=TOTAL_BG)
    mx.conditional_formatting.add(f"B{MX_TOP}:U{MX_BOT}", DataBarRule(start_type="num", start_value=0, end_type="max",
                                                                      color="9CC9DA", showValue=True))
    mx.freeze_panes = f"B{MX_TOP}"

    # ================= Dashboard ================================================
    title_bar(dash, 8, "Dashboard", "All figures in USD at the latest FX rate. Change the UW year filter here; "
              "every report sheet follows it.")
    widths(dash, {"A": 3, "B": 36, "C": 19, "D": 40, "E": 19, "F": 30, "G": 19, "H": 3})
    dash["B5"] = "UW year filter"; dash["B5"].font = f(10, True)
    dash["C5"] = "All"
    dash["C5"].font = f(11, True, INPUT_BLUE); dash["C5"].fill = fill("FFFF00"); dash["C5"].border = BOX
    dash["C5"].alignment = Alignment(horizontal="center")
    dash["D5"] = "Pick All or a single underwriting year"; dash["D5"].font = f(9, italic=True, color=INK_SOFT)
    dv(dash, "C5", type="list", formula1="lst_UWYears")

    ced_rng = f"Cedants!$A${FIRST}:$A${LAST_CED}"
    ctry_rng = f"Countries!$A${FIRST}:$A${LAST_CTRY}"
    kpis = [("Aggregate rows", '=COUNTIFS(agg_Cedant,"<>",agg_UWYear,crit_UWYear)', "#,##0"),
            ("Cedants with aggregates",
             f'=SUMPRODUCT(({ced_rng}<>"")*(COUNTIFS(agg_Cedant,{ced_rng},agg_UWYear,crit_UWYear)>0))', "#,##0"),
            ("Countries with aggregates",
             f'=SUMPRODUCT(({ctry_rng}<>"")*(COUNTIFS(agg_Country,{ctry_rng},agg_UWYear,crit_UWYear)>0))', "#,##0"),
            ("Factor (set on Factor Applied)", "=cfg_Factor", "0.00%")]
    header(dash, 7, ["Portfolio", "Value"], 2)
    for i, (lab, fx, nf) in enumerate(kpis):
        r = 8 + i
        dash.cell(r, 2, lab); dash.cell(r, 3, fx)
        style_range(dash, f"B{r}:C{r}", font=f(10))
        dash.cell(r, 3).number_format = nf
        dash.cell(r, 3).font = f(11, True)

    live = '(agg_Cedant<>"")*(UPPER(agg_Delete)<>"X")'
    header(dash, 7, ["Data checks (all years)", "Rows"], 4)
    checks = [
        ("Rows with no FX rate (excluded from USD)", f'=SUMPRODUCT({live}*(agg_Ccy<>"")*(agg_FX=""))'),
        ("Rows with no written line (written share = 0)", f'=SUMPRODUCT({live}*(agg_Written=""))'),
        ("Rows with no signed line (signed share = 0)", f'=SUMPRODUCT({live}*(agg_Signed=""))'),
        ("Rows with signed line above written line",
         f'=SUMPRODUCT({live}*(agg_Signed<>"")*(agg_Written<>"")*(agg_Signed>agg_Written))'),
        ("Rows with a cedant/country not on the lists",
         f'=SUMPRODUCT((agg_Cedant<>"")*(COUNTIF({ced_rng},agg_Cedant)=0))'
         f'+SUMPRODUCT((agg_Country<>"")*(COUNTIF({ctry_rng},agg_Country)=0))'),
        ("Rows not OK in Access (new or error)", '=SUMPRODUCT((agg_Cedant<>"")*(agg_Status<>"OK"))'),
    ]
    for i, (lab, fx) in enumerate(checks):
        r = 8 + i
        dash.cell(r, 4, lab); dash.cell(r, 5, fx)
        style_range(dash, f"D{r}:E{r}", font=f(10))
        dash.cell(r, 5).number_format = "#,##0"
        dash.cell(r, 5).font = f(11, True)
    dash.conditional_formatting.add("E8:E13", CellIsRule(operator="greaterThan", formula=["0"],
                                    fill=fill("FFF3CD"), font=Font(name=FONT, bold=True, color="7A5B00")))

    header(dash, 7, ["Sync", ""], 6)
    dash["F8"] = "Last pull from Access"; dash["G8"] = "=IF(cfg_LastPull=\"\",\"never\",cfg_LastPull)"
    dash["F9"] = "Last push to Access"; dash["G9"] = "=IF(cfg_LastPush=\"\",\"never\",cfg_LastPush)"
    style_range(dash, "F8:G9", font=f(10), number_format=STAMP)

    PT = 16   # peril table header row
    header(dash, PT, ["Peril", "100% (USD)", "Written share", "Signed share", "Written x factor",
                      "Signed x factor"], 2)
    dash.cell(PT - 1, 2, "Aggregates by peril and basis").font = f(12, True, TEAL)
    for i, (code, meaning) in enumerate(PERILS):
        r = PT + 1 + i
        dash.cell(r, 2, f"{code} - {meaning}")
        dash.cell(r, 3, f"=SUMIFS({rng(USD_100[i])},agg_UWYear,crit_UWYear)")
        dash.cell(r, 4, f"=SUMIFS({rng(USD_WRITTEN[i])},agg_UWYear,crit_UWYear)")
        dash.cell(r, 5, f"=SUMIFS({rng(USD_SIGNED[i])},agg_UWYear,crit_UWYear)")
        dash.cell(r, 6, f"=D{r}*cfg_Factor")
        dash.cell(r, 7, f"=E{r}*cfg_Factor")
        style_range(dash, f"B{r}:G{r}", font=f(10), number_format=MONEY)
        dash.cell(r, 5).font = f(10, True)
    p1, p5 = PT + 1, PT + 5
    rp = PT + 6
    dash.cell(rp, 2, "Peak single-peril aggregate")
    for c in "CDEFG":
        dash[f"{c}{rp}"] = f"=MAX({c}{p1}:{c}{p5})"
    style_range(dash, f"B{rp}:G{rp}", font=f(10, True, TOTAL_INK), bg=TOTAL_BG, number_format=MONEY)
    dash.cell(rp + 1, 2, "Peak peril")
    for c in "CDEFG":
        dash[f"{c}{rp + 1}"] = f'=IF(N({c}{rp})=0,"-",INDEX(lst_Perils,MATCH({c}{rp},{c}{p1}:{c}{p5},0)))'
    style_range(dash, f"B{rp + 1}:G{rp + 1}", font=f(10, True, TOTAL_INK), bg=TOTAL_BG,
                align=Alignment(horizontal="right"))
    dash[f"B{rp + 1}"].alignment = Alignment(horizontal="left")

    LT = rp + 4
    dash.cell(LT - 1, 2, "Largest exposures on our signed share").font = f(12, True, TEAL)
    header(dash, LT, ["Peril", "Largest cedant", "Signed share (USD)", "Largest country", "Signed share (USD)"], 2)
    dash.merge_cells(start_row=LT, start_column=6, end_row=LT, end_column=7)
    for i, (code, meaning) in enumerate(PERILS):
        r = LT + 1 + i
        cc = f"Calc!${'BCDEF'[i]}${FIRST}:${'BCDEF'[i]}${LAST_CED}"
        kc = f"Calc!${'IJKLM'[i]}${FIRST}:${'IJKLM'[i]}${LAST_CTRY}"
        dash.cell(r, 2, f"{code} - {meaning}")
        dash.cell(r, 3, f'=IF(MAX({cc})=0,"-",INDEX(Calc!$A${FIRST}:$A${LAST_CED},MATCH(MAX({cc}),{cc},0)))')
        dash.cell(r, 4, f"=MAX({cc})")
        dash.cell(r, 5, f'=IF(MAX({kc})=0,"-",INDEX(Calc!$H${FIRST}:$H${LAST_CTRY},MATCH(MAX({kc}),{kc},0)))')
        dash.cell(r, 6, f"=MAX({kc})")
        style_range(dash, f"B{r}:F{r}", font=f(10))
        for c in (4, 6):
            dash.cell(r, c).number_format = MONEY
    nr = LT + 7
    dash[f"B{nr}"] = ("100% = the cedant's full aggregate. Written / signed share = 100% x our written / signed line on "
                      "each row. Factor = share x the factor on Factor Applied. Perils are never added together: the "
                      "same sum insured is usually exposed to several perils, so the single figure is the peak peril.")
    dash[f"B{nr}"].font = f(9, italic=True, color=INK_SOFT)
    dash.merge_cells(f"B{nr}:G{nr + 2}")
    dash[f"B{nr}"].alignment = Alignment(wrap_text=True, vertical="top")

    # ================= Start ====================================================
    title_bar(st, 6, "Start", "Excel front end over an Access database. Excel = screens; Access = system of record.")
    widths(st, {"A": 3, "B": 30, "C": 42, "D": 70, "E": 3, "F": 3})
    st["B5"] = "Database"; st["B5"].font = f(12, True, TEAL)
    st["B7"] = "Database path (optional)"; st["B7"].font = f(10, True)
    st["C7"].fill = fill("FFFF00"); st["C7"].font = f(10, color=INPUT_BLUE); st["C7"].border = BOX
    st["D7"] = ("Leave blank to keep AggregateTracker.accdb next to this workbook. Use a local or network-drive "
                "path (not a OneDrive web link), e.g. S:\\Underwriting\\AggregateTracker.accdb")
    st["D7"].font = f(9, italic=True, color=INK_SOFT); st["D7"].alignment = Alignment(wrap_text=True, vertical="top")
    st.row_dimensions[7].height = 28
    st["B9"] = "Last pull from Access"; st["B10"] = "Last push to Access"
    for c in ("C9", "C10"):
        st[c].number_format = STAMP; st[c].font = f(10); st[c].border = BOX
    st["D9"] = "Written by the macros"; st["D9"].font = f(9, italic=True, color=INK_SOFT)

    st["B12"] = "Actions"; st["B12"].font = f(12, True, TEAL)
    actions = [
        (13, "Setup_Database", "First run: creates the Access file, builds tables/views (migrations), loads Universe's "
                               "predefined treaty types, classes and cedants, then this workbook's lists and any "
                               "aggregate rows. Re-run any time: it only adds what is missing."),
        (15, "Pull_All", "Reload every sheet from Access (do this before you start editing)."),
        (17, "Push_Masters", "Save Countries, Rates, Lists and Cedants. Upserts by name/code; nothing is deleted."),
        (19, "Push_Aggregates", "Validate and save the Aggregates sheet in one transaction. Blocks if any row fails "
                                "or if someone else changed a row since your pull."),
        (21, "Open_Database", "Open the .accdb in Access (tables, audit_log and the vw_* reporting queries)."),
        (23, "Add_Cedant", "Add one cedant straight to Access (name + home country); it appears in the "
                           "Aggregates dropdown at once."),
    ]
    for r, macro, text in actions:
        st.cell(r, 3, f"macro: {macro}").font = f(8, color="AAAAAA")
        st.cell(r, 4, text).font = f(10)
        st.cell(r, 4).alignment = Alignment(wrap_text=True, vertical="top")
        st.row_dimensions[r].height = 30
    st["B13"] = "Buttons appear in column C once the macros are installed (see Install below)."
    st["B13"].font = f(8, italic=True, color=INK_SOFT); st["B13"].alignment = Alignment(wrap_text=True, vertical="top")

    st["B24"] = "Install (once)"; st["B24"].font = f(12, True, TEAL)
    install = [
        "Option A - one click: put this workbook, Build-Workbook.ps1 and the vba folder together, right-click "
        "Build-Workbook.ps1 > Run with PowerShell. It writes Aggregate_Tracker.xlsm with the macros and buttons.",
        "Option B - by hand: Alt+F11 > File > Import File > the seven .bas files in the vba folder. Save as "
        ".xlsm. Then Alt+F8 > InstallButtons > Run.",
        "Then click '1. Setup / create database'. Needs Access, or the free Microsoft Access Database Engine, "
        "with the same 32/64-bit as Excel.",
    ]
    for i, t in enumerate(install):
        r = 25 + i
        st.cell(r, 2, f"Step {i + 1}").font = f(10, True)
        st.cell(r, 3, t).font = f(10)
        st.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        st.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        st.row_dimensions[r].height = 30

    st["B29"] = "Daily workflow"; st["B29"].font = f(12, True, TEAL)
    flow = ["Pull everything from Access.",
            "New cedant? Click Add cedant (or type it on Cedants, then Push reference lists + cedants).",
            "Enter or edit aggregates on Aggregates (blue columns). To remove a saved row type X in Delete (X) - "
            "do not delete the Excel row.",
            "Push aggregates. Fix any rows marked ERROR in the Status column and push again.",
            "Read the results on Dashboard, 100% Share, Written Share, Signed Share, Factor Applied and Matrix."]
    for i, t in enumerate(flow):
        st.cell(30 + i, 2, f"{i + 1}."); st.cell(30 + i, 3, t)
        st.merge_cells(start_row=30 + i, start_column=3, end_row=30 + i, end_column=4)
        st.cell(30 + i, 2).font = f(10, True); st.cell(30 + i, 3).font = f(10)
        st.cell(30 + i, 2).alignment = Alignment(horizontal="right")

    st["B36"] = "Example aggregate row"; st["B36"].font = f(12, True, TEAL)
    example = [
        ("Cedant", "Tawuniya", "Must exist on Cedants"),
        ("Country", "Saudi Arabia", "Must exist on Countries"),
        ("Zone ID", "SA-01", "Optional - blank = ALL (country total)"),
        ("UW year", 2026, "Whole year 1990-2100"),
        ("Treaty type", "Quota Share", "Optional - blank = ALL"),
        ("Class", "Property", "Optional - blank = ALL"),
        ("Ccy", "SAR", "Must exist on Rates; amounts are in this currency"),
        ("Written line %", 0.10, "Our written line on this treaty, 0-100%. Blank = written share shows 0"),
        ("Signed line %", 0.08, "Our signed line, 0-100%, not above the written line. Blank until signed"),
        ("EQ", 450000000, "100% aggregate exposed to earthquake, >= 0"),
        ("WS", 0, "100% aggregate exposed to windstorm"),
        ("Flood", 120000000, "100% aggregate exposed to flood"),
        ("SRCC", 80000000, "100% aggregate exposed to SRCC"),
        ("Others", 0, "Any other peril"),
    ]
    header(st, 37, ["Column", "Example value", "Rule"], 2)
    for i, (h, v, rule) in enumerate(example):
        r = 38 + i
        st.cell(r, 2, h).font = f(10, True)
        c = st.cell(r, 3, v); c.font = f(10, color=INPUT_BLUE)
        if h.endswith("%"):
            c.number_format = PCT
        elif isinstance(v, (int, float)) and i >= 9:
            c.number_format = MONEY
        c.alignment = Alignment(horizontal="left")
        st.cell(r, 4, rule).font = f(10)
        style_range(st, f"B{r}:D{r}")
    r = 38 + len(example) + 1
    st[f"B{r}"] = ("One row per cedant + country + zone + UW year + treaty type + class + currency (unique in Access). "
                   "Signed share = 450m x 8% = 36m SAR; at the 0.35% factor = 126k SAR.")
    st[f"B{r}"].font = f(9, italic=True, color=INK_SOFT)

    r += 2
    st[f"B{r}"] = "Colour legend"; st[f"B{r}"].font = f(12, True, TEAL)
    legend = [("Blue text", "Input you type", None, INPUT_BLUE),
              ("Yellow cell", "Setting / filter", "FFFF00", None),
              ("Grey-blue cell", "Formula - do not type", CALC_BG, None),
              ("Status = ERROR", "Row refused on push - read the message", "F8D7DA", None)]
    for i, (a, b, bg, fg) in enumerate(legend):
        rr = r + 1 + i
        st[f"B{rr}"] = a; st[f"C{rr}"] = b
        st[f"B{rr}"].font = f(10, color=fg or INK); st[f"C{rr}"].font = f(10)
        if bg:
            st[f"B{rr}"].fill = fill(bg)

    r += len(legend) + 2
    st[f"B{r}"] = "How this mirrors Universe"; st[f"B{r}"].font = f(12, True, TEAL)
    header(st, r + 1, ["Universe", "Here", "What it does"], 2)
    arch = [("client (React screens)", "Excel sheets + AggMacros.bas", "Entry grids, dashboards, buttons"),
            ("server/src/services, validation", "AggService.bas", "Validation, push/pull, transactions"),
            ("modules/*/repositories", "AggRepo.bas", "Every SQL statement"),
            ("db/pool.js, withTransaction.js", "AggDb.bas", "ADO connection, parameterised commands"),
            ("db/migrations + runMigrations.js", "AggMigrate.bas + schema_migrations", "Numbered forward-only schema changes"),
            ("startup/ensureReferenceData.js", "AggSeed.bas + Lists / Cedants sheets", "Predefined treaty types, classes, cedants"),
            ("contract_cresta_data", "cedant_aggregate", "Peril aggregates by country / CRESTA zone"),
            ("written / signed line (signed <= written)", "written_line, signed_line", "Our share of each aggregate row"),
            ("ref_exchange_rate (rate_to_usd)", "ref_exchange_rate + vw_fx_latest", "Latest rate per currency"),
            ("db/optimisticLock.js", "row_version column", "Refuses saves over someone else's change"),
            ("audit_log", "audit_log", "Who changed what, when")]
    for i, row in enumerate(arch):
        rr = r + 2 + i
        for j2, v in enumerate(row):
            st.cell(rr, 2 + j2, v).font = f(10)
        style_range(st, f"B{rr}:D{rr}")

    wb.active = 0
    wb.save(path)


if __name__ == "__main__":
    out = sys.argv[1]
    build(f"{out}/Aggregate_Tracker.xlsx", demo=False)
    build(f"{out}/Aggregate_Tracker_DEMO.xlsx", demo=True)
    print("built")

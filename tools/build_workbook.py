"""Generate Aggregate_Tracker.xlsx (clean) and Aggregate_Tracker_DEMO.xlsx.

Layout contract with the VBA modules (AggConfig.bas):
  header row 5, data from row 6 on every grid sheet;
  Aggregates A..T = inputs + sync state, U..Z = formulas.
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
STAMP = 'dd-mmm-yyyy hh:mm'

FIRST, LAST_AGG, LAST_CED, LAST_CTRY, LAST_RATE, LAST_LIST = 6, 5005, 505, 205, 105, 105

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
# cedant, country, zone, zone name, year, treaty, class, ccy, EQ, WS, Flood, SRCC, Others
DEMO_ROWS = [
    ('Demo - Alpha Insurance','Saudi Arabia','SA-01','Riyadh',2026,'Quota Share','Property','SAR',450e6,0,120e6,80e6,0),
    ('Demo - Alpha Insurance','Saudi Arabia','SA-02','Makkah',2026,'Quota Share','Property','SAR',380e6,0,210e6,60e6,0),
    ('Demo - Alpha Insurance','Saudi Arabia','SA-03','Eastern Province',2026,'CAT XL','Property','SAR',300e6,90e6,150e6,40e6,0),
    ('Demo - Alpha Insurance','United Arab Emirates','','',2026,'Quota Share','Engineering','AED',75e6,40e6,25e6,10e6,5e6),
    ('Demo - Alpha Insurance','Saudi Arabia','SA-01','Riyadh',2025,'Quota Share','Property','SAR',400e6,0,100e6,70e6,0),
    ('Demo - Baobab General','Zambia','','',2026,'Quota Share','Property','USD',12e6,0,8e6,6e6,1e6),
    ('Demo - Baobab General','Zimbabwe','','',2026,'Quota Share','Property','USD',9e6,0,5e6,7e6,0),
    ('Demo - Baobab General','Zambia','','',2025,'First Surplus','Property','USD',10e6,0,7e6,5e6,0),
    ('Demo - Crescent Takaful','United Arab Emirates','AE-DXB','Dubai',2026,'Quota Share','Property','AED',520e6,310e6,190e6,45e6,0),
    ('Demo - Crescent Takaful','United Arab Emirates','AE-AUH','Abu Dhabi',2026,'Quota Share','Property','AED',340e6,200e6,110e6,30e6,0),
    ('Demo - Crescent Takaful','Oman','','',2026,'Risk & CAT XL','Property','OMR',18e6,22e6,15e6,3e6,0),
    ('Demo - Crescent Takaful','Qatar','','',2026,'Quota Share','Energy','QAR',95e6,60e6,20e6,12e6,0),
    ('Demo - Delta Re-Assurance','Kenya','KE-NBO','Nairobi',2026,'Quota Share','Property','KES',2.5e9,0,1.2e9,0.9e9,0),
    ('Demo - Delta Re-Assurance','Tanzania','','',2026,'Quota Share','Property','USD',14e6,3e6,9e6,4e6,0),
    ('Demo - Delta Re-Assurance','Uganda','','',2026,'Quota Share','Agriculture','USD',0,0,6e6,0,4e6),
    ('Demo - Gulf Pearl Insurance','Bahrain','','',2026,'Quota Share','Property','BHD',60e6,15e6,25e6,20e6,0),
    ('Demo - Gulf Pearl Insurance','Saudi Arabia','SA-03','Eastern Province',2026,'Quota Share','Property','SAR',150e6,30e6,90e6,25e6,0),
    ('Demo - Gulf Pearl Insurance','Kuwait','','',2026,'Fac Oblig','Property','USD',35e6,8e6,12e6,15e6,0),
    ('Demo - Gulf Pearl Insurance','Jordan','','',2025,'Quota Share','Property','JOD',40e6,0,6e6,9e6,0),
]


def build(path, demo=False):
    wb = Workbook()
    st = wb.active
    st.title = "Start"
    dash = wb.create_sheet("Dashboard")
    byc = wb.create_sheet("By Cedant")
    byk = wb.create_sheet("By Country")
    mx = wb.create_sheet("Matrix")
    ag = wb.create_sheet("Aggregates")
    ce = wb.create_sheet("Cedants")
    co = wb.create_sheet("Countries")
    ra = wb.create_sheet("Rates")
    li = wb.create_sheet("Lists")
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = HEADER if ws.title in ("Start", "Dashboard", "By Cedant", "By Country", "Matrix") else TEAL

    # ---------------- defined names ---------------------------------------------
    rng = lambda col: f"Aggregates!${col}${FIRST}:${col}${LAST_AGG}"
    for nm, col in [("agg_Status", "A"), ("agg_Cedant", "C"), ("agg_Country", "D"), ("agg_UWYear", "G"),
                    ("agg_Ccy", "J"), ("agg_Delete", "R"), ("agg_FX", "U"), ("agg_EQ_USD", "V"),
                    ("agg_WS_USD", "W"), ("agg_Flood_USD", "X"), ("agg_SRCC_USD", "Y"), ("agg_Others_USD", "Z")]:
        name(wb, nm, rng(col))
    name(wb, "agg_USD_block", f"Aggregates!$V${FIRST}:$Z${LAST_AGG}")
    name(wb, "lst_Cedants", f"OFFSET(Cedants!$A${FIRST},0,0,MAX(1,COUNTA(Cedants!$A${FIRST}:$A${LAST_CED})),1)")
    name(wb, "lst_Countries", f"OFFSET(Countries!$A${FIRST},0,0,MAX(1,COUNTA(Countries!$A${FIRST}:$A${LAST_CTRY})),1)")
    name(wb, "lst_Ccy", f"OFFSET(Rates!$A${FIRST},0,0,MAX(1,COUNTA(Rates!$A${FIRST}:$A${LAST_RATE})),1)")
    name(wb, "lst_Treaty", f"OFFSET(Lists!$A${FIRST},0,0,MAX(1,COUNTA(Lists!$A${FIRST}:$A${LAST_LIST})),1)")
    name(wb, "lst_Cob", f"OFFSET(Lists!$D${FIRST},0,0,MAX(1,COUNTA(Lists!$D${FIRST}:$D${LAST_LIST})),1)")
    name(wb, "lst_Perils", "Lists!$G$6:$G$10")
    name(wb, "lst_UWYears", "Lists!$J$6:$J$27")
    name(wb, "cfg_DbPath", "Start!$C$7")
    name(wb, "cfg_LastPull", "Start!$C$9")
    name(wb, "cfg_LastPush", "Start!$C$10")
    name(wb, "sel_UWYear", "Dashboard!$C$5")
    name(wb, "crit_UWYear", 'IF(Dashboard!$C$5="All",">0",Dashboard!$C$5)')
    name(wb, "sel_PerilIdx", "MATCH(Matrix!$C$5,Lists!$G$6:$G$10,0)")

    # ================= Aggregates ===============================================
    ncol = 26
    title_bar(ag, ncol, "Aggregates by cedant / country",
              "Working copy of the Access table cedant_aggregate. Enter in blue columns, then Push to Access. "
              "Amounts in original currency, one column per peril (perils are not additive). Type X in 'Delete' to remove a saved row.")
    labels = ["Status", "ID", "Cedant", "Country", "CRESTA zone ID", "Zone name", "UW year", "Treaty type",
              "Class of business", "Currency", "As-at date", "EQ", "WS", "Flood", "SRCC", "Others", "Notes",
              "Delete (X)", "row_version", "_fp", "FX (USD per 1)", "EQ (USD)", "WS (USD)", "Flood (USD)",
              "SRCC (USD)", "Others (USD)"]
    header(ag, 5, labels)
    for c in range(21, 27):
        ag.cell(5, c).fill = fill(TEAL)
    ag.cell(4, 1, "Sync state").font = f(8, True, INK_SOFT)
    ag.cell(4, 3, "Inputs (blue) - blank zone / treaty type / class = ALL").font = f(8, True, INK_SOFT)
    ag.cell(4, 21, "Calculated - do not type here").font = f(8, True, INK_SOFT)
    widths(ag, {"A": 26, "B": 7, "C": 28, "D": 22, "E": 12, "F": 18, "G": 8, "H": 18, "I": 18, "J": 9,
                "K": 12, "L": 15, "M": 15, "N": 15, "O": 15, "P": 15, "Q": 24, "R": 8, "S": 8, "T": 8,
                "U": 12, "V": 15, "W": 15, "X": 15, "Y": 15, "Z": 15})
    ag.column_dimensions["S"].hidden = True
    ag.column_dimensions["T"].hidden = True
    blue = f(10, color=INPUT_BLUE)
    style_range(ag, f"A{FIRST}:A{LAST_AGG}", font=f(9, color=INK_SOFT), border=False)
    style_range(ag, f"B{FIRST}:B{LAST_AGG}", font=f(9, color=INK_SOFT), number_format="0", border=False)
    style_range(ag, f"C{FIRST}:R{LAST_AGG}", font=blue, border=False)
    style_range(ag, f"E{FIRST}:E{LAST_AGG}", number_format="@", border=False)
    style_range(ag, f"G{FIRST}:G{LAST_AGG}", number_format="0", border=False)
    style_range(ag, f"K{FIRST}:K{LAST_AGG}", number_format=DATE, border=False)
    style_range(ag, f"L{FIRST}:P{LAST_AGG}", number_format=MONEY, border=False)
    style_range(ag, f"R{FIRST}:R{LAST_AGG}", align=Alignment(horizontal="center"), border=False)
    for r in range(FIRST, LAST_AGG + 1):
        ag[f"U{r}"] = (f'=IF(OR($C{r}="",$J{r}="",UPPER($R{r})="X"),"",IFERROR(IF(VLOOKUP($J{r},Rates!$A${FIRST}:$C${LAST_RATE},3,FALSE)>0,'
                       f'VLOOKUP($J{r},Rates!$A${FIRST}:$C${LAST_RATE},3,FALSE),""),""))')
        for col, src in zip("VWXYZ", "LMNOP"):
            ag[f"{col}{r}"] = f'=IF($U{r}="",0,N({src}{r})*$U{r})'
    style_range(ag, f"U{FIRST}:U{LAST_AGG}", font=f(9), number_format=RATE, bg=CALC_BG, border=False)
    style_range(ag, f"V{FIRST}:Z{LAST_AGG}", font=f(10), number_format=MONEY, bg=CALC_BG, border=False)
    ag.freeze_panes = "E6"
    ag.auto_filter.ref = f"A5:Z{LAST_AGG}"
    data_ref = f"A{FIRST}:A{LAST_AGG}"
    ag.conditional_formatting.add(data_ref, FormulaRule(formula=[f'LEFT($A{FIRST},5)="ERROR"'],
                                  fill=fill("F8D7DA"), font=Font(name=FONT, color="9C0006", bold=True)))
    ag.conditional_formatting.add(data_ref, FormulaRule(formula=[f'OR(LEFT($A{FIRST},3)="NOT",LEFT($A{FIRST},4)="SKIP")'],
                                  fill=fill("FFF3CD"), font=Font(name=FONT, color="7A5B00")))
    ag.conditional_formatting.add(data_ref, FormulaRule(formula=[f'$A{FIRST}="OK"'], font=Font(name=FONT, color="2F8F6B")))
    ag.conditional_formatting.add(f"C{FIRST}:Z{LAST_AGG}", FormulaRule(formula=[f'UPPER($R{FIRST})="X"'],
                                  font=Font(name=FONT, strike=True, color="999999")))
    ag.conditional_formatting.add(f"U{FIRST}:U{LAST_AGG}", FormulaRule(formula=[f'AND($C{FIRST}<>"",$J{FIRST}<>"",$U{FIRST}="",UPPER($R{FIRST})<>"X")'],
                                  fill=fill("FFF3CD")))

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
    dv(ag, f"L{FIRST}:P{LAST_AGG}", type="decimal", operator="greaterThanOrEqual", formula1="0",
       errorTitle="Aggregate", error="Aggregates are amounts of zero or more.")
    dv(ag, f"R{FIRST}:R{LAST_AGG}", type="list", formula1='"X"')

    if demo:
        for i, row in enumerate(DEMO_ROWS):
            r = FIRST + i
            ced, ctry, zone, zname, yr, tt, cob, ccy, *amts = row
            vals = {"C": ced, "D": ctry, "E": zone or None, "F": zname or None, "G": yr, "H": tt, "I": cob,
                    "J": ccy, "K": date(2026, 9, 30)}
            for col, v in vals.items():
                ag[f"{col}{r}"] = v
            for col, v in zip("LMNOP", amts):
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
    widths(li, {"A": 24, "B": 20, "C": 3, "D": 24, "E": 9, "F": 3, "G": 9, "H": 32, "I": 3, "J": 14, "K": 3})
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
    li.freeze_panes = "A6"

    # ================= By Cedant / By Country ===================================
    def summary_sheet(ws, kind):
        is_ced = kind == "cedant"
        src, last = ("Cedants", LAST_CED) if is_ced else ("Countries", LAST_CTRY)
        key = "agg_Cedant" if is_ced else "agg_Country"
        n = last - FIRST + 1
        top, bot = FIRST + 1, FIRST + n          # data rows (row 6 = total)
        title_bar(ws, 12, f"Aggregates by {kind} (USD)",
                  f"USD at the latest rate on the Rates sheet. UW year filter is on the Dashboard. "
                  f"Peak = largest single-peril aggregate (perils are not added together).")
        ws.cell(4, 1, '="UW year: "&sel_UWYear').font = f(9, True, TEAL)
        b_lab = "Home country" if is_ced else "ISO2"
        c_lab = "Active" if is_ced else "Region"
        header(ws, 5, [kind.capitalize(), b_lab, c_lab, "Rows", "EQ", "WS", "Flood", "SRCC", "Others",
                       "Peak peril agg", "Peak peril", "Rank"])
        widths(ws, {"A": 34, "B": 20 if is_ced else 7, "C": 9 if is_ced else 13, "D": 7, "E": 16, "F": 16,
                    "G": 16, "H": 16, "I": 16, "J": 17, "K": 10, "L": 7})
        # total row
        ws.cell(FIRST, 1, "TOTAL")
        ws.cell(FIRST, 4, f"=SUM(D{top}:D{bot})")
        for col in "EFGHI":
            ws[f"{col}{FIRST}"] = f"=SUM({col}{top}:{col}{bot})"
        ws[f"J{FIRST}"] = f"=MAX(E{FIRST}:I{FIRST})"
        ws[f"K{FIRST}"] = f'=IF(N(J{FIRST})=0,"",INDEX(lst_Perils,MATCH(J{FIRST},E{FIRST}:I{FIRST},0)))'
        style_range(ws, f"A{FIRST}:L{FIRST}", font=f(10, True, TOTAL_INK), bg=TOTAL_BG)
        for r in range(top, bot + 1):
            s = r - 1   # source row on master sheet
            ws[f"A{r}"] = f'=IF({src}!A{s}="","",{src}!A{s})'
            if is_ced:
                ws[f"B{r}"] = f'=IF(OR($A{r}="",Cedants!B{s}=""),"",Cedants!B{s})'
                ws[f"C{r}"] = f'=IF($A{r}="","",IF(Cedants!C{s}="","Y",Cedants!C{s}))'
            else:
                ws[f"B{r}"] = f'=IF(OR($A{r}="",Countries!B{s}=""),"",Countries!B{s})'
                ws[f"C{r}"] = f'=IF(OR($A{r}="",Countries!D{s}=""),"",Countries!D{s})'
            ws[f"D{r}"] = f'=IF($A{r}="","",COUNTIFS({key},$A{r},agg_UWYear,crit_UWYear))'
            for col, nm in zip("EFGHI", ["agg_EQ_USD", "agg_WS_USD", "agg_Flood_USD", "agg_SRCC_USD", "agg_Others_USD"]):
                ws[f"{col}{r}"] = f'=IF($A{r}="","",SUMIFS({nm},{key},$A{r},agg_UWYear,crit_UWYear))'
            ws[f"J{r}"] = f'=IF($A{r}="","",MAX(E{r}:I{r}))'
            ws[f"K{r}"] = f'=IF(N(J{r})=0,"",INDEX(lst_Perils,MATCH(J{r},E{r}:I{r},0)))'
            ws[f"L{r}"] = f'=IF(N(J{r})=0,"",RANK(J{r},$J${top}:$J${bot}))'
        style_range(ws, f"A{top}:C{bot}", font=f(10), border=False)
        style_range(ws, f"D{top}:D{bot}", font=f(10), number_format="0;-0;\"-\"", border=False)
        style_range(ws, f"E{top}:J{bot}", font=f(10), number_format=MONEY, border=False)
        style_range(ws, f"K{top}:L{bot}", font=f(10), align=Alignment(horizontal="center"), border=False)
        ws[f"D{FIRST}"].number_format = "0"
        for col in "EFGHIJ":
            ws[f"{col}{FIRST}"].number_format = MONEY
        ws.conditional_formatting.add(f"J{top}:J{bot}", DataBarRule(start_type="num", start_value=0, end_type="max",
                                                                    color="7FB8A4", showValue=True))
        ws.conditional_formatting.add(f"A{top}:L{bot}", FormulaRule(formula=[f"MOD(ROW(),2)=0"], fill=fill("F4F9FA")))
        ws.freeze_panes = "B7"
        return top, bot

    ced_top, ced_bot = summary_sheet(byc, "cedant")
    ctry_top, ctry_bot = summary_sheet(byk, "country")

    # ================= Matrix ===================================================
    MX_HDR, MX_TOP, MX_BOT = 7, 8, 207
    title_bar(mx, 23, "Cedant x country matrix (USD)",
              "Pick a peril and change any country in row 7 (dropdowns). 'Other countries' = everything not shown.")
    mx["B5"] = "Peril"; mx["B5"].font = f(10, True)
    mx["C5"] = "EQ"; mx["C5"].font = f(11, True, INPUT_BLUE); mx["C5"].fill = fill("FFFF00"); mx["C5"].border = BOX
    dv(mx, "C5", type="list", formula1="lst_Perils")
    mx["E5"] = '="UW year: "&sel_UWYear&"   (change on Dashboard)"'; mx["E5"].font = f(9, True, TEAL)
    mx.cell(MX_HDR, 1, "Cedant")
    labels = ["Cedant"] + MATRIX_COUNTRIES + ["Other countries", "Total"]
    header(mx, MX_HDR, labels)
    for i in range(20):
        mx.cell(MX_HDR, 2 + i).fill = fill(TEAL)
    dv(mx, f"B{MX_HDR}:U{MX_HDR}", type="list", formula1="lst_Countries")
    mx.column_dimensions["A"].width = 32
    for c in range(2, 24):
        mx.column_dimensions[L(c)].width = 13
    mx.row_dimensions[MX_HDR].height = 42
    for r in range(MX_TOP, MX_BOT + 1):
        s = r - MX_TOP + FIRST
        mx[f"A{r}"] = f'=IF(Cedants!A{s}="","",Cedants!A{s})'
        for c in range(2, 22):
            cl = L(c)
            mx[f"{cl}{r}"] = (f'=IF(OR($A{r}="",{cl}${MX_HDR}=""),"",SUMIFS(INDEX(agg_USD_block,0,sel_PerilIdx),'
                              f'agg_Cedant,$A{r},agg_Country,{cl}${MX_HDR},agg_UWYear,crit_UWYear))')
        mx[f"W{r}"] = f'=IF($A{r}="","",SUMIFS(INDEX(agg_USD_block,0,sel_PerilIdx),agg_Cedant,$A{r},agg_UWYear,crit_UWYear))'
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
              "By Cedant, By Country and Matrix follow it.")
    widths(dash, {"A": 3, "B": 34, "C": 20, "D": 32, "E": 20, "F": 32, "G": 20, "H": 3})
    dash["B5"] = "UW year filter"; dash["B5"].font = f(10, True)
    dash["C5"] = "All"
    dash["C5"].font = f(11, True, INPUT_BLUE); dash["C5"].fill = fill("FFFF00"); dash["C5"].border = BOX
    dash["C5"].alignment = Alignment(horizontal="center")
    dash["D5"] = "Pick All or a single underwriting year"; dash["D5"].font = f(9, italic=True, color=INK_SOFT)
    dv(dash, "C5", type="list", formula1="lst_UWYears")

    kpis = [("Aggregate rows", "=COUNTIFS(agg_Cedant,\"<>\",agg_UWYear,crit_UWYear)"),
            ("Cedants with aggregates", f"=COUNTIF('By Cedant'!$D${ced_top}:$D${ced_bot},\">0\")"),
            ("Countries with aggregates", f"=COUNTIF('By Country'!$D${ctry_top}:$D${ctry_bot},\">0\")")]
    header(dash, 7, ["Portfolio", "Value"], 2)
    for i, (lab, fx) in enumerate(kpis):
        r = 8 + i
        dash.cell(r, 2, lab); dash.cell(r, 3, fx)
        style_range(dash, f"B{r}:C{r}", font=f(10))
        dash.cell(r, 3).number_format = "#,##0"
        dash.cell(r, 3).font = f(11, True)

    header(dash, 7, ["Data checks (all years)", "Rows"], 4)
    checks = [
        ("Rows with no FX rate (excluded from USD)",
         '=SUMPRODUCT((agg_Cedant<>"")*(UPPER(agg_Delete)<>"X")*(agg_Ccy<>"")*(agg_FX=""))'),
        ("Rows with a cedant/country not on the lists",
         f'=SUMPRODUCT((agg_Cedant<>"")*(COUNTIF(Cedants!$A${FIRST}:$A${LAST_CED},agg_Cedant)=0))'
         f'+SUMPRODUCT((agg_Country<>"")*(COUNTIF(Countries!$A${FIRST}:$A${LAST_CTRY},agg_Country)=0))'),
        ("Rows not OK in Access (new or error)", '=SUMPRODUCT((agg_Cedant<>"")*(agg_Status<>"OK"))'),
    ]
    for i, (lab, fx) in enumerate(checks):
        r = 8 + i
        dash.cell(r, 4, lab); dash.cell(r, 5, fx)
        style_range(dash, f"D{r}:E{r}", font=f(10))
        dash.cell(r, 5).number_format = "#,##0"
        dash.cell(r, 5).font = f(11, True)
    dash.conditional_formatting.add("E8:E10", CellIsRule(operator="greaterThan", formula=["0"],
                                    fill=fill("FFF3CD"), font=Font(name=FONT, bold=True, color="7A5B00")))

    header(dash, 7, ["Sync", ""], 6)
    dash["F8"] = "Last pull from Access"; dash["G8"] = "=IF(cfg_LastPull=\"\",\"never\",cfg_LastPull)"
    dash["F9"] = "Last push to Access"; dash["G9"] = "=IF(cfg_LastPush=\"\",\"never\",cfg_LastPush)"
    style_range(dash, "F8:G9", font=f(10), number_format=STAMP)
    style_range(dash, "F10:G10", font=f(10))

    header(dash, 13, ["Peril", "Aggregate (USD)", "Largest cedant", "Cedant agg (USD)", "Largest country",
                      "Country agg (USD)"], 2)
    for i, (code, meaning) in enumerate(PERILS):
        r = 14 + i
        col = "EFGHI"[i]
        cr = f"'By Cedant'!${col}${ced_top}:${col}${ced_bot}"
        kr = f"'By Country'!${col}${ctry_top}:${col}${ctry_bot}"
        dash.cell(r, 2, f"{code} - {meaning}")
        dash.cell(r, 3, f"='By Cedant'!{col}{FIRST}")
        dash.cell(r, 4, f"=IF(MAX({cr})=0,\"-\",INDEX('By Cedant'!$A${ced_top}:$A${ced_bot},MATCH(MAX({cr}),{cr},0)))")
        dash.cell(r, 5, f"=MAX({cr})")
        dash.cell(r, 6, f"=IF(MAX({kr})=0,\"-\",INDEX('By Country'!$A${ctry_top}:$A${ctry_bot},MATCH(MAX({kr}),{kr},0)))")
        dash.cell(r, 7, f"=MAX({kr})")
        style_range(dash, f"B{r}:G{r}", font=f(10))
        for c in (3, 5, 7):
            dash.cell(r, c).number_format = MONEY
        dash.cell(r, 3).font = f(10, True)
    r = 19
    dash.cell(r, 2, "Peak single-peril aggregate")
    dash.cell(r, 3, f"='By Cedant'!J{FIRST}")
    dash.cell(r, 4, f"=\"Peril: \"&'By Cedant'!K{FIRST}")
    style_range(dash, f"B{r}:G{r}", font=f(10, True, TOTAL_INK), bg=TOTAL_BG)
    dash.cell(r, 3).number_format = MONEY
    dash["B21"] = ("Perils are tracked separately and never added together: the same sum insured is usually exposed "
                   "to several perils, so the meaningful single figure is the peak peril.")
    dash["B21"].font = f(9, italic=True, color=INK_SOFT)
    dash.merge_cells("B21:G22")
    dash["B21"].alignment = Alignment(wrap_text=True, vertical="top")

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
            "Read the results on Dashboard, By Cedant, By Country and Matrix."]
    for i, t in enumerate(flow):
        st.cell(30 + i, 2, f"{i + 1}."); st.cell(30 + i, 3, t)
        st.merge_cells(start_row=30 + i, start_column=3, end_row=30 + i, end_column=4)
        st.cell(30 + i, 2).font = f(10, True); st.cell(30 + i, 3).font = f(10)
        st.cell(30 + i, 2).alignment = Alignment(horizontal="right")

    st["B36"] = "Example aggregate row"; st["B36"].font = f(12, True, TEAL)
    ex_hdr = ["Cedant", "Country", "Zone ID", "UW year", "Treaty type", "Class", "Ccy", "EQ", "WS", "Flood", "SRCC", "Others"]
    ex_val = ["Tawuniya", "Saudi Arabia", "SA-01", 2026, "Quota Share", "Property", "SAR",
              450000000, 0, 120000000, 80000000, 0]
    st["B37"] = "Column"; st["C37"] = "Example value"; st["D37"] = "Rule"
    header(st, 37, ["Column", "Example value", "Rule"], 2)
    rules = ["Must exist on Cedants", "Must exist on Countries", "Optional - blank = ALL (country total)",
             "Whole year 1990-2100", "Optional - blank = ALL", "Optional - blank = ALL",
             "Must exist on Rates; amounts are in this currency", "Aggregate exposed to earthquake, >= 0",
             "Aggregate exposed to windstorm", "Aggregate exposed to flood", "Aggregate exposed to SRCC",
             "Any other peril"]
    for i, (h, v, rule) in enumerate(zip(ex_hdr, ex_val, rules)):
        r = 38 + i
        st.cell(r, 2, h).font = f(10, True)
        c = st.cell(r, 3, v); c.font = f(10, color=INPUT_BLUE)
        if isinstance(v, (int, float)) and i >= 7:
            c.number_format = MONEY
        c.alignment = Alignment(horizontal="left")
        st.cell(r, 4, rule).font = f(10)
        style_range(st, f"B{r}:D{r}")
    st["B50"] = ("One row per cedant + country + zone + UW year + treaty type + class + currency. "
                 "That combination must be unique (enforced in Access).")
    st["B50"].font = f(9, italic=True, color=INK_SOFT)

    st["B52"] = "Colour legend"; st["B52"].font = f(12, True, TEAL)
    st["B53"] = "Blue text"; st["B53"].font = f(10, color=INPUT_BLUE); st["C53"] = "Input you type"
    st["B54"] = "Yellow cell"; st["B54"].fill = fill("FFFF00"); st["C54"] = "Setting / filter"
    st["B55"] = "Grey-blue cell"; st["B55"].fill = fill(CALC_BG); st["C55"] = "Formula - do not type"
    st["B56"] = "Status = ERROR"; st["B56"].fill = fill("F8D7DA"); st["C56"] = "Row refused on push - read the message"
    for r in range(53, 57):
        st.cell(r, 3).font = f(10)

    st["B58"] = "How this mirrors Universe"; st["B58"].font = f(12, True, TEAL)
    header(st, 59, ["Universe", "Here", "What it does"], 2)
    arch = [("client (React screens)", "Excel sheets + AggMacros.bas", "Entry grids, dashboards, buttons"),
            ("server/src/services, validation", "AggService.bas", "Validation, push/pull, transactions"),
            ("modules/*/repositories", "AggRepo.bas", "Every SQL statement"),
            ("db/pool.js, withTransaction.js", "AggDb.bas", "ADO connection, parameterised commands"),
            ("db/migrations + runMigrations.js", "AggMigrate.bas + schema_migrations", "Numbered forward-only schema changes"),
            ("startup/ensureReferenceData.js", "AggSeed.bas + Lists / Cedants sheets", "Predefined treaty types, classes, cedants"),
            ("contract_cresta_data", "cedant_aggregate", "Peril aggregates by country / CRESTA zone"),
            ("ref_exchange_rate (rate_to_usd)", "ref_exchange_rate + vw_fx_latest", "Latest rate per currency"),
            ("db/optimisticLock.js", "row_version column", "Refuses saves over someone else's change"),
            ("audit_log", "audit_log", "Who changed what, when")]
    for i, row in enumerate(arch):
        r = 60 + i
        for j, v in enumerate(row):
            st.cell(r, 2 + j, v).font = f(10)
        style_range(st, f"B{r}:D{r}")

    wb.active = 0
    wb.save(path)


if __name__ == "__main__":
    out = sys.argv[1]
    build(f"{out}/Aggregate_Tracker.xlsx", demo=False)
    build(f"{out}/Aggregate_Tracker_DEMO.xlsx", demo=True)
    print("built")

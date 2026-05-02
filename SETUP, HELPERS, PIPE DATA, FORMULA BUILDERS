# STAGE 1 — SETUP, HELPERS, PIPE DATA, FORMULA BUILDERS
# Run this cell FIRST. All subsequent stages depend on it.
# Google Colab: upload 'HDPE_Pipes_-Gentex_1_.xls' before running.

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from datetime import date

# FILE PATHS
HDPE_FILE   = '/content/HDPE_Pipes_-Gentex(1).xls'  # change to '/content/...' in Colab
OUTPUT_FILE = '/content/Uganda_Pipe_Sizing_Tool.xlsx' # change to '/content/...' in Colab

# THEME COLORS
C_DARK  = "1F3864"; C_MED   = "2E75B6"; C_LIGHT = "DEEAF1"
C_INP   = "FFF2CC"; C_INPT  = "003366"; C_FORM  = "000000"; C_LINK  = "375623"
C_SUBBG = "2E75B6"; C_SUBFG = "FFFFFF"; C_SECBG = "DEEAF1"; C_TOTBG = "D6DCE4"
C_OKBG  = "E2EFDA"; C_OKFG  = "375623"
C_WRNBG = "FFEB9C"; C_WRNFG = "9C5700"
C_ERRBG = "FFC7CE"; C_ERRFG = "9C0006"

# STYLE HELPERS
def BD(s='thin'):
    ss = Side(style=s)
    return Border(left=ss, right=ss, top=ss, bottom=ss)

def FL(c): return PatternFill('solid', fgColor=c)

def FN(bold=False, sz=10, c='000000', it=False):
    return Font(name='Calibri', size=sz, bold=bold, color=c, italic=it)

def AL(h='center', v='center', w=False, i=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=w, indent=i)

def hdr(ws, row, txt, c1, c2, sz=12, bg=C_DARK, fg='FFFFFF', ht=28):
    """Full-width title header row."""
    c = ws.cell(row=row, column=c1, value=txt)
    c.font = Font(name='Calibri', size=sz, bold=True, color=fg)
    c.fill = FL(bg); c.alignment = AL(w=True)
    ws.row_dimensions[row].height = ht
    if c2 > c1:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)

def subhdr(ws, row, txt, c1, c2, ht=18):
    """Section subheader (light-blue background)."""
    c = ws.cell(row=row, column=c1, value=txt)
    c.font = FN(bold=True, sz=10, c=C_DARK)
    c.fill = FL(C_SECBG); c.alignment = AL(h='left', i=1)
    ws.row_dimensions[row].height = ht
    if c2 > c1:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)

def col_hdr_row(ws, row, headers, c1=1, bg=C_SUBBG, fg=C_SUBFG, ht=38):
    """Column header row for data tables."""
    ws.row_dimensions[row].height = ht
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=c1 + i, value=h)
        c.font = Font(name='Calibri', size=9, bold=True, color=fg)
        c.fill = FL(bg); c.alignment = AL(w=True); c.border = BD()

def inp(ws, row, col, val=None, fmt=None):
    """User input cell — yellow background, dark blue text."""
    c = ws.cell(row=row, column=col, value=val)
    c.font = FN(c=C_INPT); c.fill = FL(C_INP)
    c.alignment = AL(); c.border = BD()
    if fmt: c.number_format = fmt
    return c

def frm(ws, row, col, formula, fmt=None, bg=None, lnk=False, h='center'):
    """Formula / calculated cell."""
    c = ws.cell(row=row, column=col, value=formula)
    c.font = FN(c=C_LINK if lnk else C_FORM)
    if bg: c.fill = FL(bg)
    c.alignment = AL(h=h); c.border = BD()
    if fmt: c.number_format = fmt
    return c

def lbl(ws, row, col, txt, bold=False, sz=10, ind=0, bg=None, wrap=False):
    """Plain label cell (no border, left-aligned)."""
    c = ws.cell(row=row, column=col, value=txt)
    c.font = FN(bold=bold, sz=sz); c.alignment = AL(h='left', i=ind, w=wrap)
    if bg: c.fill = FL(bg)
    return c

def tot(ws, row, col, formula, fmt=None):
    """Totals row cell."""
    c = ws.cell(row=row, column=col, value=formula)
    c.font = FN(bold=True); c.fill = FL(C_TOTBG)
    c.alignment = AL(); c.border = BD()
    if fmt: c.number_format = fmt

def cw(ws, d):
    """Set column widths: {'A': 20, 'B': 15, ...}"""
    for col, w in d.items():
        ws.column_dimensions[col].width = w

# PIPE DATA
df_raw = pd.read_excel(HDPE_FILE, engine='xlrd', header=None, skiprows=3)
df_raw.columns = ['Code', 'OD', 'ID', 'Wall', 'Weight', 'PN_label', '_', 'Price']
df_p = df_raw[df_raw['Code'].astype(str).str.match(r'^\d')].copy()
df_p['PN_num'] = df_p['PN_label'].str.extract(r'(\d+)').astype(int)
df_p = df_p.sort_values(['OD', 'PN_num']).reset_index(drop=True)
NP   = len(df_p)        # 74 pipe entries
DB_R1 = 4               # first data row in PIPE_DB sheet
DB_RN = DB_R1 + NP - 1  # = 77 (last data row)

# PIPE_DB column layout: A=Code, B=OD, C=PN_label, D=PN_num, E=ID, F=Wall, G=Weight, H=Price
def _dr(col_n):
    """Absolute Excel range for a PIPE_DB column."""
    L = get_column_letter(col_n)
    return f"PIPE_DB!${L}${DB_R1}:${L}${DB_RN}"

DB_OD = _dr(2)   # OD (mm)
DB_PN = _dr(4)   # PN numeric (6, 10, 16, 20)
DB_ID = _dr(5)   # Internal diameter (mm)
DB_WG = _dr(7)   # Weight (kg/m)

# PARAMETERS SHEET — Row map (column C holds values)
PR = {
    'C_HW':  5,   # Hazen-Williams C coefficient
    'VMINT': 6,   # Min velocity — transmission (m/s)
    'VMAXT': 7,   # Max velocity — transmission (m/s)
    'VMIND': 8,   # Min velocity — distribution (m/s)
    'VMAXD': 9,   # Max velocity — distribution (m/s)
    'MINOR': 10,  # Minor loss factor (k)
    'PN6':   13,  # PN-6  max working head (m)
    'PN10':  14,  # PN-10 max working head (m)
    'PN16':  15,  # PN-16 max working head (m)
    'PN20':  16,  # PN-20 max working head (m)
    'SF':    17,  # Pressure safety factor
    'RDIST': 20,  # Min residual pressure — distribution consumer (m)
    'RRES':  21,  # Min residual pressure — reservoir/transmission end (m)
    'STRANS':22,  # Allowable head loss gradient — transmission (m/km)
    'SDIST': 23,  # Allowable head loss gradient — distribution (m/km)
    'PUMPH': 24,  # Pump operating hours/day
    'PERCAP':27,  # Per capita demand (L/cap/day)
    'PHF':   28,  # Peak hour factor
    'PDF':   29,  # Peak day factor
    'NRW':   30,  # NRW/loss factor
    'GR':    31,  # Population growth rate (/yr)
    'PERIOD':32,  # Design period (years)
}
def p(k): return f"PARAMETERS!$C${PR[k]}"   # absolute cell reference helper

# KEY ROW NUMBERS (used for cross-sheet formulas)
TDH_R = {
    'GE':5, 'SWL':6, 'PWL':7, 'PSD':8, 'Q':9, 'PH':10,
    'RF':11, 'RH':12, 'RM':13,
    'DWL':16, 'SWL_E':17, 'SH':18,
    'SEG_S':23, 'SEG_E':34, 'TOT':35,
    'SUM_HDR':37, 'SH_COMP':38, 'FL_COMP':39, 'ML_COMP':40, 'VH_COMP':41,
    'TDH':43,
    'PUMP_HDR':45, 'P_Q':46, 'P_TDH':47, 'P_PEFF':48,
    'P_MEFF':49, 'P_HPOW':50, 'P_MPOW':51, 'P_VOL':52,
}
TR_R = {
    'SRC':5, 'DST':6, 'Q':7, 'HGL':8, 'S':9, 'RES':10,
    'SEG_S':16, 'SEG_E':30, 'TOT':31, 'SUM_HDR':33,
}
DI_R = {
    'POP':6, 'DPOP':7, 'PERCAP':8, 'NRW':9, 'PHF':10,
    'ADD':11, 'ADF':12, 'DES_Q':13, 'HGL':14, 'RRES':15, 'S':16,
    'SEG_S':22, 'SEG_E':41, 'TOT':42,
}

# FORMULA BUILDERS
def F_PN_num(max_press_ref):
    """Auto-select PN number from max working pressure."""
    return (f'=IF({max_press_ref}="""","""",IF({max_press_ref}<={p("PN6")}/{p("SF")},6,'
            f'IF({max_press_ref}<={p("PN10")}/{p("SF")},10,'
            f'IF({max_press_ref}<={p("PN16")}/{p("SF")},16,20))))')

def F_PN_lbl(pn_num_ref):
    """PN rating label text from PN number."""
    return (f'=IF({pn_num_ref}="""","""",IF({pn_num_ref}=6,"PN - 6",'
            f'IF({pn_num_ref}=10,"PN - 10",'
            f'IF({pn_num_ref}=16,"PN - 16","PN - 20"))))')

def F_Dreq(q_ref, s_km_ref, vmax_key):
    """
    Minimum required internal diameter (mm) from:
      - Hazen-Williams: D = [10.67·Q^1.852 / (C^1.852·S)]^(1/4.87)
      - Velocity constraint: D = sqrt(4Q / π·Vmax)
    Returns MAX of the two.
    """
    c  = p('C_HW'); v = p(vmax_key)
    hw = (f"1000*((10.67*({q_ref}/1000)^1.852)"
          f"/({c}^1.852*({s_km_ref}/1000)))^(1/4.87)")
    vl = f"1000*SQRT(4*({q_ref}/1000)/(PI()*{v}))"
    return f'=IF({q_ref}<=0,"""",IFERROR(MAX({hw},{vl}),0))'

def F_OD(d_ref, pn_ref, q_ref):
    """
    Auto-select smallest OD from PIPE_DB where:
      internal diameter ≥ D_req  AND  PN_num = selected PN
    Uses AGGREGATE(15,6,...) — smallest non-error value (Excel 2010+).
    """
    return (f'=IF(OR({q_ref}<=0,{d_ref}=""""),"""",IFERROR('
            f'AGGREGATE(15,6,{DB_OD}/(({DB_ID}>={d_ref})*({DB_PN}={pn_ref})),1)'
            f',"SIZE N/A"))')

def F_ID(od_ref, pn_ref):
    """Look up actual internal diameter for selected (OD, PN) pair."""
    return (f'=IF(OR({od_ref}="""",{od_ref}="SIZE N/A",'
            f'NOT(ISNUMBER({od_ref}))),"""",IFERROR('
            f'SUMPRODUCT(({DB_OD}={od_ref})*({DB_PN}={pn_ref})*{DB_ID}),0))')

def F_WT(od_ref, pn_ref):
    """Look up pipe weight (kg/m) for selected (OD, PN) pair."""
    return (f'=IF(OR({od_ref}="""",{od_ref}="SIZE N/A",'
            f'NOT(ISNUMBER({od_ref}))),"""",IFERROR('
            f'SUMPRODUCT(({DB_OD}={od_ref})*({DB_PN}={pn_ref})*{DB_WG}),0))')

def F_VEL(q_ref, id_ref):
    """Flow velocity (m/s): V = 4Q / (π·D²) with Q in L/s, D in mm."""
    return (f'=IF(OR({q_ref}<=0,{id_ref}="""",NOT(ISNUMBER({id_ref}))),"""",IFERROR('
            f'4*({q_ref}/1000)/(PI()*({id_ref}/1000)^2),0))')

def F_HF(q_ref, id_ref, L_ref):
    """
    Hazen-Williams friction head loss (m):
      hf = 10.67 · L · Q^1.852 / (C^1.852 · D^4.87)
    Q in L/s, D in mm, L in m → hf in m.
    """
    c = p('C_HW')
    return (f'=IF(OR({q_ref}<=0,{id_ref}="""",NOT(ISNUMBER({id_ref}))),"""",IFERROR('
            f'10.67*{L_ref}*({q_ref}/1000)^1.852/({c}^1.852*({id_ref}/1000)^4.87),0))')

def F_ML(hf_ref):
    """Minor losses = k × friction loss (k from PARAMETERS)."""
    return f'=IF({hf_ref}="""","""",IFERROR({p("MINOR")}*{hf_ref},0))'

def F_TL(hf_ref, ml_ref):
    """Total head loss = friction + minor."""
    return f'=IF({hf_ref}="""","""",IFERROR({hf_ref}+{ml_ref},0))'

def F_HGL_E(hgl_s_ref, tl_ref):
    """HGL at segment end = HGL start − total loss."""
    return f'=IF(OR({hgl_s_ref}="""",{tl_ref}=""""),"""",IFERROR({hgl_s_ref}-{tl_ref},0))'

def F_PRESS(hgl_e_ref, elev_ref):
    """Pressure head at node = HGL − ground elevation."""
    return f'=IF({hgl_e_ref}="""","""",IFERROR({hgl_e_ref}-{elev_ref},0))'

def F_STATUS(vel_ref, press_ref, od_ref, vmax_key, vmin_key, rp_key):
    """
    Status check:
      ✓ OK           — velocity in range AND pressure ≥ minimum
      ⚠ LOW VELOCITY — velocity below minimum
      ⚠ HIGH VELOCITY— velocity above maximum
      ⚠ LOW PRESSURE — pressure below minimum
      ⚠ SIZE N/A     — no pipe in catalog meets criteria
    """
    vm = p(vmax_key); vi = p(vmin_key); rp = p(rp_key)
    return (f'=IF({od_ref}="""","""",IF({od_ref}="SIZE N/A","⚠ SIZE N/A",'
            f'IF({vel_ref}="""","""",IF(AND({vel_ref}>={vi},{vel_ref}<={vm},{press_ref}>={rp}),'
            f'"✓ OK",IF({vel_ref}<{vi},"⚠ LOW VELOCITY",'
            f'IF({vel_ref}>{vm},"⚠ HIGH VELOCITY","⚠ LOW PRESSURE"))))))')

# CREATE WORKBOOK
wb = openpyxl.Workbook()
wb.remove(wb.active)

print("✅ Stage 1 complete")
print(f"   Pipe catalog loaded: {NP} entries | OD 20–450 mm | PN-6 to PN-20")
print(f"   Formula builders ready | Workbook initialised")

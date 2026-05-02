# STAGE 2 — COVER · PARAMETERS · PIPE_DB SHEETS
# Run AFTER Stage 1.

# COVER
ws_cv = wb.create_sheet("COVER")
ws_cv.sheet_view.showGridLines = False

hdr(ws_cv, 1, "WATER SUPPLY PIPE SIZING TOOL", 1, 14, sz=22, ht=42)
ws_cv.merge_cells('A1:N1')
hdr(ws_cv, 2,
    "Hazen-Williams  ·  HDPE PE100 DIN 8074 (Gentex)  ·  Uganda Water Supply Design Manual – 2nd Edition",
    1, 14, sz=10, bg=C_MED, ht=22)
ws_cv.merge_cells('A2:N2')
ws_cv.row_dimensions[3].height = 10

# Project information block
subhdr(ws_cv, 4, "PROJECT INFORMATION", 1, 9)
ws_cv.merge_cells('A4:I4')
for i, lab in enumerate(["Project Title","Project Location","Designed By",
                          "Checked By","Client / Owner","Contract No.","Date","Revision"]):
    r = 5 + i
    lbl(ws_cv, r, 1, lab, bold=True, ind=1)
    default = date.today().strftime("%d-%b-%Y") if lab == "Date" else ("Rev 0" if lab == "Revision" else None)
    inp(ws_cv, r, 2, default)
    ws_cv.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
    ws_cv.row_dimensions[r].height = 18
ws_cv.row_dimensions[13].height = 10

# Auto-populated design summary
subhdr(ws_cv, 14, "DESIGN SUMMARY  (auto-populated from calculation sheets)", 1, 14)
ws_cv.merge_cells('A14:N14')
col_hdr_row(ws_cv, 15, ["Description","Value","Unit","Source Sheet"], 1, ht=22)
summ = [
    ("Total Dynamic Head (TDH)",
     f"=IFERROR(TDH_CALC!$E${TDH_R['TDH']},\"-\")",              "m",      "TDH_CALC"),
    ("Design Flow Rate",
     f"=IFERROR(TDH_CALC!$E${TDH_R['Q']},\"-\")",                "L/s",    "TDH_CALC"),
    ("Estimated Motor Power Required",
     f"=IFERROR(TDH_CALC!$E${TDH_R['P_MPOW']},\"-\")",           "kW",     "TDH_CALC"),
    ("Transmission Main — Total Length",
     f"=IFERROR(SUM(TRANSMISSION!$D${TR_R['SEG_S']}:$D${TR_R['SEG_E']}),\"-\")", "m", "TRANSMISSION"),
    ("Distribution Network — Total Length",
     f"=IFERROR(SUM(DISTRIBUTION!$D${DI_R['SEG_S']}:$D${DI_R['SEG_E']}),\"-\")", "m", "DISTRIBUTION"),
]
for i, (desc, fm, unit, src) in enumerate(summ):
    r = 16 + i
    lbl(ws_cv, r, 1, desc, ind=1);    ws_cv.cell(row=r, column=1).border = BD()
    frm(ws_cv, r, 2, fm, '0.00', lnk=True)
    lbl(ws_cv, r, 3, unit);            ws_cv.cell(row=r, column=3).border = BD()
    lbl(ws_cv, r, 4, src);             ws_cv.cell(row=r, column=4).border = BD()
    ws_cv.row_dimensions[r].height = 18
ws_cv.row_dimensions[21].height = 10

# Colour-coding legend
subhdr(ws_cv, 22, "CELL COLOUR CODING", 1, 14)
ws_cv.merge_cells('A22:N22')
for i, (desc, bg, fg, note) in enumerate([
    ("Input Cell  (yellow bg, dark-blue text)",  C_INP,  C_INPT, "Enter data here"),
    ("Formula / Calculated  (black text)",        None,   C_FORM, "Do not edit — auto-calculated"),
    ("Cross-sheet Link  (green text)",            None,   C_LINK, "References another sheet"),
    ("✓ OK — within design limits",              C_OKBG, C_OKFG, "Velocity and pressure acceptable"),
    ("⚠ Warning / Out of Range",                C_ERRBG,C_ERRFG,"Action required — revise design"),
]):
    r = 23 + i
    c = ws_cv.cell(row=r, column=1, value=desc)
    c.font = FN(bold=True, c=fg, sz=10)
    if bg: c.fill = FL(bg)
    c.border = BD()
    n = ws_cv.cell(row=r, column=2, value=note)
    n.font = FN(sz=10); n.alignment = AL(h='left', i=1)
    ws_cv.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
    ws_cv.row_dimensions[r].height = 18
ws_cv.row_dimensions[28].height = 10

# Navigation guide
subhdr(ws_cv, 29, "WORKSHEET GUIDE", 1, 14)
ws_cv.merge_cells('A29:N29')
col_hdr_row(ws_cv, 30, ["Sheet","Purpose","Instructions"], 1, ht=20)
for i, (sn, desc, inst) in enumerate([
    ("PARAMETERS",   "Design constants — H-W C, velocities, pressure ratings, demand factors",
                     "Review and adjust defaults before running calculations"),
    ("PIPE_DB",      "Full HDPE PE100 catalog (Gentex DIN 8074) — 74 entries OD 20–450 mm",
                     "Do NOT modify — auto-referenced by all formulas"),
    ("TDH_CALC",     "Total Dynamic Head: production well → storage reservoir; pump sizing",
                     "Step 1 — enter well & reservoir survey data"),
    ("TRANSMISSION", "Transmission main — auto OD & PN selection, HGL profile, status checks",
                     "Step 2 — enter segment lengths & elevations"),
    ("DISTRIBUTION", "Distribution network — demand calc, branch sizing, node pressure checks",
                     "Step 3 — enter branch data; link HGL Start to junction nodes"),
]):
    r = 31 + i
    c = ws_cv.cell(row=r, column=1, value=sn)
    c.font = Font(name='Calibri', size=10, bold=True, color=C_MED)
    c.hyperlink = f"#{sn}!A1"; c.border = BD()
    lbl(ws_cv, r, 2, desc); ws_cv.cell(row=r, column=2).border = BD()
    lbl(ws_cv, r, 3, inst); ws_cv.cell(row=r, column=3).border = BD()
    ws_cv.row_dimensions[r].height = 18

cw(ws_cv, {'A':28,'B':22,'C':12,'D':20,'E':16,'F':16,
            'G':16,'H':16,'I':16,'J':14,'K':14,'L':14,'M':14,'N':14})

# PARAMETERS 
ws_pm = wb.create_sheet("PARAMETERS")
ws_pm.sheet_view.showGridLines = False
hdr(ws_pm, 1, "DESIGN PARAMETERS", 1, 6, sz=13, ht=28)
ws_pm.merge_cells('A1:F1')
ws_pm.row_dimensions[2].height = 8
col_hdr_row(ws_pm, 3, ["#","Parameter","Value","Unit","Reference / Notes"], 1, ht=20)

params_data = [
    ("HYDRAULIC PARAMETERS", [
        ("Hazen-Williams Coefficient C  (HDPE PE100)",         150,   "-",       "DIN 8074; Uganda WSS DM Sect 4.3",                   True),
        ("Min. Flow Velocity — Transmission Main",              0.5,  "m/s",     "Sedimentation prevention; Uganda WSS DM",             True),
        ("Max. Flow Velocity — Transmission Main",              2.5,  "m/s",     "Erosion / surge limit for HDPE",                      True),
        ("Min. Flow Velocity — Distribution",                   0.3,  "m/s",     "Sedimentation prevention; Uganda WSS DM",             True),
        ("Max. Flow Velocity — Distribution",                   1.5,  "m/s",     "Erosion limit; Uganda WSS DM",                        True),
        ("Minor Loss Factor k  (fraction of friction loss)",    0.05, "-",       "Typical for well-aligned mains; adjust per design",   True),
    ]),
    ("PRESSURE RATINGS — HDPE PE100  DIN 8074  (PE100 compound)", [
        ("PN-6   Max. Working Head",   61.2,  "m",  "6 bar  → 61.2 m H₂O  (design limit = PN head ÷ SF)", False),
        ("PN-10  Max. Working Head",  102.0,  "m",  "10 bar → 102.0 m H₂O",                               False),
        ("PN-16  Max. Working Head",  163.2,  "m",  "16 bar → 163.2 m H₂O",                               False),
        ("PN-20  Max. Working Head",  204.0,  "m",  "20 bar → 204.0 m H₂O",                               False),
        ("Pressure Safety Factor",      1.25,  "-",  "Design working head = PN head ÷ SF",                   True),
    ]),
    ("DESIGN STANDARDS — Uganda Water Supply Design Manual, 2nd Edition", [
        ("Min. Residual Pressure — Distribution Consumer",   7.0, "m",      "Rural WSS; Uganda WSS DM Sect 5",         True),
        ("Min. Residual Pressure — Reservoir / Trans. End",  2.0, "m",      "Above reservoir HWL at delivery point",   True),
        ("Allowable Head Loss Gradient — Transmission",      5.0, "m/km",   "Economic gradient; Uganda WSS DM",        True),
        ("Allowable Head Loss Gradient — Distribution",      8.0, "m/km",   "Economic gradient; Uganda WSS DM",        True),
        ("Pump Design Operating Hours per Day",              8.0, "hr/day", "Uganda WSS DM rural standard",            True),
    ]),
    ("DEMAND PARAMETERS", [
        ("Per Capita Demand",      20.0,  "L/cap/day", "Uganda rural WSS standard; adjust to local survey",     True),
        ("Peak Hour Factor (PHF)",  2.0,  "-",         "Ratio of peak hour to average hourly demand",           True),
        ("Peak Day Factor  (PDF)",  1.5,  "-",         "Ratio of peak day to average daily demand",             True),
        ("NRW / Loss Factor",       1.2,  "-",         "20% non-revenue water allowance",                       True),
        ("Population Growth Rate",  0.03, "/yr",       "3% per year; verify against census data",               True),
        ("Design Period",          20,   "years",      "Standard for Uganda rural water supply projects",        True),
    ]),
]

r = 4; n = 1
for sec_title, items in params_data:
    subhdr(ws_pm, r, sec_title, 1, 6); r += 1
    for txt, val, unit, note, editable in items:
        ws_pm.cell(row=r, column=1, value=n).font = FN(sz=9)
        ws_pm.cell(row=r, column=1).alignment = AL()
        ws_pm.cell(row=r, column=1).border = BD()
        lbl(ws_pm, r, 2, txt, ind=1); ws_pm.cell(row=r, column=2).border = BD()
        if editable:
            inp(ws_pm, r, 3, val, '0.00' if isinstance(val, float) else '0')
        else:
            frm(ws_pm, r, 3, val, '0.0', bg=C_TOTBG)
        lbl(ws_pm, r, 4, unit);    ws_pm.cell(row=r, column=4).border = BD()
        lbl(ws_pm, r, 5, note, sz=9, wrap=True)
        ws_pm.cell(row=r, column=5).border = BD()
        ws_pm.row_dimensions[r].height = 18; r += 1; n += 1
    ws_pm.row_dimensions[r].height = 8; r += 1

cw(ws_pm, {'A':5,'B':52,'C':14,'D':12,'E':62,'F':8})

# PIPE_DB 
ws_db = wb.create_sheet("PIPE_DB")
ws_db.sheet_view.showGridLines = False
ws_db.sheet_properties.tabColor = "595959"

hdr(ws_db, 1, "HDPE PE100 PIPE CATALOG — DIN 8074  |  Gentex", 1, 9, sz=13, ht=28)
ws_db.merge_cells('A1:I1')
note_c = ws_db.cell(row=2, column=1,
    value=("Source: Gentex HDPE Pipes – PE 100 – DIN 8074 Price List (effective July 6, 2018)  |  "
           "OD 20–90mm: 100m rolls  |  OD 110mm: 50m rolls  |  OD ≥125mm: 6m or 12m lengths"))
note_c.font = FN(sz=9, it=True, c='595959'); note_c.alignment = AL(h='left')
ws_db.merge_cells('A2:I2'); ws_db.row_dimensions[2].height = 16

col_hdr_row(ws_db, 3, ["Item Code","OD (mm)","PN Rating","PN (Numeric)",
                         "Internal Dia. ID (mm)","Wall Thickness (mm)",
                         "Weight (kg/m)","Price (UGX/m)","Supply Form"], 1, ht=30)

for i, row_d in df_p.iterrows():
    r = DB_R1 + i
    od = int(row_d['OD'])
    form = "100m roll" if od <= 90 else ("50m roll" if od == 110 else "6m / 12m lengths")
    vals = [int(row_d['Code']), od, str(row_d['PN_label']).strip(), int(row_d['PN_num']),
            float(row_d['ID']), float(row_d['Wall']), float(row_d['Weight']),
            float(row_d['Price']) if pd.notna(row_d['Price']) else None, form]
    fmts = ['0','0','@','0','0.00','0.00','0.000','#,##0','@']
    alt  = "F7F9FC" if i % 2 == 0 else "FFFFFF"
    for j, (v, fm) in enumerate(zip(vals, fmts)):
        c = ws_db.cell(row=r, column=j+1, value=v)
        c.font = FN(sz=9); c.number_format = fm
        c.border = BD('thin'); c.alignment = AL(); c.fill = FL(alt)
    ws_db.row_dimensions[r].height = 15

ws_db.freeze_panes = 'A4'
cw(ws_db, {'A':12,'B':10,'C':12,'D':14,'E':18,'F':20,'G':14,'H':16,'I':18})

print("✅ Stage 2 complete — COVER, PARAMETERS, PIPE_DB sheets built.")

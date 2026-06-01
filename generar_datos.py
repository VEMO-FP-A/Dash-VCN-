"""
VEMO VCN Dashboard — Generador de datos
========================================
Uso:
    py generar_datos.py "VEMO 2026 VCN FY Financials FV dynamics.xlsx"

Genera: data.json (solo este archivo se sube a GitHub cada mes)

Descarga Python: https://www.python.org/downloads/
Instalar dependencias: py -m pip install pandas openpyxl
"""

import pandas as pd
import json
import sys
import os
from datetime import datetime

MONTH_ES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

def gv(df, row, cols, div=1e6):
    return [round(float(df.iloc[row,c])/div,3)
            if pd.notna(df.iloc[row,c]) and isinstance(df.iloc[row,c],(int,float)) else 0.0
            for c in cols]

def find_row(df, label):
    for i,l in enumerate(df.iloc[:,1].tolist()):
        if pd.notna(l) and str(l).strip()==label:
            return i
    return None

def find_row_col2(df, label):
    """Busca en columna 2 (para hojas con estructura diferente)."""
    for i,l in enumerate(df.iloc[:,2].tolist()):
        if pd.notna(l) and str(l).strip()==label:
            return i
    return None

def build_ch(df, rows_dict, cols):
    result = {}
    for label,row in rows_dict.items():
        if row is not None:
            vals = gv(df, row, cols)
            if any(v!=0 for v in vals):
                result[label] = vals
    return result

def detect_cols_vcn(df):
    """Detecta las columnas de meses reales en VCN Board View."""
    cols = []
    for c in range(17, df.shape[1]):
        date_v = df.iloc[5,c]
        rev_row = find_row(df,'Revenue') or 162
        rev_v   = df.iloc[rev_row,c]
        if (pd.notna(date_v) and pd.notna(rev_v) and
                isinstance(rev_v,(int,float)) and abs(float(rev_v))>1000):
            cols.append(c)
        elif cols:
            break
    return cols

def make_labels(df, cols):
    labels = []
    for c in cols:
        v = df.iloc[5,c]
        if pd.notna(v):
            try:
                d = pd.to_datetime(v)
                labels.append(f"{MONTH_ES[d.month-1]}-{str(d.year)[2:]}")
            except:
                labels.append(f"M{c}")
        else:
            labels.append(f"M{c}")
    return labels

def procesar_vcn(df, cols):
    print("  Procesando VCN Board View...")
    rows = {
        'Revenue':   find_row(df,'Revenue')   or 162,
        'COGS':      find_row(df,'COGS')       or 172,
        'GP':        find_row(df,'Gross Profit') or 182,
        'Opex':      find_row(df,'Opex')       or 180,
        'SGA':       find_row(df,'Direct Business Lines SG&A') or 211,
        'CrossSGA':  find_row(df,'Cross Functional SG&A') or 213,
        'CorpGA':    find_row(df,'Corporate G&A') or 214,
        'CorpBonus': find_row(df,'Corporate Bonus') or 215,
        'EBITDA':    find_row(df,'EBITDA')     or 218,
        'DA':        find_row(df,'D&A')        or 245,
        'Interest':  find_row(df,'Interest expense') or 247,
        'NI':        find_row(df,'Net Income') or 254,
        'NonOp':     find_row(df,'Non Operating Revenue / (Expense)') or 239,
    }
    print(f"    Filas: {rows}")

    rev_ch = {
        'LTO Users':         find_row(df,'LTO Users (Watts)'),
        'LTO Adicional':     find_row(df,'LTO Additional Consumption '),
        'B2C Users':         find_row(df,'B2C Users (Watts)'),
        'DiDi PPC':          find_row(df,'DiDi Charging Agreement PPC '),
        'DiDi Drivers':      find_row(df,'DiDi Independent Drivers'),
        'DAE Hubs':          find_row(df,'DAE Chaas Fee '),
        'DAE: Energy Cost Reimbursment': find_row(df,'DAE: Energy Cost Reimbursment'),
        'Others Revenue':    find_row(df,'Others Revenue'),
    }
    cogs_ch = {
        'Energy Costs':      find_row(df,'Energy Costs'),
        'Energy DAE':        find_row(df,'Energy cost asociated with DAE consumption'),
        'Revenue Share':     find_row(df,'Revenue Share '),
        'Bank & Fees':       find_row(df,'Bank & Others Fees '),
        'Chargebacks':       find_row(df,'Chargebacks'),
        'Charger Insurance': find_row(df,'Charger Insurance '),
        'O&M':               find_row(df,'O&M'),
        'Maintenance Hubs':  find_row(df,'Maintenance Hubs'),
    }
    opex_ch = {
        'Cleaning':              find_row(df,'Outside services (cleaning ) '),
        'Security':              find_row(df,'Security'),
        'HUB Utilities':         find_row(df,'HUB Utilities VCN '),
        'Non-Executed Projects': find_row(df,'Non-Executed Projects '),
        'Rent Hubs':             find_row(df,'Total Rent'),
        'D&A Hubs':              find_row(df,'D&A Hubs'),
    }
    sga_ch = {
        'Payroll Admin':      find_row(df,'Payroll Expense Administrative'),
        'Payroll Staff Hubs': find_row(df,'Payroll Expense: Staff Anchored Hubs'),
        'MKT':                find_row(df,'MKT Expense'),
        'Advisory Fees':      find_row(df,'Advisory Fees'),
        'Insurance':          find_row(df,'Insurance'),
        'Subscriptions':      find_row(df,'Subscriptions'),
        'Financings':         find_row(df,'Financings'),
        'T&E':                find_row(df,'T&E'),
        'Others':             find_row(df,'Others'),
    }

    def v(k): return gv(df, rows[k], cols)

    revenue   = v('Revenue'); cogs = v('COGS'); gp = v('GP'); opex = v('Opex')
    sga = v('SGA'); cross = v('CrossSGA'); corp = v('CorpGA'); bonus = v('CorpBonus')
    ebitda = v('EBITDA'); da = v('DA'); interest = v('Interest')
    ni = v('NI'); nonop = v('NonOp')

    gpm       = [round(gp[i]/revenue[i]*100,1) if revenue[i]!=0 else 0.0 for i in range(len(gp))]
    ebitda_adj= [round(ebitda[i]-cross[i]-corp[i]-bonus[i],3) for i in range(len(ebitda))]
    ni_adj    = [round(ni[i]-nonop[i],3) for i in range(len(ni))]

    print(f"    Último mes: Rev={revenue[-1]}, GP={gp[-1]}, EBITDA={ebitda[-1]}, NI={ni[-1]}")

    return {
        'Revenue':   {'vals':revenue,    'children':build_ch(df,rev_ch,cols)},
        'COGS':      {'vals':cogs,       'children':build_ch(df,cogs_ch,cols)},
        'GP':        {'vals':gp,         'children':{}},
        'GPm':       {'vals':gpm,        'children':{}},
        'Opex':      {'vals':opex,       'children':build_ch(df,opex_ch,cols)},
        'SGA':       {'vals':sga,        'children':build_ch(df,sga_ch,cols)},
        'CrossSGA':  {'vals':cross,      'children':{}},
        'CorpGA':    {'vals':corp,       'children':{}},
        'CorpBonus': {'vals':bonus,      'children':{}},
        'EBITDA':    {'vals':ebitda,     'children':{}},
        'EBITDAadj': {'vals':ebitda_adj, 'children':{}},
        'DA':        {'vals':da,         'children':{}},
        'Interest':  {'vals':interest,   'children':{}},
        'NI':        {'vals':ni,         'children':{}},
        'NIadj':     {'vals':ni_adj,     'children':{}},
        'NonOp':     {'vals':nonop,      'children':{}},
        'CrossCorp': {
            'vals':[round(sga[i]+cross[i]+corp[i]+bonus[i],3) for i in range(len(sga))],
            'children':{
                'Direct SG&A':sga, 'Cross Functional SG&A':cross,
                'Corporate G&A':corp, 'Corporate Bonus':bonus,
            }
        },
    }

def procesar_budget(df_bud, cols_bud):
    print("  Procesando Budget...")

    def gvb(r):
        if r >= df_bud.shape[0]: return [0]*len(cols_bud)
        return gv(df_bud, r, cols_bud)

    # Detect rows in budget sheet by label
    rows_b = {
        'Revenue':  find_row(df_bud,'Revenue')                  or find_row(df_bud,'Revenue'),
        'COGS':     find_row(df_bud,'COGS'),
        'GP':       find_row(df_bud,'Gross Profit'),
        'Opex':     find_row(df_bud,'OPEX')                     or find_row(df_bud,'Opex'),
        'SGA':      find_row(df_bud,'Direct Business Lines SG&A'),
        'CrossSGA': find_row(df_bud,'Cross Functional SG&A'),
        'CorpGA':   find_row(df_bud,'Corporate G&A'),
        'CorpBonus':find_row(df_bud,'Corporate Bonus'),
        'EBITDA':   find_row(df_bud,'EBITDA'),
        'NI':       find_row(df_bud,'Net Income'),
    }
    print(f"    Budget rows: {rows_b}")

    def safe(k, default=0):
        r = rows_b.get(k)
        return gvb(r) if r is not None else [default]*len(cols_bud)

    B_Rev    = safe('Revenue')
    B_COGS   = safe('COGS')
    B_GP     = safe('GP')
    B_Opex   = safe('Opex')
    B_SGA    = safe('SGA')
    B_Cross  = safe('CrossSGA')
    B_Corp   = safe('CorpGA')
    B_Bonus  = safe('CorpBonus')
    B_EBITDA = safe('EBITDA')
    B_NI     = safe('NI')

    n = len(cols_bud)
    B_adj = [round(B_EBITDA[i]-B_Cross[i]-B_Corp[i]-B_Bonus[i],3) for i in range(n)]

    # Children — safe fallbacks for rows that may not exist
    def safe_row(r, default=None):
        if r is None or r >= df_bud.shape[0]: return [0]*n
        return gvb(r)

    rev_ch  = {'LTO Users':safe_row(find_row(df_bud,'LTO Users (Watts)')),
               'B2C Users':safe_row(find_row(df_bud,'B2C Users (Watts)')),
               'DiDi PPC': safe_row(find_row(df_bud,'PPC')),
               'Others Revenue':safe_row(find_row(df_bud,'Other Revenues'))}
    cogs_ch = {'Energy Costs':safe_row(find_row(df_bud,'Energy Costs')),
               'Revenue Share':safe_row(find_row(df_bud,'Revenue Share ')),
               'Bank & Fees': safe_row(find_row(df_bud,'Bank Fees')),
               'O&M':         safe_row(find_row(df_bud,'O&M & Other OP. Costs'))}
    opex_ch = {'Cleaning':safe_row(find_row(df_bud,'Anchored LTO Hubs: Cleaning Expense')),
               'Security':safe_row(find_row(df_bud,'Anchored LTO Hubs: Security Expense')),
               'Total Rent':safe_row(find_row(df_bud,'Anchored LTO Hubs: Rents'))}
    sga_ch  = {'Payroll Admin':safe_row(find_row(df_bud,'Payroll Expense')),
               'Payroll Staff Hubs':safe_row(find_row(df_bud,'Payroll Expense: Staff Anchored Hubs')),
               'MKT':safe_row(find_row(df_bud,'MKT Expense')),
               'Advisory Fees':safe_row(find_row(df_bud,'Legal and Software Expenses'))}

    def clean(d): return {k:v for k,v in d.items() if any(x!=0 for x in v)}

    print(f"    Budget NI último mes: {B_NI[-1]}, EBITDAadj[0]: {B_adj[0]}")
    return {
        'Revenue':B_Rev,'COGS':B_COGS,'GP':B_GP,'Opex':B_Opex,'SGA':B_SGA,
        'CrossSGA':B_Cross,'CorpGA':B_Corp,'CorpBonus':B_Bonus,
        'EBITDA':B_EBITDA,'EBITDAadj':B_adj,'NI':B_NI,
        '_RevenueCh':clean(rev_ch),'_COGSCh':clean(cogs_ch),
        '_OpexCh':clean(opex_ch),'_SGACh':clean(sga_ch),
    }

def procesar_hoja2(df_h2, cols_h2):
    print("  Procesando Hoja2 / Ops Metrics...")

    def gvh(r, dec=4):
        return [round(float(df_h2.iloc[r,c]),dec)
                if pd.notna(df_h2.iloc[r,c]) and isinstance(df_h2.iloc[r,c],(int,float)) else None
                for c in cols_h2]

    # Try to find rows dynamically, fallback to known positions
    row_map = {}
    labels2 = [str(df_h2.iloc[r,1]).strip() if pd.notna(df_h2.iloc[r,1]) else '' for r in range(df_h2.shape[0])]

    # Map of section names to find
    section_search = {
        'Portafolio Installed Capacity': 14,
        'Portafolio Contracted Capacity': 34,
        '3) Portafolio Connectors': 56,
    }

    return {
        'inst_active':gvh(16),'inst_anchored':gvh(17),'inst_pureDest':gvh(18),
        'inst_rhf':gvh(19),'inst_dae':gvh(20),
        'inst_newInst':gvh(26),'inst_newAnch':gvh(27),'inst_newPure':gvh(28),
        'inst_newRhf':gvh(29),'inst_backlog':gvh(30),'inst_backAnch':gvh(31),
        'inst_backPure':gvh(32),'inst_backRhf':gvh(33),
        'cont_active':gvh(36),'cont_anchored':gvh(37),'cont_pureDest':gvh(38),
        'cont_rhf':gvh(39),'cont_dae':gvh(40),
        'cont_newInst':gvh(41),'cont_newAnch':gvh(42),'cont_newPure':gvh(43),
        'cont_newRhf':gvh(44),'cont_backlog':gvh(45),'cont_backAnch':gvh(46),
        'cont_backPure':gvh(47),'cont_backRhf':gvh(48),
        'conn_active':gvh(56),'conn_anchored':gvh(57),'conn_pureDest':gvh(58),
        'conn_rhf':gvh(59),'conn_dae':gvh(60),
        'conn_newInst':gvh(61),'conn_newAnch':gvh(62),'conn_newPure':gvh(63),
        'conn_newRhf':gvh(64),'conn_backlog':gvh(65),'conn_backAnch':gvh(66),
        'conn_backPure':gvh(67),'conn_backRhf':gvh(68),
        'irr_active':gvh(76),'irr_anchored':gvh(77),'irr_pureDest':gvh(78),'irr_rhf':gvh(79),
        'thru_active':gvh(83),'thru_anchored':gvh(84),'thru_pureDest':gvh(85),
        'thru_rhf':gvh(86),'thru_dae':gvh(87),
        'util_active':gvh(91),'util_anchored':gvh(92),'util_pureDest':gvh(93),
        'util_rhf':gvh(94),'util_dae':gvh(95),
        'rev_total':gvh(99),'rev_anchored':gvh(100),'rev_pureDest':gvh(101),'rev_rhf':gvh(102),
        'revkwh_total':gvh(106),'revkwh_anchored':gvh(107),'revkwh_pureDest':gvh(108),'revkwh_rhf':gvh(109),
        'gpkwh_total':gvh(113),'gpmargin_total':gvh(120),
        'revconn_total':gvh(127),'revconn_anchored':gvh(128),'revconn_pureDest':gvh(129),'revconn_rhf':gvh(130),
        'uptime_total':gvh(134),'uptime_anchored':gvh(135),'uptime_pureDest':gvh(136),'uptime_rhf':gvh(137),
        'capex_total':gvh(141),'capex_anchored':gvh(142),'capex_pureDest':gvh(143),'capex_rhf':gvh(144),
        'evr_anchored':gvh(148),'evr_pureDest':gvh(149),
    }

def main():
    # ── Archivo ──
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        excels = sorted([f for f in os.listdir('.') if f.endswith('.xlsx') and
                        any(x in f for x in ['VEMO','Dash','VCN'])])
        if not excels:
            print("ERROR: No se encontró archivo Excel.")
            print("Uso: py generar_datos.py nombre_archivo.xlsx")
            sys.exit(1)
        excel_file = excels[-1]
        print(f"Usando archivo: {excel_file}")

    print(f"\n{'='*55}")
    print(f"VEMO VCN — Generando data.json")
    print(f"Archivo: {excel_file}")
    print(f"{'='*55}\n")

    xl = pd.read_excel(excel_file, sheet_name=None, header=None)
    sheets = list(xl.keys())
    print(f"Hojas: {sheets}\n")

    # Detectar hojas por nombre
    vcn_sheet   = next((s for s in sheets if 'Board View' in s and 'Detail' in s), sheets[0])
    bud_sheet   = next((s for s in sheets if 'Budget' in s and 'Board' in s),
                  next((s for s in sheets if s.strip() == 'Budget'),
                  next((s for s in sheets if 'Budget' in s), sheets[1])))
    ops_sheet   = next((s for s in sheets if 'Hoja2' in s or 'Hoja1' in s), None)

    print(f"VCN Board View : '{vcn_sheet}'")
    print(f"Budget sheet   : '{bud_sheet}'")
    print(f"Ops sheet      : '{ops_sheet}'\n")

    df_vcn = xl[vcn_sheet]
    df_bud = xl[bud_sheet]
    df_ops = xl[ops_sheet] if ops_sheet else None

    # Detectar columnas actuals
    cols_vcn = detect_cols_vcn(df_vcn)
    n_months = len(cols_vcn)
    print(f"Meses actuals detectados: {n_months} (cols {cols_vcn[0]}-{cols_vcn[-1]})")

    mo_labels = make_labels(df_vcn, cols_vcn)
    last_month = mo_labels[-1]
    print(f"Meses: {mo_labels}\n")

    # Budget cols
    cols_bud = list(range(29, 41))

    # Ops cols — only if sheet has enough columns
    ops_cols = []
    if df_ops is not None and df_ops.shape[1] > 30:
        for c in range(26, min(33, df_ops.shape[1])):
            if any(pd.notna(df_ops.iloc[r,c]) and isinstance(df_ops.iloc[r,c],(int,float)) and df_ops.iloc[r,c]!=0
                   for r in [16,56,83] if r < df_ops.shape[0]):
                ops_cols.append(c)

    ops_months_labels = make_labels(df_ops, ops_cols) if df_ops is not None and ops_cols else mo_labels[-5:]

    # Procesar
    D   = procesar_vcn(df_vcn, cols_vcn)
    B   = procesar_budget(df_bud, cols_bud)
    ops = procesar_hoja2(df_ops, ops_cols) if df_ops is not None and ops_cols else {}

    # Empaquetar
    output = {
        'meta': {
            'generated':   datetime.now().strftime('%Y-%m-%d %H:%M'),
            'last_month':  last_month,
            'n_months':    n_months,
            'source_file': os.path.basename(excel_file),
        },
        'months':      mo_labels,
        'ops_months':  ops_months_labels,
        'D':           D,
        'B':           B,
        'ops':         ops,
    }

    out_file = 'data.json'
    with open(out_file, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    size = os.path.getsize(out_file)/1024
    print(f"\n{'='*55}")
    print(f"✅  Generado: {out_file} ({size:.1f} KB)")
    print(f"    Último mes : {last_month}")
    print(f"    Meses      : {n_months}")
    print(f"{'='*55}")
    print("\n➡️  Sube data.json a GitHub y el dashboard se actualiza.")

if __name__ == '__main__':
    main()

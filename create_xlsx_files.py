"""
Creates Excel workbooks using xlsxwriter.

Every formula cell has BOTH:
  - The live Excel formula (visible in the formula bar, recalculates on F9)
  - A pre-computed cached value (displays immediately, no Protected-View issue)

treasury_rates.xlsx
  Col A   : date
  Col B-D : DGS2, DGS5, DGS10 (raw FRED data)
  Col E   : butterfly_spread  =2*DGS5-DGS2-DGS10
  Col F   : roll_mean_252     rolling 252-day mean (blank first 251 rows)
  Col G   : roll_std_252      rolling 252-day std  (blank first 251 rows)
  Col H   : z_score_252       (spread-mean)/std    (blank first 251 rows)

butterfly_trades.xlsx
  Sheet "Long Belly Trades"   36 columns, Section 1-5
  Sheet "Summary"             P&L summary
"""

import math
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as CN

# ---------------------------------------------------------------------------
# Colours
# ---------------------------------------------------------------------------
C_INP_BG   = '#D9E1F2'   # light blue  -- raw input
C_FML_BG   = '#E2EFDA'   # light green -- formula / derived
C_PNL_BG   = '#FFF2CC'   # light yellow -- P&L / constants
C_CARRY_BG = '#FCE4D6'   # light orange -- carry detail
C_TOT_BG   = '#D6DCE4'   # light grey  -- totals
C_WIN_BG   = '#C6EFCE'   # bright green -- winning trade
C_LOSS_BG  = '#FFC7CE'   # bright red  -- losing trade
C_DK_BLUE  = '#1F4E79'   # dark blue   -- main header / text accent
C_DK_GREEN = '#375623'   # dark green  -- long belly
C_BORDER   = '#AAAAAA'   # cell border

# ---------------------------------------------------------------------------
# Column layout for butterfly trade sheet (0-indexed for xlsxwriter)
# ---------------------------------------------------------------------------
COL = {
    "entry_date":        0,   # A
    "entry_DGS2":        1,   # B
    "entry_DGS5":        2,   # C
    "entry_DGS10":       3,   # D
    "entry_z_score":     4,   # E
    "exit_date":         5,   # F
    "exit_DGS2":         6,   # G
    "exit_DGS5":         7,   # H
    "exit_DGS10":        8,   # I
    "exit_z_score":      9,   # J
    "hold_days":         10,  # K
    "avg_repo_rate_pct": 11,  # L
    "actual_avg_carry":  12,  # M
    "delta_dgs2_bps":    13,  # N  = (G-B)*100
    "delta_dgs5_bps":    14,  # O  = (H-C)*100
    "delta_dgs10_bps":   15,  # P  = (I-D)*100
    "butterfly_chg_bps": 16,  # Q  = 2*O-N-P
    "belly_dv01":        17,  # R  = 45000 (constant)
    "wing_dv01":         18,  # S  = 22500 (constant)
    "entry_butterfly":   19,  # T  = 2*C-B-D
    "exit_butterfly":    20,  # U  = 2*H-G-I
    "leg_5y_pnl":        21,  # V  = -R*O
    "leg_2y_pnl":        22,  # W  = +S*N
    "leg_10y_pnl":       23,  # X  = +S*P
    "legs_sum_pnl":      24,  # Y  = V+W+X
    "mtm_pnl":           25,  # Z  = -S*Q
    "belly_notional":    26,  # AA = 100,000,000
    "wing_2y_notional":  27,  # AB = 118,420,000
    "wing_10y_notional": 28,  # AC = 26,470,000
    "belly_carry_day":   29,  # AD = (C/100-L/100)*AA/360
    "wing2y_carry_day":  30,  # AE = (L/100-B/100)*AB/360
    "wing10y_carry_day": 31,  # AF = (L/100-D/100)*AC/360
    "entry_day_carry":   32,  # AG = AD+AE+AF
    "carry_pnl":         33,  # AH = M*K
    "total_pnl":         34,  # AI = Z+AH
    "profitable":        35,  # AJ = IF(AI>0,"Y","N")
}
# Column letter lookup (e.g. CL["entry_DGS2"] = "B")
CL = {name: CN(idx) for name, idx in COL.items()}

BELLY_N  = 100_000_000
WING_2Y  = 118_420_000
WING_10Y =  26_470_000


def _v(val, default=0):
    """Return default if val is None or NaN, otherwise val."""
    if val is None:
        return default
    try:
        if math.isnan(float(val)):
            return default
    except (TypeError, ValueError):
        pass
    return val


# ---------------------------------------------------------------------------
# Format factory
# ---------------------------------------------------------------------------
def _make_formats(wb):
    B = C_BORDER

    def F(**kw):
        kw.setdefault('border', 1)
        kw.setdefault('border_color', B)
        return wb.add_format(kw)

    DT = C_DK_BLUE  # dark text for derived cells

    return {
        # --- input (light blue) ---
        'inp':       F(bg_color=C_INP_BG),
        'inp_d3':    F(bg_color=C_INP_BG, num_format='0.000'),
        'inp_int':   F(bg_color=C_INP_BG, num_format='#,##0'),

        # --- formula/derived (light green, italic dark blue) ---
        'fml_d4':    F(bg_color=C_FML_BG, italic=True, font_color=DT, num_format='0.0000'),
        'fml_d3':    F(bg_color=C_FML_BG, italic=True, font_color=DT, num_format='0.000'),
        'fml_d1':    F(bg_color=C_FML_BG, italic=True, font_color=DT, num_format='0.0'),
        'fml_int':   F(bg_color=C_FML_BG, italic=True, font_color=DT, num_format='#,##0'),

        # --- P&L / constant (light yellow) ---
        'pnl_cst':   F(bg_color=C_PNL_BG, num_format='#,##0'),
        'pnl':       F(bg_color=C_PNL_BG, italic=True, font_color=DT, num_format='#,##0'),

        # --- carry (light orange) ---
        'carry_cst': F(bg_color=C_CARRY_BG, num_format='#,##0'),
        'carry':     F(bg_color=C_CARRY_BG, italic=True, font_color=DT, num_format='#,##0'),

        # --- totals / win / loss ---
        'tot':       F(bg_color=C_TOT_BG,  italic=True, font_color=DT, num_format='#,##0'),
        'win':       F(bg_color=C_WIN_BG,  italic=True, font_color=DT, num_format='#,##0'),
        'loss':      F(bg_color=C_LOSS_BG, italic=True, font_color=DT, num_format='#,##0'),
        'win_wl':    F(bg_color=C_WIN_BG,  bold=True, align='center'),
        'loss_wl':   F(bg_color=C_LOSS_BG, bold=True, align='center'),

        # --- column headers (row 3 in trade sheet / row 2 in treasury) ---
        'hdr_inp':   F(bg_color=C_INP_BG,   bold=True, font_color=DT, align='center', valign='vcenter', text_wrap=True),
        'hdr_fml':   F(bg_color=C_FML_BG,   bold=True, font_color=DT, align='center', valign='vcenter', text_wrap=True),
        'hdr_pnl':   F(bg_color=C_PNL_BG,   bold=True, font_color=DT, align='center', valign='vcenter', text_wrap=True),
        'hdr_carry': F(bg_color=C_CARRY_BG, bold=True, font_color=DT, align='center', valign='vcenter', text_wrap=True),
        'hdr_tot':   F(bg_color=C_TOT_BG,   bold=True, font_color=DT, align='center', valign='vcenter', text_wrap=True),

        # --- section banners (light bg, dark blue bold text) ---
        'ban_inp':   F(bg_color=C_INP_BG,   bold=True, font_color=DT, align='center', valign='vcenter'),
        'ban_fml':   F(bg_color=C_FML_BG,   bold=True, font_color=DT, align='center', valign='vcenter'),
        'ban_pnl':   F(bg_color=C_PNL_BG,   bold=True, font_color=DT, align='center', valign='vcenter'),
        'ban_carry': F(bg_color=C_CARRY_BG, bold=True, font_color=DT, align='center', valign='vcenter'),
        'ban_tot':   F(bg_color=C_TOT_BG,   bold=True, font_color=DT, align='center', valign='vcenter'),

        # --- top banners (dark bg, white bold text) ---
        'ban_dark':  F(bg_color=C_DK_BLUE,  bold=True, font_color='#FFFFFF', align='center', valign='vcenter'),
        'ban_long':  F(bg_color=C_DK_GREEN, bold=True, font_color='#FFFFFF', align='center', valign='vcenter'),

        # --- misc ---
        'note':      F(italic=True, font_size=9, text_wrap=True, border=0),
        'sum_hdr':   F(bg_color=C_DK_BLUE,  bold=True, font_color='#FFFFFF', align='center'),
        'sum_long':  F(bg_color=C_DK_GREEN, bold=True, font_color='#FFFFFF', align='center'),
        'sum_cell':  F(num_format='#,##0', align='right'),
        'sum_bold':  F(bold=True, num_format='#,##0', align='right'),
        'sum_label': F(),
    }


# ===========================================================================
# 1. treasury_rates.xlsx
# ===========================================================================

def build_treasury_xlsx():
    df  = pd.read_csv("treasury_rates.csv")
    wb  = xlsxwriter.Workbook("treasury_rates.xlsx")
    ws  = wb.add_worksheet("Treasury Rates")
    fmts = _make_formats(wb)

    # Column widths (0-indexed)
    ws.set_column(0, 0, 13)   # A: date
    ws.set_column(1, 4, 11)   # B-E
    ws.set_column(5, 7, 15)   # F-H: rolling stats

    # Row 0: top banner
    ws.set_row(0, 20)
    ws.merge_range(0, 0, 0, 3, "RAW DATA  (source: FRED API)", fmts['ban_dark'])
    ws.merge_range(0, 4, 0, 7,
        "CALCULATED  --  live Excel formulas + cached values  "
        "|  rolling stats blank for first 251 rows (need 252 obs)", fmts['ban_long'])

    # Row 1: column headers
    ws.set_row(1, 48)
    hdr_cols = [
        (0, "date",                                              fmts['hdr_inp']),
        (1, "DGS2\n(%)",                                         fmts['hdr_inp']),
        (2, "DGS5\n(%)",                                         fmts['hdr_inp']),
        (3, "DGS10\n(%)",                                        fmts['hdr_inp']),
        (4, "butterfly_spread\n=2*DGS5-DGS2-DGS10",             fmts['hdr_fml']),
        (5, "roll_mean_252\n=252d rolling avg\n(blank 1st 251)", fmts['hdr_fml']),
        (6, "roll_std_252\n=252d rolling std\n(blank 1st 251)",  fmts['hdr_fml']),
        (7, "z_score_252\n=(spread-mean)/std\n(blank 1st 251)", fmts['hdr_fml']),
    ]
    for col, text, fmt in hdr_cols:
        ws.write(1, col, text, fmt)

    # Rolling-stat formulas (same formula string every row; ROW() is evaluated per cell)
    # ROW()-2 = number of data points (row 2 = header → first data is row 3 → ROW()-2=1)
    mean_fml = '=IF(ROW()-2<252,"",AVERAGE(OFFSET($E$2,ROW()-253,0,252,1)))'
    std_fml  = '=IF(ROW()-2<252,"",STDEV(OFFSET($E$2,ROW()-253,0,252,1)))'

    # Data rows start at 0-indexed row 2 (Excel row 3)
    for ri, row in enumerate(df.itertuples(index=False), start=2):
        er = ri + 1  # Excel 1-indexed row number for cell references

        ws.write(ri, 0, row.date,  fmts['inp'])
        ws.write(ri, 1, row.DGS2,  fmts['inp_d3'])
        ws.write(ri, 2, row.DGS5,  fmts['inp_d3'])
        ws.write(ri, 3, row.DGS10, fmts['inp_d3'])

        # Col E: butterfly_spread
        bfly = _v(row.butterfly_spread, 0)
        ws.write_formula(ri, 4, f"=2*C{er}-B{er}-D{er}", fmts['fml_d4'], bfly)

        # Col F-H: rolling stats -- blank cached value for first 251 data rows
        has_252 = (ri >= 253)

        mean_v = _v(row.roll_mean_252, "") if has_252 else ""
        ws.write_formula(ri, 5, mean_fml, fmts['fml_d4'], mean_v)

        std_v = _v(row.roll_std_252, "") if has_252 else ""
        ws.write_formula(ri, 6, std_fml, fmts['fml_d4'], std_v)

        z_fml = f'=IF(F{er}="","",IF(G{er}=0,"",(E{er}-F{er})/G{er}))'
        z_v   = _v(row.z_score_252, "") if has_252 else ""
        ws.write_formula(ri, 7, z_fml, fmts['fml_d3'], z_v)

    ws.freeze_panes(2, 0)  # freeze rows 1-2 (0-indexed)
    wb.close()
    print(f"Saved treasury_rates.xlsx  ({len(df)} rows, cols E-H: formula + cached value)")


# ===========================================================================
# 2. butterfly_trades.xlsx
# ===========================================================================

def _write_trade_sheet(ws, fmts, df):
    """Write the Long Belly trade sheet."""
    dir_label = "LONG BELLY  (Long $100M 5Y / Short $118.4M 2Y / Short $26.5M 10Y)"

    # -- Column widths (0-indexed first_col, last_col, width)
    ws.set_column(0,  0,  13)   # A: entry date
    ws.set_column(1,  4,  10)   # B-E: entry yields + z
    ws.set_column(5,  5,  13)   # F: exit date
    ws.set_column(6,  9,  10)   # G-J: exit yields + z
    ws.set_column(10, 10,  8)   # K: hold days
    ws.set_column(11, 11, 10)   # L: repo rate
    ws.set_column(12, 12, 14)   # M: actual carry
    ws.set_column(13, 15, 11)   # N-P: delta bps
    ws.set_column(16, 16, 14)   # Q: butterfly chg bps
    ws.set_column(17, 18, 12)   # R-S: DV01 constants
    ws.set_column(19, 25, 14)   # T-Z: entry/exit butterfly + leg P&Ls + MtM
    ws.set_column(26, 28, 14)   # AA-AC: notionals
    ws.set_column(29, 33, 16)   # AD-AH: carry detail
    ws.set_column(34, 35, 14)   # AI-AJ: total + W/L

    # -- Row 0: direction banner
    ws.set_row(0, 22)
    ws.merge_range(0, 0, 0, 35, dir_label, fmts['ban_long'])

    # -- Row 1: section banners
    ws.set_row(1, 18)
    ws.merge_range(1,  0, 1, 12, "SECTION 1 -- INPUT DATA",    fmts['ban_inp'])
    ws.merge_range(1, 13, 1, 16, "SECTION 2 -- YIELD CHANGES", fmts['ban_fml'])
    ws.merge_range(1, 17, 1, 25, "SECTION 3 -- MtM P&L",       fmts['ban_pnl'])
    ws.merge_range(1, 26, 1, 33, "SECTION 4 -- CARRY DETAIL",  fmts['ban_carry'])
    ws.merge_range(1, 34, 1, 35, "SECTION 5 -- TOTAL",         fmts['ban_tot'])

    # -- Row 2: column headers
    ws.set_row(2, 56)

    hdr_cols = [
        (COL["entry_date"],        "Entry Date",                  fmts['hdr_inp']),
        (COL["entry_DGS2"],        "Entry\nDGS2 (%)",             fmts['hdr_inp']),
        (COL["entry_DGS5"],        "Entry\nDGS5 (%)",             fmts['hdr_inp']),
        (COL["entry_DGS10"],       "Entry\nDGS10 (%)",            fmts['hdr_inp']),
        (COL["entry_z_score"],     "Entry Z\n(>= +2.0)",          fmts['hdr_inp']),
        (COL["exit_date"],         "Exit Date",                   fmts['hdr_inp']),
        (COL["exit_DGS2"],         "Exit\nDGS2 (%)",              fmts['hdr_inp']),
        (COL["exit_DGS5"],         "Exit\nDGS5 (%)",              fmts['hdr_inp']),
        (COL["exit_DGS10"],        "Exit\nDGS10 (%)",             fmts['hdr_inp']),
        (COL["exit_z_score"],      "Exit Z\n(<= 0.0)",            fmts['hdr_inp']),
        (COL["hold_days"],         "Hold\nDays",                  fmts['hdr_inp']),
        (COL["avg_repo_rate_pct"], "Avg SOFR\n(%)",               fmts['hdr_inp']),
        (COL["actual_avg_carry"],  "Actual Avg\nDaily Carry $",   fmts['hdr_inp']),
        (COL["delta_dgs2_bps"],    "dDGS2 bps\n=(G-B)*100",       fmts['hdr_fml']),
        (COL["delta_dgs5_bps"],    "dDGS5 bps\n=(H-C)*100",       fmts['hdr_fml']),
        (COL["delta_dgs10_bps"],   "dDGS10 bps\n=(I-D)*100",      fmts['hdr_fml']),
        (COL["butterfly_chg_bps"], "Bfly Chg bps\n=2*O-N-P",     fmts['hdr_fml']),
        (COL["belly_dv01"],        "Belly DV01\n[=$45,000]",      fmts['hdr_pnl']),
        (COL["wing_dv01"],         "Wing DV01\n[=$22,500]",       fmts['hdr_pnl']),
        (COL["entry_butterfly"],   "Entry Bfly\n=2*C-B-D",        fmts['hdr_fml']),
        (COL["exit_butterfly"],    "Exit Bfly\n=2*H-G-I",         fmts['hdr_fml']),
        (COL["leg_5y_pnl"],        "Leg 5Y P&L\n=-DV01b*dDGS5",  fmts['hdr_pnl']),
        (COL["leg_2y_pnl"],        "Leg 2Y P&L\n=+DV01w*dDGS2",  fmts['hdr_pnl']),
        (COL["leg_10y_pnl"],       "Leg 10Y P&L\n=+DV01w*dDGS10",fmts['hdr_pnl']),
        (COL["legs_sum_pnl"],      "Legs Sum\n=V+W+X",            fmts['hdr_pnl']),
        (COL["mtm_pnl"],           "MtM P&L\n=-DV01w*BflyCh",    fmts['hdr_pnl']),
        (COL["belly_notional"],    "Belly Notl\n[$100M]",         fmts['hdr_carry']),
        (COL["wing_2y_notional"],  "2Y Notl\n[$118.42M]",         fmts['hdr_carry']),
        (COL["wing_10y_notional"], "10Y Notl\n[$26.47M]",         fmts['hdr_carry']),
        (COL["belly_carry_day"],   "Belly Carry/d\n=(Y5-SOFR)*N5/360",  fmts['hdr_carry']),
        (COL["wing2y_carry_day"],  "2Y Carry/d\n=(SOFR-Y2)*N2/360",     fmts['hdr_carry']),
        (COL["wing10y_carry_day"], "10Y Carry/d\n=(SOFR-Y10)*N10/360",  fmts['hdr_carry']),
        (COL["entry_day_carry"],   "Entry-Day\nCarry Approx/d",   fmts['hdr_carry']),
        (COL["carry_pnl"],         "Carry P&L\n=M*K",             fmts['hdr_carry']),
        (COL["total_pnl"],         "Total P&L\n=Z+AH",            fmts['hdr_tot']),
        (COL["profitable"],        "W/L",                         fmts['hdr_tot']),
    ]
    for ci, label, fmt in hdr_cols:
        ws.write(2, ci, label, fmt)

    # -- Data rows (0-indexed starting at row 3; Excel row = ri+1)
    for ri, row in enumerate(df.itertuples(index=False), start=3):
        er = ri + 1        # Excel 1-indexed row number
        is_win = (row.profitable == "Y")

        # Shortcuts for column letters used in formulas
        B  = CL["entry_DGS2"];  C_ = CL["entry_DGS5"];  D  = CL["entry_DGS10"]
        G  = CL["exit_DGS2"];   H  = CL["exit_DGS5"];   I_ = CL["exit_DGS10"]
        N  = CL["delta_dgs2_bps"]
        O  = CL["delta_dgs5_bps"]
        P  = CL["delta_dgs10_bps"]
        Q  = CL["butterfly_chg_bps"]
        R  = CL["belly_dv01"];   S_ = CL["wing_dv01"]
        V  = CL["leg_5y_pnl"];   W  = CL["leg_2y_pnl"]; X  = CL["leg_10y_pnl"]
        L  = CL["avg_repo_rate_pct"]
        K  = CL["hold_days"]
        M  = CL["actual_avg_carry"]
        AA = CL["belly_notional"]; AB = CL["wing_2y_notional"]; AC = CL["wing_10y_notional"]
        AD = CL["belly_carry_day"]; AE = CL["wing2y_carry_day"]; AF = CL["wing10y_carry_day"]
        Z  = CL["mtm_pnl"];  AH = CL["carry_pnl"];  AI = CL["total_pnl"]

        def wv(col_name, val, fmt):
            ws.write(ri, COL[col_name], val, fmt)

        def wf(col_name, formula, fmt, cached):
            ws.write_formula(ri, COL[col_name], formula, fmt, cached)

        # -- Section 1: raw inputs (written as plain values) --
        wv("entry_date",        row.entry_date,          fmts['inp'])
        wv("entry_DGS2",        row.entry_DGS2,          fmts['inp_d3'])
        wv("entry_DGS5",        row.entry_DGS5,          fmts['inp_d3'])
        wv("entry_DGS10",       row.entry_DGS10,         fmts['inp_d3'])
        wv("entry_z_score",     row.entry_z_score,       fmts['inp_d3'])
        wv("exit_date",         row.exit_date,           fmts['inp'])
        wv("exit_DGS2",         row.exit_DGS2,           fmts['inp_d3'])
        wv("exit_DGS5",         row.exit_DGS5,           fmts['inp_d3'])
        wv("exit_DGS10",        row.exit_DGS10,          fmts['inp_d3'])
        wv("exit_z_score",      row.exit_z_score,        fmts['inp_d3'])
        wv("hold_days",         row.hold_days,           fmts['inp'])
        wv("avg_repo_rate_pct", row.avg_repo_rate_pct,   fmts['inp_d3'])
        wv("actual_avg_carry",  row.avg_daily_carry_usd, fmts['inp_int'])

        # -- Section 2: yield changes (formula + cached) --
        wf("delta_dgs2_bps",    f"=({G}{er}-{B}{er})*100",       fmts['fml_d1'], float(row.delta_dgs2_bps))
        wf("delta_dgs5_bps",    f"=({H}{er}-{C_}{er})*100",      fmts['fml_d1'], float(row.delta_dgs5_bps))
        wf("delta_dgs10_bps",   f"=({I_}{er}-{D}{er})*100",      fmts['fml_d1'], float(row.delta_dgs10_bps))
        wf("butterfly_chg_bps", f"=2*{O}{er}-{N}{er}-{P}{er}",   fmts['fml_d1'], float(row.butterfly_chg_bps))

        # -- Section 3: MtM P&L --
        wv("belly_dv01", 45000, fmts['pnl_cst'])
        wv("wing_dv01",  22500, fmts['pnl_cst'])

        wf("entry_butterfly", f"=2*{C_}{er}-{B}{er}-{D}{er}",  fmts['fml_d4'], float(row.entry_butterfly))
        wf("exit_butterfly",  f"=2*{H}{er}-{G}{er}-{I_}{er}",  fmts['fml_d4'], float(row.exit_butterfly))

        wf("leg_5y_pnl",  f"=-{R}{er}*{O}{er}",   fmts['pnl'], float(row.leg_5y_pnl_usd))
        wf("leg_2y_pnl",  f"=+{S_}{er}*{N}{er}",  fmts['pnl'], float(row.leg_2y_pnl_usd))
        wf("leg_10y_pnl", f"=+{S_}{er}*{P}{er}",  fmts['pnl'], float(row.leg_10y_pnl_usd))
        wf("mtm_pnl",     f"=-{S_}{er}*{Q}{er}",  fmts['pnl'], float(row.mtm_pnl_usd))
        wf("legs_sum_pnl", f"={V}{er}+{W}{er}+{X}{er}", fmts['pnl'], float(row.legs_sum_pnl_usd))

        # -- Section 4: carry detail --
        wv("belly_notional",    BELLY_N,  fmts['carry_cst'])
        wv("wing_2y_notional",  WING_2Y,  fmts['carry_cst'])
        wv("wing_10y_notional", WING_10Y, fmts['carry_cst'])

        # Per-leg entry-day carry (computed in Python for cached value)
        repo = row.avg_repo_rate_pct / 100
        y2   = row.entry_DGS2  / 100
        y5   = row.entry_DGS5  / 100
        y10  = row.entry_DGS10 / 100
        bc_val   = (y5  - repo) * BELLY_N  / 360
        w2c_val  = (repo - y2)  * WING_2Y  / 360
        w10c_val = (repo - y10) * WING_10Y / 360
        wf("belly_carry_day",   f"=({C_}{er}/100-{L}{er}/100)*{AA}{er}/360", fmts['carry'], round(bc_val))
        wf("wing2y_carry_day",  f"=({L}{er}/100-{B}{er}/100)*{AB}{er}/360",  fmts['carry'], round(w2c_val))
        wf("wing10y_carry_day", f"=({L}{er}/100-{D}{er}/100)*{AC}{er}/360",  fmts['carry'], round(w10c_val))

        edc_val = bc_val + w2c_val + w10c_val
        wf("entry_day_carry", f"={AD}{er}+{AE}{er}+{AF}{er}", fmts['carry'], round(edc_val))
        wf("carry_pnl",       f"={M}{er}*{K}{er}",            fmts['carry'], float(row.carry_pnl_usd))

        # -- Section 5: total --
        t_fmt  = fmts['win']    if is_win else fmts['loss']
        wl_fmt = fmts['win_wl'] if is_win else fmts['loss_wl']
        wf("total_pnl",  f"={Z}{er}+{AH}{er}",          t_fmt,  float(row.total_pnl_usd))
        wf("profitable", f'=IF({AI}{er}>0,"Y","N")',      wl_fmt, row.profitable)

    # Footer note
    note_row = 3 + len(df)
    ws.merge_range(note_row, 0, note_row, 12,
        'NOTE: Section 2-5 cells contain live Excel formulas (visible in formula bar) '
        'with pre-computed cached values that display immediately. Press F9 to force recalculate.',
        fmts['note'])
    ws.set_row(note_row, 28)

    ws.freeze_panes(3, 0)  # freeze rows 0-2 (direction + section banners + headers)


def _build_summary_sheet(ws, fmts, long_df):
    ws.set_column(0, 0, 34)
    ws.set_column(1, 2, 20)

    ws.merge_range(0, 0, 0, 2,
        "STRATEGY SUMMARY  --  Treasury Butterfly Long Belly Mean-Reversion", fmts['sum_hdr'])

    for ci, (label, fmt) in enumerate(zip(
        ["Metric", "Long Belly", "Jun 2016 - Jun 2026"],
        [fmts['sum_hdr'], fmts['sum_long'], fmts['sum_hdr']]
    )):
        ws.write(1, ci, label, fmt)

    def s(df, col):
        return int(df[col].sum()) if len(df) > 0 else 0
    def n(df):
        return len(df)
    def w(df):
        return int((df["profitable"] == "Y").sum()) if len(df) > 0 else 0

    rows = [
        ("Total trades",        n(long_df),                     ""),
        ("Winning trades",      w(long_df),                     ""),
        ("Losing trades",       n(long_df) - w(long_df),        ""),
        ("Total MtM P&L ($)",   s(long_df, "mtm_pnl_usd"),     ""),
        ("Total Carry P&L ($)", s(long_df, "carry_pnl_usd"),   ""),
        ("Total P&L ($)",       s(long_df, "total_pnl_usd"),   ""),
    ]

    for r_off, (label, lv, _) in enumerate(rows, start=2):
        is_pnl = isinstance(lv, int) and abs(lv) > 999
        is_last = (r_off == 2 + len(rows) - 1)
        val_fmt = fmts['sum_bold'] if (is_last and is_pnl) else (fmts['sum_cell'] if is_pnl else fmts['sum_label'])
        ws.write(r_off, 0, label, fmts['sum_label'])
        ws.write(r_off, 1, lv,    val_fmt)
        ws.write(r_off, 2, "",    fmts['sum_label'])


def build_trades_xlsx():
    long_df = pd.read_csv("butterfly_trades_long.csv") if pd.io.common.file_exists("butterfly_trades_long.csv") else pd.DataFrame()

    wb   = xlsxwriter.Workbook("butterfly_trades.xlsx")
    fmts = _make_formats(wb)

    ws_long = wb.add_worksheet("Long Belly Trades")
    ws_sum  = wb.add_worksheet("Summary")

    if len(long_df) > 0:
        _write_trade_sheet(ws_long, fmts, long_df)
    _build_summary_sheet(ws_sum, fmts, long_df)

    wb.close()
    print(f"Saved butterfly_trades.xlsx  ({len(long_df)} long trades  --  formula + cached value per cell)")


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    build_treasury_xlsx()
    build_trades_xlsx()
    print("\nDone.")
    print("  - All formula cells show pre-computed values immediately on open")
    print("  - Live Excel formulas visible in formula bar; press F9 to recalculate")
    print("  - 252d rolling stats: blank for first 251 rows (require full 252-day window)")
    print("  Legend: Blue=input  Green=derived  Yellow=P&L  Orange=carry  Grey=total")

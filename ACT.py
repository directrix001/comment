"""
ACT.py  –  fully guarded KPI actuals engine  (v2)
--------------------------------------------------
Same four-layer defence as Submit_2.py:

  Layer 1 – Input normalisation  : None paths→"", None DFs→empty DataFrame
  Layer 2 – Presence flags       : HAS_* booleans block all lookups when
                                   source DF or file is absent.
  Layer 3 – Integer guardrail    : _int(v) on EVERY value from a dict/list
                                   lookup before arithmetic → no TypeError.
  Layer 4 – Safe formula builders: return Z ("=0") when row==0 or col==0,
                                   or when required file is missing.

Every possible crash point identified in analysis is covered.
"""

import os
import calendar
import numpy as np
import pandas as pd
from datetime import datetime

from Nissan_Helper_act_function import (
    find_row_numbers_act,
    find_values_pandas_HFM,
    SearchResults,
    generate_excel_formula,
    extract_formula,
    paste_values_KPI_PL,
    relief_search_terms_new,
    find_row_and_column,
    find_row_number_CMFI,
    generate_excel_formula_1,
    find_itp_and_month_value,
    find_second_column_for_outside_profit,
    generate_excel_formula_2,
    goodwill_raw_data_extraction,
    Goodwill_rows,
    paste_values_goodwill,
)
from utils.logger import logger

logger.info("ACT triggered")

# ─────────────────────────────────────────────────────────────────────────────
#  SENTINEL
# ─────────────────────────────────────────────────────────────────────────────
Z = "=0"


# ═══════════════════════════════════════════════════════════════════════════════
#  LAYER 3  –  INTEGER GUARDRAIL
# ═══════════════════════════════════════════════════════════════════════════════

def _int(v) -> int:
    """Convert any value to int safely; return 0 on failure."""
    if v is None:
        return 0
    try:
        return int(v)
    except (TypeError, ValueError):
        logger.warning(f"_int(): cannot convert {v!r} to int – using 0")
        return 0


def _int_list(lst) -> list:
    """Convert every element of a list/tuple to int via _int()."""
    if not lst:
        return [0]
    return [_int(x) for x in lst]


def _row_from(d: dict, key: str) -> int:
    """
    Extract first element of d[key] as int.
    Handles: list value, scalar value, None, missing key.
    Always returns int ≥ 0.
    """
    try:
        val = d.get(key)
        if val is None:
            return 0
        if isinstance(val, (list, tuple)):
            return _int(val[0]) if val else 0
        return _int(val)
    except Exception:
        return 0


# ═══════════════════════════════════════════════════════════════════════════════
#  LAYER 1  –  PRESENCE CHECKERS
# ═══════════════════════════════════════════════════════════════════════════════

def _df_ok(df, name: str) -> bool:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        logger.warning(f"[SKIP] DataFrame '{name}' absent/empty – dependent cells → =0")
        return False
    return True


def _file_ok(path: str, name: str) -> bool:
    if not path or not os.path.isfile(path):
        logger.warning(f"[SKIP] File '{name}' not found ({path!r}) – formulas → =0")
        return False
    return True


# ═══════════════════════════════════════════════════════════════════════════════
#  LAYER 4  –  SAFE FORMULA BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def _safe_gen_formula(row: int, col: int, sheet: str,
                      file_path: str = "") -> str:
    row, col = _int(row), _int(col)
    if not row or not col:
        logger.debug(f"_safe_gen_formula skipped: r={row} c={col} sheet={sheet}")
        return Z
    if file_path and not _file_ok(file_path, f"cross-wb [{sheet}]"):
        return Z
    try:
        return generate_excel_formula(row, col, sheet,
                                      file_path if file_path else None)
    except Exception as e:
        logger.warning(f"generate_excel_formula(r={row},c={col},{sheet}): {e}")
        return Z


def _safe_gen_formula_1(row: int, col: int, sheet: str,
                        file_path: str = "") -> str:
    """Wrapper around generate_excel_formula_1 with same guards."""
    row, col = _int(row), _int(col)
    if not row or not col:
        logger.debug(f"_safe_gen_formula_1 skipped: r={row} c={col} sheet={sheet}")
        return Z
    if file_path and not _file_ok(file_path, f"cross-wb-1 [{sheet}]"):
        return Z
    try:
        return generate_excel_formula_1(row, col, sheet,
                                        file_path if file_path else None)
    except Exception as e:
        logger.warning(f"generate_excel_formula_1(r={row},c={col},{sheet}): {e}")
        return Z


def _safe_extract(file_path: str, sheet: str,
                  row: int, col: int, label: str = "") -> str:
    row, col = _int(row), _int(col)
    if not row or not col:
        return Z
    if not _file_ok(file_path, label or f"template extract({sheet},{row},{col})"):
        return Z
    try:
        result = extract_formula(file_path, sheet, row, col)
        return result if result else Z
    except Exception as e:
        logger.warning(f"extract_formula({sheet},{row},{col}): {e}")
        return Z


def _join(*formulas) -> str:
    """Build '=a+b+…' skipping Z sentinels."""
    parts = [x[1:] for x in formulas if x and x != Z]
    return f"={''.join(parts)}" if parts else Z


def _compound(a: str, op: str, b: str) -> str:
    if a == Z or b == Z:
        return Z
    return f"={a[1:]}{op}{b[1:]}"


# ═══════════════════════════════════════════════════════════════════════════════
#  SAFE COLUMN FINDER  (returns 0 when df absent or 'apr' not found)
# ═══════════════════════════════════════════════════════════════════════════════

def _find_col_months(df: pd.DataFrame, month_number: int, name: str) -> int:
    """
    Locate the column whose row contains 'apr' (case-insensitive),
    then offset by (month_number-4)%12.
    Returns 0 if df is empty or 'apr' is not found.
    """
    if not _df_ok(df, name):
        return 0
    try:
        for _, row in df.iterrows():
            for col_name, value in row.items():
                if str(value).lower() == 'apr':
                    col_idx = _int(df.columns.get_loc(col_name))
                    offset  = (month_number - 4) % 12
                    return col_idx + offset
        logger.warning(f"_find_col_months({name}): 'apr' not found – using 0")
        return 0
    except Exception as e:
        logger.warning(f"_find_col_months({name}): {e} – using 0")
        return 0


def _find_col_months_rd(df: pd.DataFrame, month_number: int, name: str) -> int:
    """Same but looks for 'april' (long form used in R&D sheets)."""
    if not _df_ok(df, name):
        return 0
    try:
        for _, row in df.iterrows():
            for col_name, value in row.items():
                if str(value).lower() == 'april':
                    col_idx = _int(df.columns.get_loc(col_name))
                    offset  = (month_number - 4) % 12
                    return col_idx + offset
        logger.warning(f"_find_col_months_rd({name}): 'april' not found – using 0")
        return 0
    except Exception as e:
        logger.warning(f"_find_col_months_rd({name}): {e} – using 0")
        return 0


# ═══════════════════════════════════════════════════════════════════════════════
#  SAFE GOODWILL ROW EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════

def _gw_row(lst, label: str) -> int:
    """
    Safely extract first element of a goodwill row list, add 2 (as original
    code does), and return as int.  Returns 0 on any failure.
    """
    try:
        if not lst:
            logger.warning(f"Goodwill row list '{label}' empty – using 0")
            return 0
        return _int(lst[0]) + 2
    except Exception as e:
        logger.warning(f"Goodwill row '{label}': {e} – using 0")
        return 0


# ═══════════════════════════════════════════════════════════════════════════════
#  ACT FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

def ACT(
    file_path_P_L, file_path_P_L_1,
    df_P_L_ROP_NE, df_P_L_HQ,
    file_path_OSP, df_OSP,
    file_path_HMF, df_HMF_Managerial,
    file_path_New_KPI, df_P_L_New_KPI,
    file_HQ_Item, df_hQ_items,
    file_goodwill,
    df_r_and_d, df_MZ_conso, df_P_L_Relief,
    df_P_L_PL, df_P_L_Reclass, df_P_L_MCCP, df_P_L_OCS,
    df_New_KPI_CFMI, df_New_goodwill, df_Act,
    df_1c_osp,
    file_paths_High_Level_Variance, df_high_variance_feuil,
    file_paths_Relief_KD, df_Relief_KD,
    file_paths_Purchase_Price, df_Purchase_Price,
    file_paths_UP_Calc_Template_Simulation, df_UP_Calc_Template_Simulation,
    file_paths_NANO_HFM, df_nano_hfm,
    file_paths_ITP, df_itp,
    file_path_Template, file_path_Template_1,
    month_number, Var_Converter, var_Cfmi,
    file_path_Template_man_1, file_path_goodwill_template,
    result, output_path,
):
    logger.info("ACT started")

    # ── Layer 1: normalise all inputs ─────────────────────────────────────────
    def _p(v): return v if isinstance(v, str) else ""
    def _d(v): return v if isinstance(v, pd.DataFrame) else pd.DataFrame()

    file_path_P_L                    = _p(file_path_P_L)
    file_path_P_L_1                  = _p(file_path_P_L_1)
    file_path_OSP                    = _p(file_path_OSP)
    file_path_HMF                    = _p(file_path_HMF)
    file_path_New_KPI                = _p(file_path_New_KPI)
    file_HQ_Item                     = _p(file_HQ_Item)
    file_goodwill                    = _p(file_goodwill)
    file_paths_High_Level_Variance   = _p(file_paths_High_Level_Variance)
    file_paths_Relief_KD             = _p(file_paths_Relief_KD)
    file_paths_Purchase_Price        = _p(file_paths_Purchase_Price)
    file_paths_UP_Calc_Template_Simulation = _p(file_paths_UP_Calc_Template_Simulation)
    file_paths_NANO_HFM              = _p(file_paths_NANO_HFM)
    file_paths_ITP                   = _p(file_paths_ITP)
    file_path_Template               = _p(file_path_Template)
    file_path_Template_1             = _p(file_path_Template_1)
    file_path_Template_man_1         = _p(file_path_Template_man_1)
    file_path_goodwill_template      = _p(file_path_goodwill_template)
    result                           = _p(result)
    output_path                      = _p(output_path)

    df_P_L_ROP_NE      = _d(df_P_L_ROP_NE)
    df_P_L_HQ          = _d(df_P_L_HQ)
    df_OSP             = _d(df_OSP)
    df_HMF_Managerial  = _d(df_HMF_Managerial)
    df_P_L_New_KPI     = _d(df_P_L_New_KPI)
    df_hQ_items        = _d(df_hQ_items)
    df_r_and_d         = _d(df_r_and_d)
    df_MZ_conso        = _d(df_MZ_conso)
    df_P_L_Relief      = _d(df_P_L_Relief)
    df_P_L_PL          = _d(df_P_L_PL)
    df_P_L_Reclass     = _d(df_P_L_Reclass)
    df_P_L_MCCP        = _d(df_P_L_MCCP)
    df_P_L_OCS         = _d(df_P_L_OCS)
    df_New_KPI_CFMI    = _d(df_New_KPI_CFMI)
    df_New_goodwill    = _d(df_New_goodwill)
    df_Act             = _d(df_Act)
    df_1c_osp          = _d(df_1c_osp)
    df_high_variance_feuil = _d(df_high_variance_feuil)
    df_Relief_KD       = _d(df_Relief_KD)
    df_Purchase_Price  = _d(df_Purchase_Price)
    df_UP_Calc_Template_Simulation = _d(df_UP_Calc_Template_Simulation)
    df_nano_hfm        = _d(df_nano_hfm)
    df_itp             = _d(df_itp)

    month_number = _int(month_number)

    # ── Layer 2: presence flags ───────────────────────────────────────────────
    HAS_RELIEF      = _df_ok(df_P_L_Relief,    "df_P_L_Relief")
    HAS_HMF         = _df_ok(df_HMF_Managerial,"df_HMF_Managerial")
    HAS_ACT         = _df_ok(df_Act,           "df_Act")
    HAS_NEW_KPI     = _df_ok(df_P_L_New_KPI,  "df_P_L_New_KPI")
    HAS_PL_PL       = _df_ok(df_P_L_PL,       "df_P_L_PL")
    HAS_OSP         = _df_ok(df_OSP,           "df_OSP")
    HAS_CFMI        = _df_ok(df_New_KPI_CFMI,  "df_New_KPI_CFMI")
    HAS_GOODWILL    = _df_ok(df_New_goodwill,  "df_New_goodwill")

    HAS_FILE_HMF         = _file_ok(file_path_HMF,          "HMF file")
    HAS_FILE_OSP         = _file_ok(file_path_OSP,          "OSP file")
    HAS_FILE_GOODWILL    = _file_ok(file_goodwill,           "Goodwill file")
    HAS_FILE_RESULT      = _file_ok(result,                  "result (P&L output) file")
    HAS_FILE_TMPL_1      = _file_ok(file_path_Template_1,   "Template_1 file")
    HAS_FILE_TMPL_MAN_1  = _file_ok(file_path_Template_man_1,"Template_man_1 file")
    HAS_FILE_GW_TMPL     = _file_ok(file_path_goodwill_template,"Goodwill template file")

    # ── Fiscal year label ─────────────────────────────────────────────────────
    current_date = datetime.now()
    fy_start = current_date.year if current_date.month >= 4 else current_date.year - 1
    current_year = f"{fy_start}-{str(fy_start + 1)[-2:]}"

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 1  –  COLUMN INDICES  (all int, 0 when df absent)
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 1: resolving column indices")

    col_relief    = _find_col_months(df_P_L_Relief,   month_number, "df_P_L_Relief")   if HAS_RELIEF  else 0
    col_hmf       = _find_col_months(df_HMF_Managerial, month_number, "df_HMF_Managerial") if HAS_HMF else 0
    act_col_adj   = _find_col_months(df_Act,          month_number, "df_Act")           if HAS_ACT     else 0
    col_new_kpi   = _find_col_months(df_P_L_New_KPI,  month_number, "df_P_L_New_KPI")  if HAS_NEW_KPI else 0
    col_pl        = _find_col_months(df_P_L_PL,       month_number, "df_P_L_PL")       if HAS_PL_PL   else 0
    col_cfmi      = _find_col_months(df_New_KPI_CFMI, month_number, "df_New_KPI_CFMI") if HAS_CFMI    else 0

    # write column for Act sheet (1-based for openpyxl)
    col_act_write = _int(act_col_adj) + 1

    # goodwill column offset
    adj                = (month_number - 4) % 12
    col_adj_goodwill   = _int(2 + adj)
    col_gw_write       = col_adj_goodwill + 1   # 1-based for openpyxl

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 2  –  OSP ROW / COL
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 2: OSP lookups")

    if HAS_OSP and HAS_FILE_OSP:
        try:
            row_OSP,   col_OSP   = find_itp_and_month_value(df_OSP, month_number)
            row_OSP,   col_OSP   = _int(row_OSP),   _int(col_OSP)
        except Exception as e:
            logger.warning(f"find_itp_and_month_value: {e}"); row_OSP = col_OSP = 0
        try:
            row_OSP_1, col_OSP_1 = find_second_column_for_outside_profit(df_OSP, month_number)
            row_OSP_1, col_OSP_1 = _int(row_OSP_1), _int(col_OSP_1)
        except Exception as e:
            logger.warning(f"find_second_column_for_outside_profit: {e}"); row_OSP_1 = col_OSP_1 = 0
    else:
        row_OSP = col_OSP = row_OSP_1 = col_OSP_1 = 0

    sheet_osp = "Master"
    formula_act_osp   = _safe_gen_formula_1(row_OSP,   col_OSP,   sheet_osp, file_path_OSP)
    formula_act_osp_1 = _safe_gen_formula_1(row_OSP_1, col_OSP_1, sheet_osp, file_path_OSP)
    formula_act_osp_1 = f"=-{formula_act_osp_1[1:]}" if formula_act_osp_1 != Z else Z
    formula_act_osp_sum = (
        f"=-(IFERROR({formula_act_osp[1:]},0))" if formula_act_osp != Z else Z
    )

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 3  –  HFM Managerial re-read (original code re-reads the file)
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 3: HFM Managerial re-read")

    if HAS_FILE_HMF:
        try:
            find_values_pandas_HFM(df_HMF_Managerial)
        except Exception as e:
            logger.warning(f"find_values_pandas_HFM: {e}")
        try:
            df_HMF_Managerial = pd.read_excel(file_path_HMF, sheet_name='Managerial')
            logger.info("HMF Managerial re-read OK")
        except Exception as e:
            logger.warning(f"pd.read_excel(HMF): {e} – continuing with existing df")
    else:
        logger.warning("[SKIP] HMF file absent – skipping re-read")

    # ── SearchResults attributes (may be 0 if search failed) ─────────────────
    # Wrap every attribute access in _int() so arithmetic is always safe
    def _sr(attr: str) -> int:
        try:
            return _int(getattr(SearchResults, attr, 0))
        except Exception:
            return 0

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 4  –  HFM formula block
    #  All gated on HAS_HMF + HAS_FILE_HMF + col_hmf > 0
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 4: building HFM formula block")

    sheet_man = "Managerial"

    def _fhmf(row_attr: str) -> str:
        if not HAS_HMF or not HAS_FILE_HMF or not col_hmf:
            return Z
        return _safe_gen_formula(_sr(row_attr), col_hmf, sheet_man, file_path_HMF)

    sheet_new  = "New KPI Actual"
    sheet_pl   = "P&L Actual"
    sheet_act  = "Act"

    formula_act_1  = _safe_gen_formula_1(34, col_new_kpi, sheet_new, result) if (HAS_FILE_RESULT and col_new_kpi) else Z
    formula_act_2  = _fhmf("wholesale_volume")
    formula_act_3  = _fhmf("gross_sales_wo_rd")
    formula_act_4  = _fhmf("rd_revenues")
    formula_act_5  = _fhmf("ci")
    formula_act_6  = _fhmf("mfg_fixed_plants")
    formula_act_7  = _fhmf("vendor_tooling")
    formula_act_8  = _fhmf("rd_costs")
    formula_act_9  = _fhmf("mfg_fixed_out_of_scope")
    formula_act_10 = _fhmf("fmi")
    formula_act_11 = _fhmf("ga_and_other_items")

    sga_row = _sr("sga_total")
    formula_act_12 = (
        _safe_gen_formula(sga_row + 1, col_hmf, sheet_man, file_path_HMF)
        if (HAS_HMF and HAS_FILE_HMF and col_hmf and sga_row) else Z
    )
    formula_act_13 = _fhmf("warranty_rs")

    # compound HFM formulas
    formula_sum_act_1 = _join(formula_act_3, f"+{formula_act_4[1:]}") if (formula_act_3 != Z and formula_act_4 != Z) else Z

    Formula_ex_act_1   = _safe_extract(file_path_Template_man_1, sheet_act, 30, act_col_adj, "ex_act_1")
    formula_sum_act_2  = _compound(formula_act_1, "-", Formula_ex_act_1)

    Formula_ex_act_d17 = _safe_extract(file_path_Template_man_1, sheet_act, 16, act_col_adj, "ex_act_d17")
    formula_sum_act_3  = _compound(formula_act_13, "-", Formula_ex_act_d17)

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 5  –  GOODWILL ROWS
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 5: goodwill row extraction")

    month_name           = calendar.month_name[month_number]
    file_path_goodwill_gen = ""
    if output_path:
        file_path_goodwill_gen = os.path.join(
            output_path, f"Goodwill_{month_name}_output_FY{current_year}.xlsx"
        )
    HAS_FILE_GW_GEN = _file_ok(file_path_goodwill_gen, "Goodwill generated output")

    formula_sum_act_10 = (
        _safe_gen_formula(21, col_adj_goodwill, "GW Budget", file_path_goodwill_gen)
        if HAS_FILE_GW_GEN else Z
    )

    # goodwill raw rows
    search_terms = [
        "goodwill", "ghq (cfmi)", "nmgb buyback - scope change to hq",
        "ger tcs", "act", "mobility",
    ]

    goodwill_rows = ghq_cfmi_rows = nmgb_buyback_rows = []
    ger_tcs_rows  = act_rows = mobility_rows = []
    column_indices = []

    if HAS_GOODWILL and HAS_FILE_GOODWILL:
        try:
            (column_indices, goodwill_rows, ghq_cfmi_rows, nmgb_buyback_rows,
             ger_tcs_rows, act_rows, mobility_rows) = goodwill_raw_data_extraction(
                df_New_goodwill, month_number, search_terms
            )
            # Ensure all row lists contain ints
            goodwill_rows    = _int_list(goodwill_rows)
            ghq_cfmi_rows    = _int_list(ghq_cfmi_rows)
            nmgb_buyback_rows= _int_list(nmgb_buyback_rows)
            ger_tcs_rows     = _int_list(ger_tcs_rows)
            act_rows         = _int_list(act_rows)
            mobility_rows    = _int_list(mobility_rows)
        except Exception as e:
            logger.warning(f"goodwill_raw_data_extraction: {e} – all GW rows → 0")
    else:
        logger.warning("[SKIP] Goodwill DF or file absent – GW formulas → =0")

    sheet_gew = "GW Budget"

    def _gw_f(row_list, label: str) -> str:
        """Build goodwill cross-wb formula, guarding on empty list and file."""
        r = _gw_row(row_list, label)   # int, 0 on failure
        if not r or not col_adj_goodwill or not HAS_FILE_GOODWILL:
            return Z
        return _safe_gen_formula_1(r, col_adj_goodwill, sheet_gew, file_goodwill)

    f1 = _gw_f(goodwill_rows,     "goodwill")
    f2 = _gw_f(ghq_cfmi_rows,     "ghq_cfmi")
    f3 = _gw_f(nmgb_buyback_rows, "nmgb_buyback")
    f4 = _gw_f(ger_tcs_rows,      "ger_tcs")
    f5 = _gw_f(act_rows,          "act")
    f6 = _gw_f(mobility_rows,     "mobility")

    logger.info(f"Goodwill f1={f1}")

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 6  –  RELIEF ROWS  (all int via _row_from)
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 6: relief row extraction")

    if HAS_RELIEF:
        try:
            rel_dict = relief_search_terms_new(df_P_L_Relief) or {}
        except Exception as e:
            logger.warning(f"relief_search_terms_new: {e}"); rel_dict = {}
    else:
        rel_dict = {}

    def _rrow(k): return _row_from(rel_dict, k)

    row_recall          = _rrow("recall service/campaigns")
    row_kd_nmuk         = _rrow("kd itp (nmuk)")
    row_kd_nmisa        = _rrow("kd itp (nmisa)")
    row_kd_nmgr         = _rrow("kd itp (nmgr)")
    row_connectivity    = _rrow("connectivity")
    row_itp_nissan      = _rrow("itp cbu nissan")
    row_itp_infiniti    = _rrow("itp cbu infiniti")
    row_iln_nmuk        = _rrow("iln export (nmuk)")
    row_iln_nmisa       = _rrow("iln export (nmisa)")
    row_err_mz          = _rrow("error correction in mz (?)")
    row_export          = _rrow("export")
    row_mz_imp          = _rrow("mz impairment")
    row_ga_imp          = _rrow("g&a impairment")    # used in +1 arithmetic below
    row_cev             = _rrow("compact ev/vt impairment")
    row_bat_nmuk        = _rrow("battery 40 kw itp ( nmuk )")
    row_bat_nmisa       = _rrow("battery 40 kw itp ( nmisa )")
    row_bat_cancel      = _rrow("battery cancelation")
    row_err_wd5         = _rrow("error correction in mz (after wd5)")
    row_top_daim        = _rrow("top daimler correction (after wd5)")
    row_inf_fmi         = _rrow("infiniti fmi correction (after wd5)")

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 7  –  CFMI ROW / COL
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 7: CFMI row extraction")

    if HAS_CFMI:
        try:
            cfmi_rows = find_row_number_CMFI(df_New_KPI_CFMI)
            cfmi_rows = _int_list(cfmi_rows or [0])
        except Exception as e:
            logger.warning(f"find_row_number_CMFI: {e}"); cfmi_rows = [0]
    else:
        cfmi_rows = [0]

    cfmi_row = cfmi_rows[0] if cfmi_rows else 0

    formula_act_cfmi = (
        _safe_gen_formula(cfmi_row, col_cfmi, "CFMI")
        if (cfmi_row and col_cfmi) else Z
    )

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 8  –  FORMULA BUILDING
    #  Every formula builder is gated on HAS_* + col > 0 + row > 0.
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 8: building formulas")

    sheet_rel = "Relief"
    sheet_new = "New KPI Actual"
    sheet_pl  = "P&L Actual"
    sheet_act = "Act"

    # ── helpers bound to their source ─────────────────────────────────────────
    def _frel(row: int) -> str:
        row = _int(row)
        if not HAS_RELIEF or not row or not col_relief:
            return Z
        return _safe_gen_formula(row, col_relief, sheet_rel, result)

    def _fnew(row: int) -> str:
        row = _int(row)
        if not HAS_FILE_RESULT or not row or not col_new_kpi:
            return Z
        return _safe_gen_formula_1(row, col_new_kpi, sheet_new, result)

    def _fpl(row: int) -> str:
        row = _int(row)
        if not HAS_FILE_RESULT or not row or not col_pl:
            return Z
        return _safe_gen_formula_1(row, col_pl, sheet_pl, result)

    def _ex(row: int, col: int, label: str = "") -> str:
        return _safe_extract(file_path_Template_man_1, sheet_act,
                             _int(row), _int(col), label)

    # ── relief formulas ───────────────────────────────────────────────────────
    formula_relief_1 = _frel(row_recall)

    # ── New KPI formulas ──────────────────────────────────────────────────────
    formula_act_new_2  = _fnew(282)
    Formula_ex_new_2   = _ex(57,  act_col_adj, "ex_new_2")
    formula_newkpi_sum_2 = _compound(formula_act_new_2, "-", Formula_ex_new_2)

    formula_act_new_3  = _fnew(251)
    Formula_ex_new_4   = _ex(83,  act_col_adj, "ex_new_4")
    formula_newkpi_sum_4 = _compound(formula_act_new_3, "-", Formula_ex_new_4)

    formula_act_new_5  = _fnew(313)
    Formula_ex_new_41  = _ex(138, act_col_adj, "ex_new_41")
    formula_newkpi_sum_41 = _compound(formula_act_new_5, "-", Formula_ex_new_41)

    formula_act_new_6  = _fnew(335)
    formula_act_new_7  = _fnew(45)
    Formula_ex_new_16  = _ex(214, act_col_adj, "ex_new_16")
    formula_pl_sum_20  = (
        f"={Formula_ex_new_16[1:]}+{formula_act_new_7[1:]}"
        if (Formula_ex_new_16 != Z and formula_act_new_7 != Z) else Z
    )

    formula_act_new_8  = _fnew(437)
    formula_act_new_9  = _fnew(344)
    formula_act_new_10 = _fnew(44)
    formula_act_new_11 = _fnew(52)
    Formula_ex_new_11  = _ex(219, act_col_adj, "ex_new_11")
    # sum: 8+9+10+11 - ex
    _sum_parts = [x[1:] for x in [formula_act_new_8, formula_act_new_9,
                                    formula_act_new_10, formula_act_new_11] if x != Z]
    formula_newkpi_sum_11 = (
        f"={'  +  '.join(_sum_parts)}-{Formula_ex_new_11[1:]}"
        if (_sum_parts and Formula_ex_new_11 != Z) else Z
    )

    # ── P&L formulas ──────────────────────────────────────────────────────────
    formula_act_pl_1  = _fpl(307)
    formula_act_pl_2  = _fpl(309)
    formula_pl_sum_1  = (
        f"=-({formula_act_pl_1[1:]}+{formula_act_pl_2[1:]})"
        if (formula_act_pl_1 != Z and formula_act_pl_2 != Z) else Z
    )

    formula_act_pl_3  = _fpl(305)
    formula_act_pl_4  = _fpl(336)
    formula_act_pl_19 = _fpl(310)
    formula_act_pl_22 = _fpl(308)
    formula_act_pl_24 = _fpl(9)
    formula_act_pl_25 = _fpl(13)
    formula_act_pl_26 = _fpl(166)
    formula_act_pl_27 = _fpl(162)
    formula_act_pl_29 = _fpl(164)
    formula_act_pl_30 = _fpl(161)

    formula_act_new_28 = _fnew(319)
    formula_pl_sum_28 = (
        f"={formula_act_pl_27[1:]}+{formula_act_new_28[1:]}"
        if (formula_act_pl_27 != Z and formula_act_new_28 != Z) else Z
    )
    formula_pl_sum_10 = (
        f"={formula_act_pl_22[1:]}+{formula_act_pl_19[1:]}"
        if (formula_act_pl_22 != Z and formula_act_pl_19 != Z) else Z
    )

    # ── Relief compound formulas  (all row_* are int; arithmetic is safe) ─────
    # iln export
    f_iln_nmuk  = _frel(row_iln_nmuk)
    f_iln_nmisa = _frel(row_iln_nmisa)
    # ga_imp + 1: row_ga_imp is int from _row_from → safe
    f_ga_imp_p1 = _frel(_int(row_ga_imp) + 1) if row_ga_imp else Z

    formula_act_rel_1 = f_iln_nmuk
    formula_act_rel_2 = f_iln_nmisa
    formula_act_rel_3 = f_ga_imp_p1
    formula_rel_sum_1 = (
        f"=-({'+'.join(x[1:] for x in [formula_act_rel_1, formula_act_rel_2, formula_act_rel_3] if x != Z)})"
        if any(x != Z for x in [formula_act_rel_1, formula_act_rel_2, formula_act_rel_3]) else Z
    )

    formula_act_rel_4 = _frel(row_kd_nmuk)
    formula_act_rel_5 = _frel(row_kd_nmisa)
    formula_act_rel_6 = _frel(row_kd_nmgr)
    formula_act_rel_7 = _frel(row_itp_nissan)
    formula_act_rel_8 = _frel(row_itp_infiniti)
    # itp_cbu_nissan - 1: row_itp_nissan is int → safe
    formula_act_rel_9 = _frel(_int(row_itp_nissan) - 1) if row_itp_nissan else Z

    _r2_parts = [x[1:] for x in [formula_act_rel_4, formula_act_rel_5,
                                   formula_act_rel_6, formula_act_rel_7,
                                   formula_act_rel_8] if x != Z]
    formula_rel_sum_2 = f"=-({'+'.join(_r2_parts)})" if _r2_parts else Z

    Formula_ex_new_rel_1 = _ex(165, act_col_adj, "ex_rel_1")
    _r3_parts = [x[1:] for x in [
        formula_act_rel_1, formula_act_rel_2, formula_act_rel_4,
        formula_act_rel_5, formula_act_rel_6, formula_act_rel_7,
        formula_act_rel_8, formula_act_rel_9,
    ] if x != Z]
    formula_rel_sum_3 = (
        f"=-({'+'.join(_r3_parts)})-{Formula_ex_new_rel_1[1:]}"
        if (_r3_parts and Formula_ex_new_rel_1 != Z) else Z
    )

    formula_act_new_12 = _fnew(211)
    formula_act_new_13 = _fnew(220)
    formula_act_new_21 = _fnew(468)
    Formula_ex_new_rel_12 = _ex(246, act_col_adj, "ex_rel_12")
    formula_rel_sum_11 = _compound(formula_act_new_21, "-", Formula_ex_new_rel_12)

    formula_act_rel_10 = _frel(row_connectivity)
    formula_rel_sum_4  = _compound(formula_act_new_12, "-", formula_act_rel_10)

    Formula_ex_new_rel_2 = _ex(192, act_col_adj, "ex_rel_2")
    formula_rel_sum_5 = (
        f"=-{formula_act_rel_10[1:]}+{formula_act_new_13[1:]}-{Formula_ex_new_rel_2[1:]}"
        if all(x != Z for x in [formula_act_rel_10, formula_act_new_13, Formula_ex_new_rel_2])
        else Z
    )

    # ── f7 / f8: goodwill P&L and template cross-ref ──────────────────────────
    f7 = _fpl(102)
    # f8 uses col_act_write (1-based) for reading the template
    f8 = (
        _safe_gen_formula_1(23, col_act_write, sheet_act, file_path_Template_1)
        if HAS_FILE_TMPL_1 and col_act_write else Z
    )

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 9  –  WRITE KPI OUTPUT FILE
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 9: writing KPI output file")

    sheet_names = ['CFMI', 'Act']

    cell_updates = [
        # CFMI sheet
        [
            (4, col_cfmi + 1, var_Cfmi),
        ],
        # Act sheet  (col_act_write = act_col_adj + 1, 1-based)
        [
            (5,   col_act_write, formula_act_2),
            (6,   col_act_write, formula_sum_act_1),
            (7,   col_act_write, formula_act_5),
            (12,  col_act_write, formula_act_6),
            (13,  col_act_write, formula_act_7),
            (14,  col_act_write, formula_act_8),
            (16,  col_act_write, formula_sum_act_3),
            (17,  col_act_write, formula_relief_1),
            (18,  col_act_write, formula_act_9),
            (23,  col_act_write, formula_act_10),
            (26,  col_act_write, formula_act_11),
            (29,  col_act_write, formula_act_12),
            (30,  col_act_write, formula_sum_act_2),
            (33,  col_act_write, formula_act_osp_sum),
            (37,  col_act_write, formula_act_osp_1),
            (57,  col_act_write, formula_newkpi_sum_2),
            (64,  col_act_write, formula_act_new_3),
            (84,  col_act_write, formula_newkpi_sum_4),
            (114, col_act_write, formula_pl_sum_1),
            (126, col_act_write, formula_act_pl_3),
            (138, col_act_write, formula_newkpi_sum_41),
            (141, col_act_write, formula_rel_sum_1),
            (144, col_act_write, formula_rel_sum_2),
            (165, col_act_write, formula_rel_sum_3),
            (171, col_act_write, formula_rel_sum_4),
            (192, col_act_write, formula_rel_sum_5),
            (198, col_act_write, formula_act_new_6),
            (209, col_act_write, formula_sum_act_10),
            (210, col_act_write, formula_act_cfmi),
            (213, col_act_write, formula_act_pl_4),
            (214, col_act_write, formula_pl_sum_20),
            (219, col_act_write, formula_newkpi_sum_11),
            (246, col_act_write, formula_rel_sum_11),
            (253, col_act_write, formula_act_new_7),
            (260, col_act_write, formula_pl_sum_10),
            (267, col_act_write, formula_act_pl_30),
            (268, col_act_write, formula_act_pl_29),
            (270, col_act_write, formula_pl_sum_28),
            (283, col_act_write, formula_act_pl_26),
            (292, col_act_write, formula_act_pl_24),
            (293, col_act_write, formula_act_pl_25),
            (295, col_act_write, Var_Converter),
        ],
    ]

    file_path_new_KPI_latest = Z   # sentinel; overwritten on success
    if HAS_FILE_TMPL_1:
        try:
            file_path_new_KPI_latest = paste_values_KPI_PL(
                file_path_Template_1, sheet_names, cell_updates,
                month_number, output_path,
            )
            logger.info(f"ACT KPI output written → {file_path_new_KPI_latest}")
        except Exception as e:
            logger.warning(f"paste_values_KPI_PL (ACT): {e}")
    else:
        logger.warning("[SKIP] Template_1 absent – KPI output not written")

    # ── f8 updated with actual output path ────────────────────────────────────
    if _file_ok(str(file_path_new_KPI_latest), "KPI latest output"):
        f8 = _safe_gen_formula_1(
            23, _int(col_act_write) - 1, sheet_act, str(file_path_new_KPI_latest)
        )
    else:
        logger.warning("[SKIP] f8 update skipped – KPI latest output not available")

    # ══════════════════════════════════════════════════════════════════════════
    #  STEP 10  –  WRITE GOODWILL OUTPUT FILE
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("ACT Step 10: writing Goodwill output file")

    sheet_names_goodwill = ["GW Budget"]

    cell_updates_goodwill = [
        [
            (5,  col_gw_write, f1),
            (6,  col_gw_write, f2),
            (9,  col_gw_write, f3),
            (10, col_gw_write, f4),
            (14, col_gw_write, f5),
            (15, col_gw_write, f6),
            (20, col_gw_write, f7),
            (21, col_gw_write, f8),
            (24, col_gw_write, var_Cfmi),
        ],
    ]

    if HAS_FILE_GW_TMPL:
        try:
            paste_values_goodwill(
                file_path_goodwill_template, sheet_names_goodwill,
                cell_updates_goodwill, month_number, output_path,
            )
            logger.info("ACT Goodwill output written")
        except Exception as e:
            logger.warning(f"paste_values_goodwill: {e}")
    else:
        logger.warning("[SKIP] Goodwill template absent – GW output not written")

    logger.info("ACT completed")

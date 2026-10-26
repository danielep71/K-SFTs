"""Build the macro-free K-SFTs workbook template.

Writes src/workbook/SFT_Template.xlsx from the fixtures, in the layout that
src/core/CORE_Config.bas describes: every table in column C, located through
the workbook name of its first header cell; every parameter read through the
workbook name equal to its code. The template holds no VBA: save it as .xlsm
and import the modules of src/ (see INSTALLATION.md). The sheet buttons call
RunSFT, ValidateInputs and ClearOutputs by name.

Usage: python tools/build_template.py
"""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Any

import xlsxwriter  # type: ignore[import-untyped,import-not-found,unused-ignore]
from xlsxwriter.utility import (  # type: ignore[import-untyped,import-not-found,unused-ignore]
    xl_col_to_name,
    xl_rowcol_to_cell,
)

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "src/workbook/SFT_Template.xlsx"

FIXTURE = json.loads((ROOT / "tests/derivations/prototype/demo-portfolio.json").read_text())
TABLES = json.loads((ROOT / "tests/derivations/prototype/regulatory-tables.json").read_text())
EXPECTED = json.loads((ROOT / "tests/derivations/prototype/demo-portfolio.expected.json").read_text())["results"]

NAVY, BLUE, BAND, GOLD, INPUT, GREY = "#1F3864", "#2E75B6", "#265A99", "#FFC000", "#FFF2CC", "#BFBFBF"
TAB_DOC, TAB_IN, TAB_OUT, TAB_TEST = "#1F3864", "#BF9000", "#2E75B6", "#548235"
FONT = "Aptos"

wb = xlsxwriter.Workbook(str(OUT))
wb.set_properties({"title": "", "author": "", "company": "", "comments": ""})
wb.formats[0].set_font_name(FONT)
wb.formats[0].set_font_size(10)

_cache: dict[tuple[tuple[str, Any], ...], Any] = {}


def fmt(**kw: Any) -> Any:
    key = tuple(sorted(kw.items()))
    if key not in _cache:
        base = {"font_name": FONT, "valign": "vcenter"}
        base.update(kw)
        _cache[key] = wb.add_format(base)
    return _cache[key]


F: dict[str, Any] = {
    "brand": fmt(font_size=14, font_color=GOLD, bg_color="#000000"),
    "band1": fmt(bg_color="#000000"),
    "title": fmt(font_size=18, bold=True, font_color="#FFFFFF", bg_color=BAND),
    "band2": fmt(bg_color=BAND),
    "note": fmt(font_size=10, italic=True, font_color="#595959"),
    "note_wrap": fmt(font_size=9, font_color="#595959", text_wrap=True, valign="top"),
    "runinfo": fmt(font_size=10, italic=True, font_color="#7F6000"),
    "section": fmt(font_size=14, bold=True, font_color=NAVY, bottom=2, bottom_color=NAVY),
    "hdr_in": fmt(font_size=9, bold=True, font_color="#FFFFFF", bg_color=NAVY, text_wrap=True, border=1,
                  border_color=GREY),
    "hdr_prm": fmt(font_size=11, bold=True, font_color="#FFFFFF", bg_color=NAVY, text_wrap=True, border=1,
                   border_color=GREY),
    "hdr_out": fmt(font_size=10, bold=True, font_color="#FFFFFF", bg_color=BLUE, text_wrap=True, border=1,
                   border_color=GREY, align="center"),
    "hdr_out_l": fmt(font_size=10, bold=True, font_color="#FFFFFF", bg_color=BLUE, text_wrap=True, border=1,
                     border_color=GREY),
    "code": fmt(font_size=10, bold=True),
    "desc": fmt(font_size=9),
    "ref": fmt(font_size=9, font_color="#595959"),
    "ref_wrap": fmt(font_size=9, font_color="#595959", text_wrap=True, valign="top"),
    "in_txt": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="@"),
    "in_ctr": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="@", align="center"),
    "in_amt": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="#,##0"),
    "in_num": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="0.####"),
    "in_pct": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="0.00%", align="center"),
    "in_pct3": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="0.000%", align="center"),
    "in_rate": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="0.000000"),
    "in_date": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="yyyy-mm-dd", align="center"),
    "in_bool": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, align="center"),
    "out_txt": fmt(font_size=10, num_format="@"),
    "rd_h": fmt(font_size=14, bold=True, font_color=NAVY, bottom=2, bottom_color=NAVY, text_wrap=True),
    "rd_k": fmt(font_size=10, bold=True, text_wrap=True, bottom=3, bottom_color=GREY),
    "rd_v": fmt(font_size=9, text_wrap=True, bottom=3, bottom_color=GREY),
    "rd_link": fmt(font_size=9, text_wrap=True, bottom=3, bottom_color=GREY, font_color="#0563C1", underline=1),
    "cf_bad": wb.add_format({"bg_color": "#F8CBAD"}),
    "cf_warn": wb.add_format({"bg_color": "#FFE699"}),
    "cf_pass": wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"}),
    "cf_fail": wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"}),
}


def banner(ws: Any, title: str, last_col: int) -> None:
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 4)
    ws.set_column(1, 1, 2.5)
    ws.set_row(0, 24)
    ws.set_row(1, 26.1)
    for c in range(1, last_col + 1):
        ws.write_blank(0, c, None, F["band1"])
        ws.write_blank(1, c, None, F["band2"])
    ws.write(0, 1, "K-SFTs", F["brand"])
    ws.write(1, 1, title, F["title"])


def name(nm: str, ws_title: str, row: int, col: int, absolute: bool = True) -> None:
    wb.define_name(nm, f"='{ws_title}'!" + xl_rowcol_to_cell(row, col, row_abs=absolute, col_abs=absolute))


def buttons(ws: Any, row: int, col: int) -> None:
    for i, (caption, macro) in enumerate([("Run SFT", "RunSFT"), ("Validate inputs", "ValidateInputs"),
                                          ("Clear outputs", "ClearOutputs")]):
        ws.insert_button(row, col, {"macro": macro, "caption": caption, "width": 112, "height": 24,
                                    "x_offset": 4 + i * 122, "y_offset": 6})


def date(text: str) -> dt.datetime | None:
    return dt.datetime.fromisoformat(text) if text else None


# ----------------------------------------------------------------------------
# README
# ----------------------------------------------------------------------------
readme = wb.add_worksheet("README")
readme.set_tab_color(TAB_DOC)
banner(readme, "ReadMe", 4)
readme.set_column(2, 2, 46)
readme.set_column(3, 3, 140)
readme.set_column(4, 4, 4)
readme.write(2, 2, "Exposure value E*, risk-weighted exposure amount and own funds requirement of securities "
             "financing transactions under the financial collateral comprehensive method with supervisory "
             "haircuts: EU CRR (Art. 220-227), the default, and Basel CRE22, chosen per netting set.", F["note"])
readme.set_row(2, 34)
README = [
    ("How to use", None),
    ("1. Enable macros", "This file is the macro-free template. Save it as an Excel Macro-Enabled Workbook "
     "(.xlsm), import the modules of src/core, src/modules and src/workbook (INSTALLATION.md), then open "
     "it and click \"Enable Content\". The SFT_* cells on FormulaTest show #NAME? until the modules are "
     "imported; press Ctrl+Alt+F9 afterwards."),
    ("2. Parameters", "Params: reporting date, reporting currency, default Regime (CRR or BCBS), holding "
     "periods, haircut weights, the supervisory haircut table (H10, 10 business days, daily remargining), "
     "the risk-weight table and FX rates (units of reporting currency per 1 unit). Yellow cells are inputs."),
    ("3. Netting sets", "NettingSets: one row per netting set - whether an eligible master netting agreement "
     "covers it, its settlement currency, the counterparty's exposure class, CQS or SCRA grade, the "
     "remargining frequency, the 20-day and dispute drivers of the holding period, the 0% haircut "
     "declaration and optional risk-weight and regime overrides."),
    ("4. Transactions", "Transactions: one row per leg. A leg is cash or one security issue the bank lent, "
     "sold or posted (LENT, E) or borrowed, bought or received (RECEIVED, C). Legs sharing a TransactionID "
     "form one transaction and repeat its netting set, type, start and maturity dates. Amounts can be in "
     "any currency of the FX table."),
    ("5. Run", "Click \"Run SFT\" (macro RunSFT). Results, Positions, LegCalc and Checks are rebuilt. "
     "\"Validate inputs\" runs the checks only and clears the other output sheets; \"Clear outputs\" "
     "empties the output sheets."),
    ("6. Review", "Checks lists every error (transaction rejected, netting set withheld, run stopped), "
     "warning and information message, including each difference from the regulatory parameter set "
     "(WHAT-IF). TestCatalogue compares Results with independently derived values; FormulaTest unit-tests "
     "the SFT_* functions."),
    ("7. Regime", "Basel_vs_CRR lists the differences between CRE22 / CRE20 and the CRR and how the "
     "workbook treats them. Regime-dependent rows of the Params tables are switched per netting set."),
    ("Sign and input conventions", None),
    ("Side", "LENT: cash or securities lent, sold under repo or posted as collateral (E). RECEIVED: cash or "
     "securities borrowed, bought under reverse repo or received as collateral (C). A reverse repo is "
     "cash LENT and securities RECEIVED; a repo is securities LENT and cash RECEIVED."),
    ("Transaction type", "REPO, REVREPO, SECLEND, SECBORROW: repo-style transactions, minimum holding period "
     "5 business days. MARGINLEND: other capital-market-driven transaction, 10 business days. A netting set "
     "holding both uses 10 [CRE22.57-22.58; CRR Art. 224(2)]."),
    ("Asset type", "CASH; DEBT with issuer SOV / INST / CORP / SEC, issue CQS (blank = unrated) and security "
     "maturity date; EQUITY with listing MAIN (main index), LISTED (recognised exchange) or OTHER; GOLD; "
     "FUND with its look-through haircut H10 (or the highest haircut of any asset the fund may invest in); "
     "OTHER. Every non-cash leg names its Issue ID: legs of one issue are netted under a master netting "
     "agreement."),
    ("Eligibility", "A leg whose haircut key is not in the haircut table is not eligible: lent, it takes "
     "NonEligibleHaircut (30%) [CRE22.54; Art. 224(4)]; received, it is not recognised as collateral. "
     "Unrated debt of an institution is treated as CQS 2-3 [CRE22.39(d); Art. 197(4), 224(6)]."),
    ("Methodology (per netting set)", None),
    ("Netting (eligible MNA)", "E* = max(0, sum E - sum C + 0.4 x net + 0.6 x gross / sqrt(N) + sum |E_fx| x "
     "Hfx); net = |sum E_s H_s|, gross = sum |E_s| H_s over security issues s, N = issues with |E_s| at "
     "least one tenth of the largest [CRE22.64-22.65; CRR Art. 220(3)]"),
    ("Single transaction (no MNA)", "E* = max(0, E x (1 + He) - C x (1 - Hc - Hfx)) per transaction, summed "
     "over the netting set; baskets use value-weighted haircuts [CRE22.46, 22.55; CRR Art. 223]"),
    ("Haircuts", "H = H10 x sqrt((N_R + T_M - 1) / 10); T_M = holding period, N_R = business days between "
     "remargining [CRE22.59; CRR Art. 226]. Currency mismatch Hfx = 8% scaled the same way [CRE22.52; "
     "Art. 224 Table 4]."),
    ("Holding period", "5 BD repo-style, 10 BD margin lending; 20 BD for a netting set with more than 5,000 "
     "transactions in the previous quarter or illiquid collateral; doubled after margin disputes; an "
     "override can only lengthen it [CRE22.57-22.58; CRR Art. 224(2), 285(3)-(4)]."),
    ("0% haircut", "Declared per netting set; the engine checks the conditions it can see (repo-style only, "
     "cash and central government securities, one currency, daily remargining) and takes the rest as "
     "declared [CRE22.60; CRR Art. 227]."),
    ("Capital", "RWA = E* x RW of the counterparty; own funds requirement = 8% x RWA [CRE22.48, RBC20.1; CRR "
     "Art. 92(1)(c)]. Institutions: ECRA by CQS, SCRA grade A / A_ENH / B / C when unrated, short-term "
     "weights when every transaction has an original maturity of 3 months or less [CRE20.18-20.21; Art. "
     "120-121]. QCCP trade exposures 2% [CRE54.7; Art. 306(1)(a)]."),
    ("Sheets", None),
    ("Params", "Inputs - parameters, haircut table, risk-weight table, FX rates"),
    ("NettingSets", "Inputs - netting sets, counterparties and holding-period drivers"),
    ("Transactions", "Inputs - transaction legs"),
    ("Results", "Output - E*, RWA and own funds per netting set, and TOTAL CRR / TOTAL BCBS over the VALID "
     "netting sets of each regime"),
    ("Positions", "Output - per netting set with netting: each security issue (E_s, H_s, counted in N) and "
     "each currency (FX term); without netting: each transaction"),
    ("LegCalc", "Output - every leg: haircut key, eligibility, residual maturity, H10, holding period, scale, "
     "H, Hfx, adjusted value"),
    ("Checks", "Output - errors (transaction rejected, netting set withheld, run stopped), warnings and "
     "information"),
    ("TestCatalogue", "Regression - expected values from tests/derivations/sft_reference.py against Results"),
    ("FormulaTest", "Unit tests of the SFT_* worksheet functions"),
    ("Basel_vs_CRR", "Documentation - Basel CRE22 / CRE20 vs EU CRR and workbook treatment"),
    ("Worksheet functions (usable in any cell)", None),
    ("SFT_HoldingPeriod(MinimumBD, [LargeOrIlliquid], [Disputes], [OverrideBD], [LargeBD])", "Holding period T_M"),
    ("SFT_ScaledHaircut(H10, HoldingBD, [RemarginBD], [BaseBD])", "Haircut for T_M and N_R"),
    ("SFT_ExposureSingle(E, He, C, Hc, [Hfx])", "E* of one transaction"),
    ("SFT_NetExposure(Es, Hs)", "Net exposure |sum E_s H_s|"),
    ("SFT_GrossExposure(Es, Hs)", "Gross exposure sum |E_s| H_s"),
    ("SFT_IssueCount(Es, [Threshold])", "N: issues of at least one tenth of the largest"),
    ("SFT_ExposureNetted(SumE, SumC, Net, Gross, N, [FXTerm], [NetWeight], [GrossWeight])", "E* of a netting set"),
    ("SFT_CapitalRequirement(EAD, RiskWeight, [CapitalRatio])", "Own funds requirement; RWA with ratio 1"),
    ("Scope and limitations", None),
    ("Not covered", "SA-CCR and derivatives; internal models (IMM, VaR, own-estimate haircuts); "
     "the simple approach; secured lending and commodity repos; haircut floors for non-centrally cleared "
     "SFTs (CRE56, not in the CRR); maturity mismatch; credit risk mitigation on the exposure side other "
     "than collateral; cross-product netting with derivatives; CVA; the SCRA sovereign floor; capital "
     "buffers and Pillar 2."),
    ("Status", "Prototype. Implements the regulatory set CRE22-CRR-2026-10-24, checked against CRE22 "
     "(effective 1 January 2023) and the consolidated CRR of 26.06.2026. Expected values are reproduced by "
     "an independent Python derivation, not by published worked examples."),
    ("Sample data", "Illustrative counterparties, securities and FX rates. BCBS-MNA repeats CRR-MNA under "
     "Basel; DEMO-REJECTED shows how a rejected transaction withholds a netting set's E*."),
    ("Code", None),
    ("Modules", "M_Main (macros), CORE_Engine (calculation), SFT_Formulas (public SFT_* functions), "
     "CORE_Config (layout constants and the regulatory parameter set), CORE_Util (helpers). No external "
     "references: no Scripting.Dictionary, no WorksheetFunction."),
]
r = 4
for key, value in README:
    if value is None:
        readme.set_row(r, 30)
        readme.write(r, 2, key, F["rd_h"])
        readme.write_blank(r, 3, None, F["rd_h"])
    else:
        readme.set_row(r, 30 if len(value) < 170 else 44 if len(value) < 340 else 58)
        readme.write(r, 2, key, F["rd_k"])
        readme.write(r, 3, value, F["rd_v"])
    r += 1
readme.freeze_panes(3, 0)

# ----------------------------------------------------------------------------
# PARAMS
# ----------------------------------------------------------------------------
P = FIXTURE["params"]
PARAMS = [
    ("AsOfDate", "Reporting (calculation) date", date(P["AsOfDate"]), "in_date", ""),
    ("ReportingCcy", "Reporting currency (all amounts converted into it)", P["ReportingCcy"], "in_ctr", ""),
    ("Regime", "Default regime: CRR (EU) or BCBS (Basel CRE22); override per netting set", P["Regime"], "in_ctr",
     "see Basel_vs_CRR sheet"),
    ("DaysPerYear", "Calendar days per year for the residual maturity of debt securities", P["DaysPerYear"],
     "in_num", "convention"),
    ("HaircutBaseBD", "Holding period of the haircut table (BD)", P["HaircutBaseBD"], "in_num",
     "CRE22.49 / CRR Art. 224 Table 1 (10-day column)"),
    ("HoldingRepoBD", "Minimum holding period, repo-style transactions (BD)", P["HoldingRepoBD"], "in_num",
     "CRE22.57 / CRR Art. 224(2)"),
    ("HoldingCapitalMarketBD", "Minimum holding period, other capital-market-driven transactions (BD)",
     P["HoldingCapitalMarketBD"], "in_num", "CRE22.57 / CRR Art. 224(2)"),
    ("HoldingLargeBD", "Holding period, >5,000 transactions or illiquid collateral (BD)", P["HoldingLargeBD"],
     "in_num", "CRE22.58 / CRR Art. 224(2), 285(3)"),
    ("FXHaircut", "Currency-mismatch haircut Hfx, 10 BD", P["FXHaircut"], "in_pct", "CRE22.52 / CRR Art. 224 Table 4"),
    ("NonEligibleHaircut", "Haircut of non-eligible securities lent, 10 BD", P["NonEligibleHaircut"], "in_pct",
     "CRE22.54 / CRR Art. 224(4)"),
    ("NetWeight", "Weight of the net exposure (systematic term)", P["NetWeight"], "in_num",
     "CRE22.64-22.65 / CRR Art. 220(3)"),
    ("GrossWeight", "Weight of the gross exposure (idiosyncratic term)", P["GrossWeight"], "in_num",
     "CRE22.64-22.65 / CRR Art. 220(3)"),
    ("IssueCountThreshold", "Issues below this share of the largest |E_s| are not counted in N",
     P["IssueCountThreshold"], "in_num", "CRE22.65 / CRR Art. 220(3)"),
    ("CapitalRatio", "Own funds requirement per unit of RWA", P["CapitalRatio"], "in_pct",
     "RBC20.1 / CRR Art. 92(1)(c)"),
]
prm = wb.add_worksheet("Params")
prm.set_tab_color(TAB_IN)
banner(prm, "Parameters", 10)
prm.set_column(2, 2, 26)
prm.set_column(3, 3, 52)
prm.set_column(4, 4, 16)
prm.set_column(5, 9, 14)
prm.set_column(10, 10, 38)
prm.write(2, 2, "Codes in column C are read by the engine through the workbook name equal to each code; only the "
          "yellow values are meant to be changed.", F["note"])
prm.set_row(2, 30)
prm.set_row(4, 30)
for c, h in enumerate(["Code", "Description", "Value", "Reference"]):
    prm.write(4, 2 + c, h, F["hdr_prm"])
name("Params_Code", "Params", 4, 2)
name("Params_Value", "Params", 4, 4)
r = 5
for code, desc, value, style, ref in PARAMS:
    prm.set_row(r, 20.1)
    prm.write(r, 2, code, F["code"])
    prm.write(r, 3, desc, F["desc"])
    if isinstance(value, dt.datetime):
        prm.write_datetime(r, 4, value, F[style])
    else:
        prm.write(r, 4, value, F[style])
    prm.write(r, 5, ref, F["ref"])
    name(code, "Params", r, 4)
    r += 1
prm.data_validation("E6", {"validate": "date", "criteria": ">", "value": dt.datetime(2000, 1, 1)})
prm.data_validation("E8", {"validate": "list", "source": ["CRR", "BCBS"]})

r += 1
prm.set_row(r, 26)
prm.write(r, 2, "Supervisory haircuts H10 (10 business days, daily remargining) - CRE22.49 / CRR Art. 224",
          F["section"])
r += 1
prm.set_row(r, 30)
for c, h in enumerate(["HC_Key", "Asset type", "Issuer category", "Credit quality", "Residual maturity",
                       "Haircut H10", "Regimes", "Reference"]):
    prm.write(r, 2 + c, h, F["hdr_prm"])
name("HC_Table", "Params", r, 2)
r += 1
for key, asset, issuer, cq, band, h10, regimes, ref in TABLES["haircuts"]:
    prm.write(r, 2, key, F["code"])
    prm.write(r, 3, asset, F["desc"])
    prm.write(r, 4, issuer, F["desc"])
    prm.write(r, 5, cq, F["desc"])
    prm.write(r, 6, band, F["desc"])
    prm.write(r, 7, h10, F["in_pct3"])
    prm.write(r, 8, regimes, F["in_ctr"])
    prm.write(r, 9, ref, F["ref"])
    r += 1
prm.write(r, 3, "Keys not in this table are not eligible. Add rows directly below the last one (no blank rows).",
          F["ref"])

r += 2
prm.set_row(r, 26)
prm.write(r, 2, "Risk weights - CRE20 / CRR Art. 114, 120-122; QCCP CRE54.7 / Art. 306(1)(a)", F["section"])
r += 1
prm.set_row(r, 30)
for c, h in enumerate(["RW_Key", "Exposure class", "Grade", "Risk weight", "Regimes", "Reference"]):
    prm.write(r, 2 + c, h, F["hdr_prm"])
name("RW_Table", "Params", r, 2)
r += 1
for key, cls, grade, weight, regimes, ref in TABLES["risk_weights"]:
    prm.write(r, 2, key, F["code"])
    prm.write(r, 3, cls, F["desc"])
    prm.write(r, 4, grade, F["desc"])
    prm.write(r, 5, weight, F["in_pct"])
    prm.write(r, 6, regimes, F["in_ctr"])
    prm.write(r, 7, ref, F["ref"])
    r += 1
prm.write(r, 3, "ST_ = every transaction of the netting set has an original maturity of 3 months or less. "
          "A_ENH = SCRA grade A with CET1 >= 14% and leverage ratio >= 5%.", F["ref"])

r += 2
prm.set_row(r, 26)
prm.write(r, 2, "FX rates (illustrative - replace with market rates)", F["section"])
r += 1
prm.set_row(r, 30)
for c, h in enumerate(["FX_Ccy", "Currency", "Units of reporting ccy per 1 unit"]):
    prm.write(r, 2 + c, h, F["hdr_prm"])
name("FX_Table", "Params", r, 2)
r += 1
CCY_NAMES = {"EUR": "Euro", "USD": "US dollar", "GBP": "Pound sterling", "JPY": "Japanese yen", "CHF": "Swiss franc"}
for ccy, rate in FIXTURE["fx"].items():
    prm.write(r, 2, ccy, F["code"])
    prm.write(r, 3, CCY_NAMES.get(ccy, ""), F["desc"])
    prm.write(r, 4, rate, F["in_rate"])
    r += 1
prm.write(r, 3, "Add currencies directly below the last row (no blank rows in column C).", F["ref"])
prm.freeze_panes(5, 0)

# ----------------------------------------------------------------------------
# NETTING SETS
# ----------------------------------------------------------------------------
NS_COLS = [
    ("NettingSetID", 18, "in_txt", "id"),
    ("Counterparty", 44, "in_txt", "counterparty"),
    ("Eligible master netting agreement (Y/N)", 12, "in_bool", "mna"),
    ("Settlement ccy (blank = reporting)", 11, "in_ctr", "settlement_ccy"),
    ("Counterparty class (SOV / INST / CORP / QCCP)", 13, "in_ctr", "class"),
    ("Counterparty CQS (1-6, blank = unrated)", 11, "in_ctr", "cqs"),
    ("SCRA grade (A / A_ENH / B / C)", 11, "in_ctr", "scra"),
    ("Remargin frequency N_R (BD)", 11, "in_num", "remargin_bd"),
    (">5,000 trades or illiquid collateral (Y/N)", 13, "in_bool", "large"),
    ("Margin disputes (Y/N)", 10, "in_bool", "disputes"),
    ("Holding period override (BD)", 11, "in_num", "holding_override"),
    ("0% haircut conditions met (Y/N)", 11, "in_bool", "zero_haircut"),
    ("Risk weight override", 11, "in_pct", "rw_override"),
    ("Regime override (blank = Params)", 11, "in_ctr", "regime"),
]
NS_ROWS = 200
ns = wb.add_worksheet("NettingSets")
ns.set_tab_color(TAB_IN)
banner(ns, "Netting sets", 2 + len(NS_COLS) - 1)
ns.merge_range(2, 2, 2, 2 + len(NS_COLS) - 1, "", F["note_wrap"])
ns.write(2, 2, "One row per netting set. Eligible master netting agreement = Y: all its transactions are netted "
         "(CRE22.62-22.65 / CRR Art. 220); N: each transaction is calculated on its own (CRE22.46 / Art. 223). "
         "An unrated institution needs its SCRA grade. Y in '0% haircut' declares that the conditions of "
         "CRE22.60 / Art. 227 are met.", F["note_wrap"])
ns.set_row(2, 44)
ns.set_row(5, 62)
HDR_NS = 5
for c, (h, w, style, _) in enumerate(NS_COLS):
    ns.set_column(2 + c, 2 + c, w)
    ns.write(HDR_NS, 2 + c, h, F["hdr_in"])
name("NettingSets_Header", "NettingSets", HDR_NS, 2)
for i in range(NS_ROWS):
    rr = HDR_NS + 1 + i
    src = FIXTURE["netting_sets"][i] if i < len(FIXTURE["netting_sets"]) else None
    for c, (h, w, style, key) in enumerate(NS_COLS):
        value = src[key] if src else ""
        if value == "" or value is None:
            ns.write_blank(rr, 2 + c, None, F[style])
        else:
            ns.write(rr, 2 + c, value, F[style])
last_ns = HDR_NS + NS_ROWS
def ns_range(col: int) -> str:
    return f"{xl_col_to_name(col)}{HDR_NS + 2}:{xl_col_to_name(col)}{last_ns + 1}"


rng = ns_range
ns.data_validation(rng(2), {"validate": "custom", "value": f"=COUNTIF($C${HDR_NS + 2}:$C${last_ns + 1},C{HDR_NS + 2})=1",
                            "error_message": "Netting-set IDs must be unique."})
for col in (4, 10, 11, 13):
    ns.data_validation(rng(col), {"validate": "list", "source": ["Y", "N"]})
ns.data_validation(rng(6), {"validate": "list", "source": ["SOV", "INST", "CORP", "QCCP"]})
ns.data_validation(rng(7), {"validate": "list", "source": ["1", "2", "3", "4", "5", "6"]})
ns.data_validation(rng(8), {"validate": "list", "source": ["A", "A_ENH", "B", "C"]})
ns.data_validation(rng(9), {"validate": "decimal", "criteria": ">=", "value": 1})
ns.data_validation(rng(12), {"validate": "decimal", "criteria": ">=", "value": 0})
ns.data_validation(rng(14), {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 12.5})
ns.data_validation(rng(15), {"validate": "list", "source": ["CRR", "BCBS"]})
ns.autofilter(HDR_NS, 2, last_ns, 2 + len(NS_COLS) - 1)
ns.freeze_panes(HDR_NS + 1, 3)
buttons(ns, 3, 4)
ns.set_row(3, 34)

# ----------------------------------------------------------------------------
# TRANSACTIONS
# ----------------------------------------------------------------------------
TX_COLS = [
    ("LegID", 13, "in_txt", "leg"),
    ("TransactionID", 13, "in_txt", "tx"),
    ("NettingSetID", 16, "in_txt", "ns"),
    ("Transaction type", 12, "in_ctr", "type"),
    ("Start date", 11, "in_date", "start"),
    ("Maturity date (blank = open)", 11, "in_date", "maturity"),
    ("Side (LENT / RECEIVED)", 11, "in_ctr", "side"),
    ("Asset type", 10, "in_ctr", "asset"),
    ("Issue ID", 17, "in_txt", "issue"),
    ("Issuer type (SOV / INST / CORP / SEC)", 11, "in_ctr", "issuer"),
    ("Issue CQS (1-6, blank = unrated)", 10, "in_ctr", "cqs"),
    ("Security maturity date", 11, "in_date", "sec_maturity"),
    ("Equity listing (MAIN / LISTED / OTHER)", 11, "in_ctr", "listing"),
    ("Fund haircut H10", 10, "in_pct", "fund_h"),
    ("Market value", 15, "in_amt", "mv"),
    ("Currency", 9, "in_ctr", "ccy"),
    ("Comment", 44, "in_txt", "comment"),
]
TX_ROWS = 5000
tx = wb.add_worksheet("Transactions")
tx.set_tab_color(TAB_IN)
banner(tx, "Transactions", 2 + len(TX_COLS) - 1)
tx.write(2, 2, "One row per leg: LENT = given by the bank (E), RECEIVED = received (C).", F["note"])
tx.set_row(2, 34)
HDR_TX = 3
tx.set_row(HDR_TX, 52)
for c, (h, w, style, _) in enumerate(TX_COLS):
    tx.set_column(2 + c, 2 + c, w)
    tx.write(HDR_TX, 2 + c, h, F["hdr_in"])
name("Transactions_Header", "Transactions", HDR_TX, 2)
for i in range(TX_ROWS):
    rr = HDR_TX + 1 + i
    src = FIXTURE["legs"][i] if i < len(FIXTURE["legs"]) else None
    for c, (h, w, style, key) in enumerate(TX_COLS):
        value = src[key] if src else ""
        if value == "" or value is None:
            tx.write_blank(rr, 2 + c, None, F[style])
        elif style == "in_date":
            tx.write_datetime(rr, 2 + c, date(value), F[style])
        else:
            tx.write(rr, 2 + c, value, F[style])
last_tx = HDR_TX + TX_ROWS
def tx_range(col: int) -> str:
    return f"{xl_col_to_name(col)}{HDR_TX + 2}:{xl_col_to_name(col)}{last_tx + 1}"


rng = tx_range
tx.data_validation(rng(2), {"validate": "custom", "value": f"=COUNTIF($C${HDR_TX + 2}:$C${last_tx + 1},C{HDR_TX + 2})=1",
                            "error_message": "Leg IDs must be unique."})
tx.data_validation(rng(5), {"validate": "list", "source": ["REPO", "REVREPO", "SECLEND", "SECBORROW", "MARGINLEND"]})
tx.data_validation(rng(8), {"validate": "list", "source": ["LENT", "RECEIVED"]})
tx.data_validation(rng(9), {"validate": "list", "source": ["CASH", "DEBT", "EQUITY", "GOLD", "FUND", "OTHER"]})
tx.data_validation(rng(11), {"validate": "list", "source": ["SOV", "INST", "CORP", "SEC"]})
tx.data_validation(rng(12), {"validate": "list", "source": ["1", "2", "3", "4", "5", "6"]})
tx.data_validation(rng(14), {"validate": "list", "source": ["MAIN", "LISTED", "OTHER"]})
tx.data_validation(rng(15), {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 1})
tx.data_validation(rng(16), {"validate": "decimal", "criteria": ">", "value": 0})
tx.autofilter(HDR_TX, 2, last_tx, 2 + len(TX_COLS) - 1)
tx.freeze_panes(HDR_TX + 1, 4)
buttons(tx, 2, 9)

# ----------------------------------------------------------------------------
# OUTPUT SHEETS
# ----------------------------------------------------------------------------


def output_sheet(title: str, tab_title: str, header_row: int, columns: list[tuple[str, int]], note: str | None,
                 header_name: str, freeze_col: int = 4, run_buttons: bool = False) -> Any:
    ws = wb.add_worksheet(tab_title)
    ws.set_tab_color(TAB_OUT)
    banner(ws, title, 2 + len(columns) - 1)
    if note:
        ws.write(2, 2, note, F["note"])
    ws.set_row(2, 30)
    ws.set_row(header_row, 44)
    for c, (h, w) in enumerate(columns):
        ws.set_column(2 + c, 2 + c, w)
        ws.write(header_row, 2 + c, h, F["hdr_out_l"] if c < 2 else F["hdr_out"])
    name(header_name, tab_title, header_row, 2)
    ws.freeze_panes(header_row + 1, freeze_col)
    if run_buttons:
        buttons(ws, 2, 9)
    return ws


RS_COLS = [("NettingSetID", 16), ("Counterparty", 40), ("Regime", 8), ("MNA", 6), ("Transactions used", 11),
           ("Legs used", 8), ("Holding period T_M (BD)", 9), ("Remargin N_R (BD)", 9), ("Haircut scale", 9),
           ("Sum E (lent)", 15), ("Sum C (received)", 15), ("Net exposure", 13), ("Gross exposure", 13),
           ("N (issues)", 7), ("0.4 x net", 13), ("0.6 x gross / sqrt(N)", 13), ("FX term", 13), ("E* (EAD)", 15),
           ("Exposure class", 9), ("Risk weight key", 16), ("Risk weight", 9), ("RWA", 15),
           ("Own funds requirement", 14), ("Status", 40)]
res = output_sheet("Results", "Results", 5, RS_COLS, None, "Results_Header", run_buttons=True)
res.write(3, 2, "No results yet: press Run SFT.", F["runinfo"])
res.write(4, 2, "Withheld netting sets (INVALID, INCOMPLETE) are shaded; totals add VALID netting sets of one "
          "regime only.", F["runinfo"])
res.conditional_format("C7:Z100000", {"type": "formula",
                                      "criteria": '=OR(LEFT($Z7,7)="INVALID",LEFT($Z7,10)="INCOMPLETE")',
                                      "format": F["cf_bad"]})

PS_COLS = [("NettingSetID", 16), ("Kind", 13), ("Position", 18), ("Asset type / transaction type", 12),
           ("Currency", 9), ("Legs", 6), ("Holding period T_M (BD)", 9), ("Lent / posted (E)", 15),
           ("Received / held (C)", 15), ("Net E_s = E - C", 15), ("H (scaled) / Hfx", 10), ("E_s x H_s", 13),
           ("|E_s| x H_s", 13), ("Counted in N / currency role", 13), ("FX term", 13), ("E x (1 + H)", 15),
           ("C x (1 - H - Hfx)", 15), ("E* (transaction)", 15)]
output_sheet("Positions", "Positions", 4, PS_COLS,
             "With netting: one row per security issue and per currency of each VALID netting set. Without "
             "netting: one row per transaction (single-transaction formula).", "Positions_Header")

LC_COLS = [("LegID", 13), ("TransactionID", 13), ("NettingSetID", 16), ("Status", 13), ("Transaction type", 11),
           ("Side", 10), ("Asset type", 9), ("Haircut key", 22), ("Eligible", 8), ("Residual maturity (years)", 10),
           ("Value (rep ccy)", 15), ("Signed value (E +, C -)", 15), ("H10", 9), ("Holding period T_M (BD)", 9),
           ("Scale", 8), ("H", 10), ("Hfx (no netting)", 10), ("Adjusted value", 15), ("Regime", 8),
           ("Messages", 70)]
output_sheet("Leg-level calculation", "LegCalc", 4, LC_COLS,
             "Every leg with its haircut; excluded legs show the reason in Messages.", "LegCalc_Header")

CK_COLS = [("Severity", 12), ("Sheet", 14), ("Row", 7), ("Reference", 26), ("Message", 120)]
ck = output_sheet("Checks", "Checks", 4, CK_COLS,
                  "ERROR: input rejected; the netting set's E* is withheld (INCOMPLETE or INVALID) or the run "
                  "stops. WARNING: accepted with an assumption. INFO: information only.", "Checks_Header", 3)
ck.conditional_format("C6:G100000", {"type": "formula", "criteria": '=$C6="ERROR"', "format": F["cf_bad"]})
ck.conditional_format("C6:G100000", {"type": "formula", "criteria": '=$C6="WARNING"', "format": F["cf_warn"]})

# ----------------------------------------------------------------------------
# TEST CATALOGUE
# ----------------------------------------------------------------------------
MEASURES: dict[str, tuple[str, str, float | None]] = {  # measure: (Results column letter, expected key, tolerance)
    "E*": ("T", "e_star", 0.01), "RWA": ("X", "rwa", 0.01), "Own funds": ("Y", "own_funds", 0.01),
    "Status": ("Z", "status", None), "Holding period": ("I", "holding_bd", 1e-9), "N": ("P", "n", 0),
    "Net exposure": ("N", "net", 0.01), "Gross exposure": ("O", "gross", 0.01), "FX term": ("S", "fx_term", 0.01),
    "Risk weight": ("W", "rw", 1e-12), "Risk weight key": ("V", "rw_key", None),
}
tc = wb.add_worksheet("TestCatalogue")
tc.set_tab_color(TAB_TEST)
banner(tc, "Test catalogue", 11)
for c, w in enumerate([12, 16, 8, 16, 18, 18, 9, 10, 46]):
    tc.set_column(2 + c, 2 + c, w)
tc.write(2, 2, "Expected values from tests/derivations/sft_reference.py (an independent implementation of the "
         "same specification) against Results after Run SFT. Class ILLUSTRATIVE: no published worked example "
         "exists for these portfolios.", F["note"])
tc.set_row(2, 30)
tc.set_row(5, 34)
HEAD = ["TestID", "NettingSetID", "Regime", "Measure", "Expected", "Calculated", "Tolerance", "Result", "Source"]
for c, h in enumerate(HEAD):
    tc.write(5, 2 + c, h, F["hdr_out_l"] if c < 2 else F["hdr_out"])
rows: list[tuple[str, str, str, Any]] = []
for nsid, res_ in EXPECTED.items():
    if "e_star" not in res_:
        rows.append((nsid, res_["regime"], "Status", res_["status"]))
        continue
    for m in ["E*", "RWA", "Own funds", "Status", "Holding period", "Risk weight", "Risk weight key"]:
        rows.append((nsid, res_["regime"], m, res_[MEASURES[m][1]]))
    if "n" in res_:
        for m in ["Net exposure", "Gross exposure", "N", "FX term"]:
            rows.append((nsid, res_["regime"], m, res_[MEASURES[m][1]]))
r = 6
num = fmt(font_size=10, num_format="#,##0.00")
for k, (nsid, regime, measure, expected) in enumerate(rows, 1):
    letter, _, tol = MEASURES[measure]
    tc.write(r, 2, f"SFT-{k:03d}", F["out_txt"])
    tc.write(r, 3, nsid, F["out_txt"])
    tc.write(r, 4, regime, F["out_txt"])
    tc.write(r, 5, measure, F["out_txt"])
    tc.write(r, 6, expected, num if isinstance(expected, (int, float)) else F["out_txt"])
    calc = f'=IFERROR(INDEX(Results!${letter}:${letter},MATCH($D{r + 1},Results!$C:$C,0)),"")'
    tc.write_formula(r, 7, calc, num if tol is not None else F["out_txt"], "")
    if tol is None:
        tc.write(r, 8, "exact", F["out_txt"])
        res_f = f'=IF(H{r + 1}="","NOT RUN",IF(H{r + 1}=G{r + 1},"PASS","FAIL"))'
    else:
        tc.write(r, 8, tol, num)
        res_f = (f'=IF(H{r + 1}="","NOT RUN",IF(ISNUMBER(H{r + 1}),IF(ABS(H{r + 1}-G{r + 1})<=I{r + 1},'
                 f'"PASS","FAIL"),"FAIL"))')
    tc.write_formula(r, 9, res_f, F["out_txt"], "NOT RUN")
    tc.write(r, 10, "tests/derivations/prototype/demo-portfolio.expected.json", F["ref"])
    r += 1
last_tc = r
tc.write_formula(3, 2, f'=COUNTIF(J7:J{last_tc},"PASS")&" of "&COUNTA(C7:C{last_tc})&" tests PASS; "&'
                       f'COUNTIF(J7:J{last_tc},"FAIL")&" FAIL; "&COUNTIF(J7:J{last_tc},"NOT RUN")&" NOT RUN"',
                 fmt(font_size=11, bold=True, font_color=NAVY), "")
tc.conditional_format(f"J7:J{last_tc}", {"type": "cell", "criteria": "==", "value": '"PASS"', "format": F["cf_pass"]})
tc.conditional_format(f"J7:J{last_tc}", {"type": "cell", "criteria": "==", "value": '"FAIL"', "format": F["cf_fail"]})
tc.freeze_panes(6, 4)

# ----------------------------------------------------------------------------
# FORMULA TEST
# ----------------------------------------------------------------------------
ft = wb.add_worksheet("FormulaTest")
ft.set_tab_color(TAB_TEST)
banner(ft, "Formula tests", 8)
for c, w in enumerate([10, 60, 16, 16, 10, 50]):
    ft.set_column(2 + c, 2 + c, w)
ft.write(2, 2, "Each row calls one SFT_* function; Result compares it with a hand-calculated value within the "
         "tolerance. The cells show #NAME? until the VBA modules are imported.", F["note"])
ft.set_row(2, 30)
ft.write(3, 2, "Tolerance", F["code"])
ft.write(3, 3, 1e-9, F["in_num"])
wb.define_name("FormulaTest_Tolerance", "='FormulaTest'!$D$4")
ft.set_row(5, 30)
for c, h in enumerate(["TestID", "Formula", "Value", "Expected", "Result", "Hand calculation"]):
    ft.write(5, 2 + c, h, F["hdr_out_l"] if c < 2 else F["hdr_out"])
TESTS = [
    ("SFT_HoldingPeriod(5)", 5, "repo-style minimum"),
    ("SFT_HoldingPeriod(5,TRUE)", 20, "large or illiquid netting set"),
    ("SFT_HoldingPeriod(5,TRUE,TRUE)", 40, "20 doubled for disputes"),
    ("SFT_HoldingPeriod(10,FALSE,FALSE,15)", 15, "override lengthens"),
    ("SFT_HoldingPeriod(10,FALSE,FALSE,7)", 10, "override never shortens"),
    ("SFT_HoldingPeriod(0)", "#NUM!", "minimum below 1"),
    ("SFT_ScaledHaircut(0.02,5)", 0.02 * (5 / 10) ** 0.5, "0.02 x sqrt(5/10): Art. 224 5-day column 1.414%"),
    ("SFT_ScaledHaircut(0.04,20,3)", 0.04 * (22 / 10) ** 0.5, "0.04 x sqrt((3+20-1)/10)"),
    ("SFT_ScaledHaircut(0.08,10)", 0.08, "Hfx at 10 days"),
    ("SFT_ScaledHaircut(0.02,0)", "#NUM!", "holding period below 1"),
    ("SFT_ExposureSingle(100,0,102,0.02)", 100 - 102 * 0.98, "100 - 102 x 0.98"),
    ("SFT_ExposureSingle(100,0,102,0.02,0.08)", 100 - 102 * 0.90, "100 - 102 x (1 - 0.02 - 0.08)"),
    ("SFT_ExposureSingle(100,0,120,0.02)", 0, "over-collateralised: floored at 0"),
    ("SFT_ExposureSingle(-1,0,1,0)", "#NUM!", "negative E"),
    ("SFT_NetExposure({100,-50},{0.02,0.04})", 0, "|100 x 0.02 - 50 x 0.04|"),
    ("SFT_GrossExposure({100,-50},{0.02,0.04})", 4, "100 x 0.02 + 50 x 0.04"),
    ("SFT_IssueCount({100,-50,5})", 2, "5 < 100 / 10 is not counted"),
    ("SFT_IssueCount({100,-50,10})", 3, "10 = 100 / 10 is counted"),
    ("SFT_IssueCount({0,0})", 0, "no position"),
    ("SFT_ExposureNetted(100,102,2.04,2.04,1)", 100 - 102 + 2.04, "N = 1: equals the single formula"),
    ("SFT_ExposureNetted(0,0,1,4,4,0.5)", 0.4 + 0.6 * 4 / 2 + 0.5, "0.4 x 1 + 0.6 x 4 / sqrt(4) + 0.5"),
    ("SFT_ExposureNetted(0,0,0,1,0)", "#NUM!", "gross > 0 with N = 0"),
    ("SFT_CapitalRequirement(1000000,0.2)", 16000, "1,000,000 x 20% x 8%"),
    ("SFT_CapitalRequirement(1000000,0.2,1)", 200000, "RWA"),
    ("SFT_CapitalRequirement(1,13)", "#NUM!", "risk weight above 1250%"),
]
r = 6
for k, (formula, expected, note) in enumerate(TESTS, 1):
    ft.write(r, 2, f"F-{k:02d}", F["out_txt"])
    ft.write_string(r, 3, "=" + formula, F["out_txt"])
    ft.write_formula(r, 4, "=" + formula, fmt(font_size=10, num_format="0.000000000"), "")
    ft.write(r, 5, expected, fmt(font_size=10, num_format="0.000000000") if not isinstance(expected, str) else F["out_txt"])
    rr = r + 1
    ft.write_formula(r, 6, f'=IF(ISERROR(E{rr}),IF(F{rr}="#NUM!",IF(ERROR.TYPE(E{rr})=6,"PASS","FAIL"),"FAIL"),'
                           f'IF(ISNUMBER(F{rr}),IF(ABS(E{rr}-F{rr})<=FormulaTest_Tolerance,"PASS","FAIL"),"FAIL"))',
                     F["out_txt"], "")
    ft.write(r, 7, note, F["ref"])
    r += 1
ft.write_formula(3, 5, f'=COUNTIF(G7:G{r},"PASS")&" of "&COUNTA(C7:C{r})&" PASS"',
                 fmt(font_size=11, bold=True, font_color=NAVY), "")
ft.conditional_format(f"G7:G{r}", {"type": "cell", "criteria": "==", "value": '"PASS"', "format": F["cf_pass"]})
ft.conditional_format(f"G7:G{r}", {"type": "cell", "criteria": "==", "value": '"FAIL"', "format": F["cf_fail"]})
ft.freeze_panes(6, 0)

# ----------------------------------------------------------------------------
# BASEL VS CRR
# ----------------------------------------------------------------------------
bc = wb.add_worksheet("Basel_vs_CRR")
bc.set_tab_color(TAB_DOC)
banner(bc, "Basel vs CRR", 7)
for c, w in enumerate([22, 46, 46, 14, 46, 26]):
    bc.set_column(2 + c, 2 + c, w)
bc.write(2, 2, "Basel CRE22 (effective 1 January 2023) and CRE20 against the consolidated CRR of 26.06.2026, "
         "checked on 2026-10-24.", F["note"])
bc.set_row(2, 30)
bc.set_row(4, 34)
for c, h in enumerate(["Topic", "BCBS (Basel framework)", "EU CRR", "Same in both?", "Workbook treatment",
                       "References"]):
    bc.write(4, 2 + c, h, F["hdr_in"])
BVC = [
    ("Exposure method", "Comprehensive approach with supervisory haircuts; VaR models and IMM as alternatives; "
     "SA-CCR does not apply to SFTs.", "Financial collateral comprehensive method with supervisory volatility "
     "adjustments; own estimates, VaR and IMM as alternatives.", "Yes", "Supervisory haircuts only.",
     "CRE22.40-22.65; CRR Art. 220-227, 271(1)"),
    ("Haircut table", "10-day haircuts by rating, residual maturity and issuer (sovereign, other, "
     "securitisation).", "Art. 224 Table 1 with 20-, 10- and 5-day columns; the 10-day column equals CRE22.49 "
     "since CRR3.", "Yes", "One H10 table on Params, scaled by the holding period.", "CRE22.49; Art. 224 Table 1"),
    ("Equities and gold", "Main index 20%, other listed 30%, gold 20%.", "Table 3: main index 20%, other "
     "listed 30%, gold 20% (10-day).", "Yes", "Rows EQUITY_MAIN, EQUITY_LISTED, GOLD.",
     "CRE22.49; Art. 224 Table 3"),
    ("Currency mismatch", "Hfx 8% at 10 days.", "Table 4: 8% at 10 days.", "Yes", "FXHaircut, scaled; "
     "netting uses the settlement currency of the agreement.", "CRE22.52, 22.65; Art. 220(3), 224 Table 4"),
    ("Holding periods", "5 BD repo-style, 10 BD other capital market, 20 BD secured lending; 10 BD when a "
     "netting set mixes repo-style and other capital market transactions.", "5, 10 and 20 BD; aligned with "
     "the margin period of risk of Art. 285(2)-(4) for large or illiquid netting sets.", "Yes, except the "
     "mixed rule", "The longer minimum (10) for mixed netting sets in both regimes; secured lending out of "
     "scope.", "CRE22.57-22.58; Art. 224(2)"),
    ("20-day holding period and disputes", "More than 5,000 trades, illiquid collateral; doubled after more "
     "than two disputes.", "Through Art. 285(3)-(4).", "Yes", "Flags on NettingSets.",
     "CRE22.58; Art. 224(2), 285(3)-(4)"),
    ("Scaling", "H = H10 x sqrt((N_R + T_M - 1) / 10).", "H = H_M x sqrt((N_R + T_M - 1) / T_M), with H_M "
     "the T_M column.", "Yes (same figure)", "SFT_ScaledHaircut.", "CRE22.59; Art. 226"),
    ("Netting formula", "E* = max(0, sum E - sum C + 0.4 net + 0.6 gross / sqrt(N) + FX), N excluding issues "
     "below one tenth of the largest.", "Same formula since CRR3; counts groups of the same securities and "
     "commodity types.", "Yes", "SFT_ExposureNetted; commodities out of scope.", "CRE22.64-22.65; Art. 220(3)"),
    ("Non-eligible securities lent", "Haircut 30%.", "Haircut of non-main-index listed equities (30%).", "Yes",
     "NonEligibleHaircut; non-eligible collateral received is not recognised.", "CRE22.54; Art. 224(4)"),
    ("Funds", "Highest haircut of any security the fund can invest in, or look-through.", "Weighted average "
     "of the underlying, or the highest when unknown.", "Yes", "Entered as Fund haircut H10.",
     "CRE22.49; Art. 224(5)"),
    ("0% haircut", "Repo-style transactions with core market participants meeting the conditions.",
     "Art. 227 conditions.", "Yes", "Declared per netting set; observable conditions checked.",
     "CRE22.60, 22.36-22.37; Art. 227"),
    ("Haircut floors", "Minimum haircut floors for non-centrally cleared SFTs with some counterparties.",
     "Not in the consolidated text.", "No", "Not implemented in either regime.", "CRE56"),
    ("Institution risk weights", "ECRA 20/30/50/100/150; short-term 20/20/20/50/150; SCRA 40 (30) / 75 / "
     "150, short-term 20/50/150; original maturity of 3 months or less.", "Art. 120-121: the same weights "
     "and original-maturity test.", "Yes", "RW_Table; short-term only when every transaction qualifies; the "
     "SCRA sovereign floor is not applied.", "CRE20.18-20.21; Art. 120-121"),
    ("Corporate and sovereign risk weights", "Sovereigns 0/20/50/100/150; corporates 20/50/75/100/150, "
     "unrated 100%.", "Art. 114, 122: the same weights.", "Yes", "RW_Table; transitional and output-floor "
     "treatments not modelled.", "CRE20.7, CRE20 corporates; Art. 114, 122"),
    ("QCCP", "2% on trade exposures of a clearing member.", "2% under Art. 306(1)(a).", "Yes", "Class QCCP.",
     "CRE54.7; Art. 306(1)(a)"),
    ("Capital requirement", "Total capital at least 8% of RWA.", "Art. 92(1)(c).", "Yes", "CapitalRatio; "
     "buffers and Pillar 2 not included.", "RBC20.1; Art. 92(1)(c)"),
]
r = 5
for row in BVC:
    bc.set_row(r, 62)
    for c, v in enumerate(row):
        bc.write(r, 2 + c, v, F["rd_k"] if c == 0 else F["rd_v"])
    r += 1
bc.freeze_panes(5, 3)

# Code names, sh + noun (docs/VBA_HOUSE_STYLE.md). Without them Excel assigns
# defaults when the VBA project is first opened, which the import script
# reports as a changed sheet structure.
CODE_NAMES = {"README": "shReadme", "Params": "shParams", "NettingSets": "shNettingSets",
              "Transactions": "shTransactions", "Results": "shResults", "Positions": "shPositions",
              "LegCalc": "shLegCalc", "Checks": "shChecks", "TestCatalogue": "shTestCatalogue",
              "FormulaTest": "shFormulaTest", "Basel_vs_CRR": "shBaselVsCRR"}
wb.set_vba_name("ThisWorkbook")
for sheet in wb.worksheets():
    sheet.set_vba_name(CODE_NAMES[sheet.get_name()])

wb.close()
print("written", OUT.relative_to(ROOT), "-", len(rows), "catalogue tests,", len(TESTS), "formula tests")

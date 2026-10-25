"""Independent reference implementation of the K-SFTs calculation.

Computes, for every netting set of a fixture, the fully adjusted exposure
value E* of securities financing transactions under the financial collateral
comprehensive method with supervisory haircuts, the risk weight, the
risk-weighted exposure amount and the own funds requirement, following the
specification in docs/methodology/README.md:

  single transaction (no eligible master netting agreement)
      E* = max(0, sum_i E_i (1 + H_i) - sum_j C_j (1 - H_j - Hfx_j))
                                            [CRE22.46; CRR Art. 223(5)]
  netting set under an eligible master netting agreement
      E* = max(0, sum E - sum C + 0.4 net + 0.6 gross / sqrt(N)
                  + sum_k |E_k^fx| Hfx)       [CRE22.65; CRR Art. 220(3)]
  haircut scaling   H = H10 sqrt((N_R + T_M - 1) / 10)
                                            [CRE22.59; CRR Art. 226]
  RWA = E* x RW;  own funds requirement = 8% x RWA

The VBA engine and this module are written from the same specification but
share no code; tests/derivations/prototype/demo-portfolio.expected.json is generated from this module and the
TestCatalogue sheet compares the workbook with it.

Usage:
  python tests/derivations/sft_reference.py              # print a summary
  python tests/derivations/sft_reference.py --write      # write tests/derivations/prototype
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
REPO_STYLE = {"REPO", "REVREPO", "SECLEND", "SECBORROW"}
CAPITAL_MARKET = {"MARGINLEND"}
TYPES = REPO_STYLE | CAPITAL_MARKET
ASSETS = {"CASH", "DEBT", "EQUITY", "GOLD", "FUND", "OTHER"}
CLASSES = {"SOV", "INST", "CORP", "QCCP"}


def date(text):
    if text in ("", None):
        return None
    return dt.date.fromisoformat(text)


def yes(text):
    return str(text).strip().upper() in ("Y", "YES", "TRUE", "1")


def add_months(day: dt.date, months: int) -> dt.date:
    """VBA DateAdd("m") semantics: same day, clipped to the month's end."""
    month = day.month - 1 + months
    year = day.year + month // 12
    month = month % 12 + 1
    for d in range(day.day, 27, -1):
        try:
            return dt.date(year, month, d)
        except ValueError:
            continue
    return dt.date(year, month, day.day)


def maturity_band(years: float) -> str:
    if years <= 1:
        return "LE1"
    if years <= 3:
        return "1TO3"
    if years <= 5:
        return "3TO5"
    if years <= 10:
        return "5TO10"
    return "GT10"


def haircut_key(leg, as_of, days_per_year):
    """Return (key, residual maturity years or None, note) of a leg."""
    asset = leg["asset"]
    if asset == "CASH":
        return "CASH", None, ""
    if asset == "GOLD":
        return "GOLD", None, ""
    if asset == "EQUITY":
        listing = str(leg["listing"]).upper()
        if listing == "MAIN":
            return "EQUITY_MAIN", None, ""
        if listing == "LISTED":
            return "EQUITY_LISTED", None, ""
        return "", None, "equity not listed"
    if asset == "DEBT":
        mat = date(leg["sec_maturity"])
        years = (mat - as_of).days / days_per_year
        issuer = leg["issuer"]
        cat = {"SOV": "SOV", "INST": "OTH", "CORP": "OTH", "SEC": "SEC"}[issuer]
        cqs = leg["cqs"]
        note = ""
        if cqs in ("", None):
            if issuer == "INST":
                band = "CQS23"
                note = "unrated institution debt treated as CQS 2-3"
            else:
                return "", years, "unrated debt"
        else:
            cqs = int(cqs)
            band = {1: "CQS1", 2: "CQS23", 3: "CQS23", 4: "CQS4"}.get(cqs, "CQS56")
        if band == "CQS4":
            return f"DEBT_{cat}_CQS4_ALL", years, note
        return f"DEBT_{cat}_{band}_{maturity_band(years)}", years, note
    return "", None, "asset type not eligible"


def holding_period(types, ns, p):
    base = max(p["HoldingRepoBD"] if t in REPO_STYLE else p["HoldingCapitalMarketBD"] for t in types)
    if yes(ns["large"]):
        base = max(base, p["HoldingLargeBD"])
    if ns["holding_override"] not in ("", None):
        base = max(base, float(ns["holding_override"]))
    if yes(ns["disputes"]):
        base *= 2
    return base


def risk_weight_key(ns, short_term):
    cls = ns["class"]
    cqs = ns["cqs"]
    if cls == "QCCP":
        return "QCCP"
    if cls == "SOV":
        return f"SOV_CQS{int(cqs)}" if cqs not in ("", None) else "SOV_UNRATED"
    if cls == "CORP":
        return f"CORP_CQS{int(cqs)}" if cqs not in ("", None) else "CORP_UNRATED"
    st = "ST_" if short_term else ""
    if cqs not in ("", None):
        return f"INST_{st}CQS{int(cqs)}"
    return f"INST_{st}SCRA_{ns['scra']}"


def calculate(fixture, tables):
    p = fixture["params"]
    as_of = date(p["AsOfDate"])
    dpy = p["DaysPerYear"]
    fx = fixture["fx"]
    hc = {r[0]: r[5] for r in tables["haircuts"]}
    rw = {r[0]: r[3] for r in tables["risk_weights"]}
    results = {}
    legs_by_ns = {}
    for leg in fixture["legs"]:
        legs_by_ns.setdefault(leg["ns"], []).append(leg)
    for ns in fixture["netting_sets"]:
        regime = ns["regime"] or p["Regime"]
        legs = legs_by_ns.get(ns["id"], [])
        res = {"regime": regime, "status": "VALID"}
        if not legs:
            res["status"] = "NO TRANSACTIONS"
            results[ns["id"]] = res
            continue
        # group legs into transactions and reject invalid ones
        txs = {}
        for leg in legs:
            txs.setdefault(leg["tx"], []).append(leg)
        rejected = [t for t, ls in txs.items() if any(lg["asset"] not in ASSETS or lg["type"] not in TYPES for lg in ls)]
        if rejected:
            res["status"] = f"INCOMPLETE: {len(rejected)} of {len(txs)} transaction(s) rejected"
            results[ns["id"]] = res
            continue
        settle = ns["settlement_ccy"] or p["ReportingCcy"]
        nr = float(ns["remargin_bd"]) if ns["remargin_bd"] not in ("", None) else 1.0
        zero = yes(ns["zero_haircut"])
        mna = yes(ns["mna"])

        def leg_view(leg, t_m):
            key, years, _ = haircut_key(leg, as_of, dpy)
            value = leg["mv"] * fx[leg["ccy"]]
            scale = math.sqrt((nr + t_m - 1) / p["HaircutBaseBD"])
            if leg["asset"] == "FUND":
                h10, eligible = float(leg["fund_h"]), True
            elif key in hc:
                h10, eligible = hc[key], True
            else:
                h10, eligible = p["NonEligibleHaircut"], False
            h = 0.0 if zero else h10 * scale
            hfx = 0.0 if zero else p["FXHaircut"] * scale
            recognised = eligible or leg["side"] == "LENT"
            return dict(value=value, h=h, hfx=hfx, recognised=recognised, sign=1 if leg["side"] == "LENT" else -1)

        short_term = all(
            lg["maturity"] and add_months(date(lg["start"]), 3) >= date(lg["maturity"])
            for ls in txs.values() for lg in ls)

        if mna:
            t_m = holding_period({lg["type"] for lg in legs}, ns, p)
            views = [(lg, leg_view(lg, t_m)) for lg in legs]
            views = [(lg, v) for lg, v in views if v["recognised"]]
            sum_e = sum(v["value"] for lg, v in views if v["sign"] > 0)
            sum_c = sum(v["value"] for lg, v in views if v["sign"] < 0)
            issues = {}
            ccys = {}
            for lg, v in views:
                ccys[lg["ccy"]] = ccys.get(lg["ccy"], 0.0) + v["sign"] * v["value"]
                if lg["asset"] == "CASH":
                    continue
                e = issues.setdefault(lg["issue"].upper(), {"es": 0.0, "h": v["h"]})
                e["es"] += v["sign"] * v["value"]
            net = abs(sum(i["es"] * i["h"] for i in issues.values()))
            gross = sum(abs(i["es"]) * i["h"] for i in issues.values())
            largest = max([abs(i["es"]) for i in issues.values()] or [0.0])
            n = sum(1 for i in issues.values()
                    if abs(i["es"]) > 0 and abs(i["es"]) >= p["IssueCountThreshold"] * largest)
            hfx = 0.0 if zero else p["FXHaircut"] * math.sqrt((nr + t_m - 1) / p["HaircutBaseBD"])
            fx_term = sum(abs(v) * hfx for k, v in ccys.items() if k != settle)
            sys_term = p["NetWeight"] * net
            idio_term = p["GrossWeight"] * gross / math.sqrt(n) if n > 0 else 0.0
            e_star = max(0.0, sum_e - sum_c + sys_term + idio_term + fx_term)
            res.update(holding_bd=t_m, sum_e=sum_e, sum_c=sum_c, net=net, gross=gross, n=n,
                       systematic=sys_term, idiosyncratic=idio_term, fx_term=fx_term)
        else:
            e_star = 0.0
            sum_e = sum_c = fx_term = 0.0
            t_max = 0.0
            for t, ls in txs.items():
                t_m = holding_period({ls[0]["type"]}, ns, p)
                t_max = max(t_max, t_m)
                views = [(lg, leg_view(lg, t_m)) for lg in ls]
                exp_ccy = {lg["ccy"] for lg, v in views if v["sign"] > 0}
                assert len(exp_ccy) == 1
                exp_ccy = exp_ccy.pop()
                ev = sum(v["value"] * (1 + v["h"]) for lg, v in views if v["sign"] > 0 and v["recognised"])
                cv = 0.0
                for lg, v in views:
                    if v["sign"] < 0 and v["recognised"]:
                        h_fx = v["hfx"] if lg["ccy"] != exp_ccy else 0.0
                        cv += v["value"] * (1 - v["h"] - h_fx)
                        fx_term += v["value"] * h_fx
                        sum_c += v["value"]
                sum_e += sum(v["value"] for lg, v in views if v["sign"] > 0 and v["recognised"])
                e_star += max(0.0, ev - cv)
            res.update(holding_bd=t_max, sum_e=sum_e, sum_c=sum_c, fx_term=fx_term)
        key = risk_weight_key(ns, short_term)
        weight = float(ns["rw_override"]) if ns["rw_override"] not in ("", None) else rw[key]
        res.update(e_star=e_star, rw_key=key, rw=weight, rwa=e_star * weight,
                   own_funds=e_star * weight * p["CapitalRatio"])
        results[ns["id"]] = res
    return results


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    fixture = json.loads((ROOT / "tests/derivations/prototype/demo-portfolio.json").read_text())
    tables = json.loads((ROOT / "tests/derivations/prototype/regulatory-tables.json").read_text())
    results = calculate(fixture, tables)
    for k, r in results.items():
        if "e_star" in r:
            print(f"{k:15s} {r['regime']:5s} T={r['holding_bd']:>4} E*={r['e_star']:>16,.2f} "
                  f"RW={r['rw']:.2%} RWA={r['rwa']:>16,.2f} OF={r['own_funds']:>14,.2f} {r['status']}")
        else:
            print(f"{k:15s} {r['regime']:5s} {r['status']}")
    if args.write:
        out = ROOT / "tests/derivations/prototype/demo-portfolio.expected.json"
        out.write_text(json.dumps({
            "source": "tests/derivations/sft_reference.py on tests/derivations/prototype/demo-portfolio.json",
            "results": results}, indent=1) + "\n")
        print("written", out.relative_to(ROOT))


if __name__ == "__main__":
    main()

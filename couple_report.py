#!/usr/bin/env python3
"""
Couple Monthly Financial Report Generator
Export your Google Sheet as CSV, place it in the same folder as this script,
and run:  python couple_report.py
Optional args:
  --csv   path/to/file.csv
  --month 3  (default: latest month in data)
  --year  2026
  --out   path/to/output.pdf

Dependencies: pip install pandas matplotlib seaborn fpdf2
"""

import argparse
import os
import sys
import warnings
from datetime import datetime

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
import seaborn as sns
from fpdf import FPDF

warnings.filterwarnings("ignore")

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_FILENAME = "couple_finances.csv"  # rename to match your export
PDF_FILENAME = "Monthly_Financial_Report.pdf"
SPLURGE_LIMIT = 5_000  # $ per person per year

PERSONS = ["Isa", "Toio"]

EXPENSE_CATEGORIES = []  # populated dynamically from CSV after load()

MONTH_NAMES = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

_TEMP_IMGS = []  # cleaned up after PDF is written


# ── PDF CLASS ─────────────────────────────────────────────────────────────────
class PDF(FPDF):
    def __init__(self, month_label):
        super().__init__()
        self._lbl = month_label

    def header(self):
        self.set_fill_color(41, 128, 185)
        self.rect(0, 0, 210, 11, "F")
        self.set_font("Arial", "B", 8)
        self.set_text_color(255, 255, 255)
        self.set_xy(8, 1.5)
        self.cell(80, 8, "GetReady Intelligence · Couple Financial Tracker")
        self.set_font("Arial", "", 8)
        self.set_xy(90, 1.5)
        self.cell(75, 8, f"Report Period: {self._lbl}", align="C")
        self.set_font("Arial", "B", 8)
        self.set_xy(155, 1.5)
        self.cell(45, 8, f"Page {self.page_no()}", align="R")
        self.set_text_color(0, 0, 0)
        self.set_y(18)  # consistent top margin on every page

    def footer(self):
        self.set_y(-11)
        self.set_font("Arial", "I", 7)
        self.set_text_color(150, 150, 150)
        self.cell(
            0,
            8,
            f"Generated {datetime.now().strftime('%d %b %Y %H:%M')} · GetReady Intelligence",
            align="C",
        )
        self.set_text_color(0, 0, 0)


# ── HELPERS ───────────────────────────────────────────────────────────────────
def _clean(val):
    if pd.isna(val):
        return 0.0
    return float(str(val).replace("$", "").replace(",", "").strip() or 0)


def fc(v):
    return f"-${abs(v):,.2f}" if v < 0 else f"${v:,.2f}"


def fp(v):
    return f"{v:.1f}%"


def delta(cur, ref):
    if ref == 0:
        return "N/A"
    d = (cur - ref) / abs(ref) * 100
    return ("+" if d > 0 else "-") + f" {abs(d):.1f}%"


def save_fig(fig, name):
    p = os.path.join(SCRIPT_DIR, name)
    fig.savefig(p, bbox_inches="tight", dpi=150)
    plt.close(fig)
    _TEMP_IMGS.append(p)
    return p


def cleanup():
    for f in _TEMP_IMGS:
        try:
            os.remove(f)
        except Exception:
            pass


def _bar_style(ax, currency=True):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    if currency:
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))


# ── DATA ──────────────────────────────────────────────────────────────────────
def load(path):
    global EXPENSE_CATEGORIES
    df = pd.read_csv(path)
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True)
    df["Amount"] = df["Amount"].apply(_clean)
    df["Name"] = df["Name"].fillna("").astype(str).str.strip().replace("", "Joint")
    df["Splurge"] = df["Splurge"].fillna("").astype(str).str.strip()
    df["Category"] = df["Category"].str.strip()
    df["_yr"] = df["Date"].dt.year
    df["_mo"] = df["Date"].dt.month
    df["_ym"] = df["Date"].dt.to_period("M")
    df["_exp"] = df["Category"] != "Income"
    df["_splurge"] = df["Splurge"].str.lower() == "splurge"
    EXPENSE_CATEGORIES = sorted(
        df.loc[df["_exp"], "Category"].dropna().unique().tolist()
    )
    return df


def period(df, y, m):
    return df[(df["_yr"] == y) & (df["_mo"] == m)].copy()


def prior_3m(df, cy, cm):
    """Return (DataFrame, n_months) for the 3 complete months before (cy, cm)."""
    targets = set()
    for i in range(1, 4):
        mo, yr = cm - i, cy
        if mo <= 0:
            mo += 12
            yr -= 1
        targets.add((yr, mo))
    mask = [(yr, mo) in targets for yr, mo in zip(df["_yr"], df["_mo"])]
    sub = df[mask].copy()
    n = sub["_ym"].nunique()
    return sub, n


def summarize(d, divisor=1):
    """Income, expenses, net savings, savings rate. divisor for monthly avg."""
    if d is None or d.empty:
        return dict(income=0.0, expenses=0.0, net=0.0, rate=0.0)
    inc = d[d["Category"] == "Income"]["Amount"].sum() / divisor
    exp = d[d["_exp"]]["Amount"].sum() / divisor
    net = inc - exp
    rate = net / inc * 100 if inc else 0.0
    return dict(income=inc, expenses=exp, net=net, rate=rate)


def cat_totals(d, divisor=1):
    if d is None or d.empty:
        return pd.Series({c: 0.0 for c in EXPENSE_CATEGORIES})
    exp = d[d["_exp"]]
    return (
        exp.groupby("Category")["Amount"]
        .sum()
        .reindex(EXPENSE_CATEGORIES, fill_value=0)
        / divisor
    )


def pfilter(d, person):
    """Filter df by person name."""
    if d is None or d.empty:
        return pd.DataFrame()
    return d[d["Name"] == person].copy()


def summarize_person(person_df, all_df, divisor=1):
    """Person's own income vs their 50% share of ALL couple expenses."""
    inc = (
        person_df[person_df["Category"] == "Income"]["Amount"].sum() / divisor
        if person_df is not None and not person_df.empty
        else 0.0
    )
    exp = (
        all_df[all_df["_exp"]]["Amount"].sum() / divisor / 2
        if all_df is not None and not all_df.empty
        else 0.0
    )
    net = inc - exp
    rate = net / inc * 100 if inc else 0.0
    return dict(income=inc, expenses=exp, net=net, rate=rate)


def combined_person_df(d, person):
    """Return person's rows plus Joint rows with Amount halved."""
    own = pfilter(d, person)
    joint = pfilter(d, "Joint").copy()
    if not joint.empty:
        joint["Amount"] = joint["Amount"] / 2
    return pd.concat([own, joint], ignore_index=True)


# ── CHARTS ────────────────────────────────────────────────────────────────────
def chart_cat_compare(cm_df, smly_df, avg_df, n90, fname):
    cats = EXPENSE_CATEGORIES
    cm_v = cat_totals(cm_df).values
    sm_v = cat_totals(smly_df).values
    av_v = cat_totals(avg_df, n90).values if n90 else np.zeros(len(cats))
    x, w = np.arange(len(cats)), 0.26

    fig, ax = plt.subplots(figsize=(15, 5))
    ax.bar(x - w, cm_v, w, label="This Month", color="#2980b9")
    ax.bar(x, sm_v, w, label="Same Month LY", color="#e74c3c")
    ax.bar(x + w, av_v, w, label="90-Day Avg", color="#f39c12", alpha=0.85)
    ax.set_xticks(x)
    ax.set_xticklabels(cats, rotation=45, ha="right", fontsize=8)
    ax.set_title(
        "Category Spend: This Month vs Same Month LY vs 90-Day Avg", fontweight="bold"
    )
    ax.legend(fontsize=9)
    _bar_style(ax)
    fig.tight_layout()
    return save_fig(fig, fname)


def chart_trend(df, cy, cm, fname):
    ml = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ]
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    for ax, yr in zip(axes, [cy - 1, cy]):
        d = df[df["_yr"] == yr]
        if yr == cy:
            d = d[d["_mo"] <= cm]  # exclude current incomplete month
        if d.empty:
            ax.set_title(f"{yr} - No Data")
            continue
        mo = (
            d.groupby("_mo")
            .apply(
                lambda g: pd.Series(
                    {
                        "Income": g[g["Category"] == "Income"]["Amount"].sum(),
                        "Expenses": g[g["_exp"]]["Amount"].sum(),
                    }
                )
            )
            .reset_index()
        )
        lbl = [ml[m - 1] for m in mo["_mo"]]
        ax.plot(lbl, mo["Income"], "o-", color="#2ecc71", label="Income", lw=2)
        ax.plot(lbl, mo["Expenses"], "o-", color="#e74c3c", label="Expenses", lw=2)
        ax.fill_between(
            range(len(mo)), mo["Income"], mo["Expenses"], alpha=0.08, color="#3498db"
        )
        ax.set_title(f"{yr} - Monthly Trend", fontweight="bold")
        ax.set_xticks(range(len(mo)))
        ax.set_xticklabels(lbl, rotation=45, fontsize=8)
        ax.legend(fontsize=8)
        ax.grid(linestyle="--", alpha=0.4)
        _bar_style(ax)
    fig.suptitle("Income vs Expenses - Annual Trend", fontweight="bold")
    fig.tight_layout()
    return save_fig(fig, fname)


def _chart_metric_3m(df, cy, cm, metric_key, title, is_pct, fname):
    """Grouped bar chart: last 3 complete months vs same month LY, plus LY avg line."""
    ml = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ]

    targets = []
    for i in range(3, 0, -1):
        mo, yr = cm - i, cy
        if mo <= 0:
            mo += 12
            yr -= 1
        targets.append((yr, mo))
    targets.append((cy, cm))  # include current report month as 4th bar

    lbls = [
        f"{ml[mo - 1]} '{str(yr)[2:]}" if yr != cy else ml[mo - 1] for yr, mo in targets
    ]
    curr_vals = [summarize(period(df, yr, mo))[metric_key] for yr, mo in targets]
    ly_vals = [summarize(period(df, yr - 1, mo))[metric_key] for yr, mo in targets]

    # LY full-year monthly average
    ly_year = df[df["_yr"] == cy - 1]
    n_ly = ly_year["_mo"].nunique()
    ly_avg = summarize(ly_year, n_ly)[metric_key] if n_ly else None

    # Color scheme by metric type
    _palette = {
        "income":   ("#1a7a1a", "#6dbf6d"),   # dark green / lighter green
        "expenses": ("#c0392b", "#e89090"),   # dark red / lighter red
        "rate":     ("#00aa44", "#66ddaa"),   # bright green / lighter green
        "net":      ("#00aa44", "#66ddaa"),
    }
    col_curr, col_ly = _palette.get(metric_key, ("#2980b9", "#85c1e9"))
    col_avg = "#f39c12"  # orange dashed avg line stays neutral

    x = np.arange(len(targets))
    w = 0.35
    all_vals = curr_vals + ly_vals
    max_v = max((abs(v) for v in all_vals), default=1) or 1
    offset = max_v * 0.025

    fig, ax = plt.subplots(figsize=(9, 4.4))
    bars1 = ax.bar(x - w / 2, curr_vals, w, label=str(cy), color=col_curr)
    bars2 = ax.bar(x + w / 2, ly_vals, w, label=f"Same Month {cy - 1}", color=col_ly)

    for bars in (bars1, bars2):
        for bar in bars:
            h = bar.get_height()
            if h == 0:
                continue
            lbl = f"{h:.1f}%" if is_pct else f"${h:,.0f}"
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                h + (offset if h >= 0 else -offset * 2),
                lbl,
                ha="center",
                va="bottom" if h >= 0 else "top",
                fontsize=7.5,
                fontweight="bold",
            )

    if ly_avg is not None:
        lbl_avg = f"{ly_avg:.1f}%" if is_pct else f"${ly_avg:,.0f}"
        ax.axhline(
            ly_avg,
            color=col_avg,
            ls="--",
            lw=1.5,
            label=f"{cy - 1} Avg ({lbl_avg}/mo)",
        )

    ax.set_xticks(x)
    ax.set_xticklabels(lbls)
    ax.set_title(title, fontweight="bold")
    ax.axhline(0, color="black", lw=0.8)
    ax.legend(fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False)
    _bar_style(ax, currency=not is_pct)
    if is_pct:
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    fig.tight_layout(rect=[0, 0.12, 1, 1])
    return save_fig(fig, fname)


def chart_savings_rate(df, cy, cm, fname):
    """Grouped bars = net savings $ vs same month LY."""
    ml = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

    targets = []
    for i in range(3, 0, -1):
        mo, yr = cm - i, cy
        if mo <= 0:
            mo += 12; yr -= 1
        targets.append((yr, mo))
    targets.append((cy, cm))

    lbls = [f"{ml[mo-1]} '{str(yr)[2:]}" if yr != cy else ml[mo-1] for yr, mo in targets]

    curr_net = [summarize(period(df, yr, mo))["net"]  for yr, mo in targets]
    ly_net   = [summarize(period(df, yr-1, mo))["net"] for yr, mo in targets]

    x = np.arange(len(targets))
    w = 0.35
    max_net = max((abs(v) for v in curr_net + ly_net), default=1) or 1
    offset = max_net * 0.025

    fig, ax = plt.subplots(figsize=(9, 4.4))

    bars1 = ax.bar(x - w/2, curr_net, w, label=f"{cy} Net ($)", color="#00aa44")
    bars2 = ax.bar(x + w/2, ly_net,   w, label=f"Same Month {cy-1} Net ($)", color="#66ddaa")

    for bars, vals in ((bars1, curr_net), (bars2, ly_net)):
        for bar, v in zip(bars, vals):
            if v == 0: continue
            ax.text(bar.get_x() + bar.get_width()/2,
                    v + (offset if v >= 0 else -offset*2),
                    f"${v:,.0f}", ha="center",
                    va="bottom" if v >= 0 else "top",
                    fontsize=7.5, fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels(lbls)
    ax.set_title(f"{cy} Net Savings - Last 4 Months vs Same Month LY", fontweight="bold")
    ax.axhline(0, color="black", lw=0.8)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.legend(fontsize=7.5, loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=2, frameon=False)
    _bar_style(ax)

    fig.tight_layout(rect=[0, 0.1, 1, 1])
    return save_fig(fig, fname)


def chart_income_3m(df, cy, cm, fname):
    return _chart_metric_3m(
        df,
        cy,
        cm,
        "income",
        f"{cy} Income - Last 4 Months vs Same Month LY",
        False,
        fname,
    )


def chart_expense_3m(df, cy, cm, fname):
    return _chart_metric_3m(
        df,
        cy,
        cm,
        "expenses",
        f"{cy} Expenses - Last 4 Months vs Same Month LY",
        False,
        fname,
    )


def chart_splurge(sd, fname):
    names = list(sd.keys())
    used = [sd[n]["used"] for n in names]
    rem = [max(0, SPLURGE_LIMIT - sd[n]["used"]) for n in names]
    x, w = np.arange(len(names)), 0.45

    fig, ax = plt.subplots(figsize=(7, 4))
    b1 = ax.bar(x, used, w, label="Used", color="#e74c3c")
    ax.bar(x, rem, w, bottom=used, label="Remaining", color="#2ecc71", alpha=0.75)
    ax.axhline(
        SPLURGE_LIMIT,
        color="#c0392b",
        ls="--",
        lw=1.5,
        label=f"Limit ${SPLURGE_LIMIT:,}",
    )
    ax.set_xticks(x)
    ax.set_xticklabels(names, fontsize=12)
    ax.set_title("Splurge Budget Tracker (YTD)", fontweight="bold")
    ax.legend()
    _bar_style(ax)
    for bar, v in zip(b1, used):
        if v > 0:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                v / 2,
                f"${v:,.0f}",
                ha="center",
                va="center",
                color="white",
                fontweight="bold",
                fontsize=10,
            )
    fig.tight_layout()
    return save_fig(fig, fname)


def chart_person_cats(cm_df, person, fname, df=None, cy=None, cm=None):
    sub = cm_df  # caller is responsible for pre-filtering/combining
    if sub.empty:
        return None
    exp_cm = sub[sub["_exp"]].groupby("Category")["Amount"].sum()
    exp_cm = exp_cm[exp_cm > 0].sort_values()
    if exp_cm.empty:
        return None

    cats = exp_cm.index.tolist()
    cm_vals = [exp_cm.get(c, 0) for c in cats]

    # Same month last year & LY monthly avg (requires full df)
    has_ly = df is not None and cy is not None and cm is not None
    ly_vals, ly_avg_vals = None, None
    if has_ly:
        ly_month_df = combined_person_df(df[(df["_yr"] == cy - 1) & (df["_mo"] == cm)], person)
        ly_exp = ly_month_df[ly_month_df["_exp"]].groupby("Category")["Amount"].sum() if not ly_month_df.empty else pd.Series(dtype=float)
        ly_vals = [ly_exp.get(c, 0) for c in cats]

        ly_full_df = combined_person_df(df[df["_yr"] == cy - 1], person)
        n_ly = ly_full_df["_mo"].nunique() if not ly_full_df.empty else 0
        if n_ly:
            ly_total = ly_full_df[ly_full_df["_exp"]].groupby("Category")["Amount"].sum()
            ly_avg_vals = [ly_total.get(c, 0) / n_ly for c in cats]

    y = np.arange(len(cats))
    fig_h = max(4, len(cats) * (0.75 if has_ly else 0.5))
    fig, ax = plt.subplots(figsize=(9, fig_h))

    if has_ly and ly_vals is not None and ly_avg_vals is not None:
        w = 0.28
        bars_cm = ax.barh(y + w, cm_vals, w, label="This Month", color="#2980b9")
        bars_ly = ax.barh(y, ly_vals, w, label=f"Same Month {cy - 1}", color="#e74c3c", alpha=0.85)
        bars_avg = ax.barh(y - w, ly_avg_vals, w, label=f"{cy - 1} Monthly Avg", color="#f39c12", alpha=0.85)
        all_bars = [(bars_cm, cm_vals), (bars_ly, ly_vals), (bars_avg, ly_avg_vals)]
        for bars, vals in all_bars:
            for bar, v in zip(bars, vals):
                if v > 0:
                    ax.text(v + max(cm_vals) * 0.01, bar.get_y() + bar.get_height() / 2,
                            f"${v:,.0f}", va="center", fontsize=7)
        ax.set_yticks(y)
        ax.set_yticklabels(cats)
        ax.legend(fontsize=8)
    else:
        colors = sns.color_palette("Set3", len(cats))
        bars = ax.barh(y, cm_vals, color=colors)
        for bar, v in zip(bars, cm_vals):
            ax.text(v + max(cm_vals) * 0.01, bar.get_y() + bar.get_height() / 2,
                    f"${v:,.0f}", va="center", fontsize=8)
        ax.set_yticks(y)
        ax.set_yticklabels(cats)

    ax.set_title(f"{person} - Expense Breakdown (This Month vs LY)", fontweight="bold")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    _bar_style(ax, currency=False)
    fig.tight_layout()
    return save_fig(fig, fname)


def chart_pie(cm_df, title, fname):
    cats = cat_totals(cm_df)
    cats = cats[cats > 0]
    if cats.empty:
        return None
    fig, ax = plt.subplots(figsize=(8, 6))
    _, _, autos = ax.pie(
        cats.values,
        labels=cats.index,
        autopct="%1.1f%%",
        startangle=140,
        pctdistance=0.82,
        colors=sns.color_palette("Set2", len(cats)),
    )
    for t in autos:
        t.set_fontsize(8)
    ax.set_title(title, fontweight="bold")
    fig.tight_layout()
    return save_fig(fig, fname)


def chart_person_splurge_monthly(df, cy, cm, fname):
    """Splurge spend for the last 3 complete months, with prior-year monthly avg."""
    ml = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ]

    # Build last 3 complete months + current report month
    targets = []
    for i in range(3, 0, -1):
        mo, yr = cm - i, cy
        if mo <= 0:
            mo += 12
            yr -= 1
        targets.append((yr, mo))
    targets.append((cy, cm))  # 4th point = current report month

    lbls = [
        f"{ml[mo - 1]} '{str(yr)[2:]}" if yr != cy else ml[mo - 1] for yr, mo in targets
    ]

    fig, ax = plt.subplots(figsize=(10, 4))
    colors = {"Isa": "#9b59b6", "Toio": "#3498db"}

    for p in PERSONS:
        vals = [
            df[
                (df["_yr"] == yr)
                & (df["_mo"] == mo)
                & df["_splurge"]
                & (df["Name"] == p)
            ]["Amount"].sum()
            for yr, mo in targets
        ]
        ax.plot(lbls, vals, "o-", color=colors[p], label=p, lw=2)
        for x, v in enumerate(vals):
            ax.annotate(
                f"${v:,.0f}",
                (x, v),
                textcoords="offset points",
                xytext=(0, 8),
                ha="center",
                fontsize=8,
                color=colors[p],
                fontweight="bold",
            )

        # Prior-year monthly average splurge line
        ly = df[(df["_yr"] == cy - 1) & df["_splurge"] & (df["Name"] == p)]
        if not ly.empty:
            n_mo = ly["_mo"].nunique()
            ly_avg = ly["Amount"].sum() / n_mo if n_mo else 0
            ax.axhline(
                ly_avg,
                color=colors[p],
                ls=":",
                lw=1.5,
                alpha=0.7,
                label=f"{p} {cy - 1} avg (${ly_avg:,.0f}/mo)",
            )

    ax.axhline(
        SPLURGE_LIMIT / 12,
        color="#c0392b",
        ls="--",
        lw=1.5,
        label=f"Monthly limit (${SPLURGE_LIMIT / 12:,.0f})",
    )
    ax.set_title("Splurge Spend - Last 3 Months vs Prior Year Avg", fontweight="bold")
    ax.legend(fontsize=8)
    ax.grid(linestyle="--", alpha=0.4)
    _bar_style(ax)
    fig.tight_layout()
    return save_fig(fig, fname)


# ── PDF PRIMITIVES ────────────────────────────────────────────────────────────
def sec_hdr(pdf, title, r=41, g=128, b=185):
    pdf.set_fill_color(r, g, b)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 7, f"  {title}", fill=True, ln=True)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(1)


def kpi_boxes(pdf, kpis):
    """kpis = list of (label, value_str, subtitle_str)."""
    w = 190 / len(kpis)
    x0 = pdf.get_x()
    y0 = pdf.get_y()
    for i, (lbl, val, sub) in enumerate(kpis):
        x = x0 + i * w
        pdf.set_fill_color(236, 240, 241)
        pdf.rect(x, y0, w - 2, 22, "F")
        pdf.set_xy(x + 2, y0 + 2)
        pdf.set_font("Arial", "", 7)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(w - 4, 4, lbl)
        pdf.set_xy(x + 2, y0 + 7)
        pdf.set_font("Arial", "B", 11)
        pdf.set_text_color(44, 62, 80)
        pdf.cell(w - 4, 7, str(val))
        pdf.set_xy(x + 2, y0 + 15)
        pdf.set_font("Arial", "I", 7)
        pdf.set_text_color(120, 120, 120)
        pdf.cell(w - 4, 5, str(sub))
    pdf.set_xy(x0, y0 + 24)
    pdf.set_text_color(0, 0, 0)


def tbl(pdf, headers, rows, widths, aligns=None):
    if aligns is None:
        aligns = ["L"] + ["C"] * (len(headers) - 1)
    # header
    pdf.set_fill_color(41, 128, 185)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 8.5)
    for i, h in enumerate(headers):
        pdf.cell(widths[i], 7, str(h), fill=True, align=aligns[i])
    pdf.ln()
    # rows
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "", 8)
    for ri, row in enumerate(rows):
        pdf.set_fill_color(245, 248, 250) if ri % 2 == 0 else pdf.set_fill_color(
            255, 255, 255
        )
        for i, cell in enumerate(row):
            pdf.cell(widths[i], 6.5, str(cell), fill=True, align=aligns[i])
        pdf.ln()
    pdf.ln(3)


# ── PDF LAYOUT HELPERS ────────────────────────────────────────────────────────
_TOP_MARGIN = 18  # must match header set_y()


def _section_gap(pdf):
    """Thin grey rule + small vertical gap between sections (no forced page break).
    Skips the leading gap when at the top of the page to keep all pages consistent."""
    if pdf.get_y() > _TOP_MARGIN + 4:
        pdf.ln(4)
    y = pdf.get_y()
    pdf.set_draw_color(210, 210, 210)
    pdf.line(10, y, 200, y)
    pdf.set_draw_color(0, 0, 0)
    pdf.ln(4)


def _check_space(pdf, needed_mm):
    """Force a page break only when the remaining vertical space is insufficient."""
    remaining = pdf.h - pdf.get_y() - pdf.b_margin
    if remaining < needed_mm:
        pdf.add_page()


# ── PAGE BUILDERS ─────────────────────────────────────────────────────────────
def page_summary(pdf, cm_s, smly_s, avg_s, n90, lbl_cm, lbl_ly):
    _section_gap(pdf)
    _check_space(pdf, 25)
    sec_hdr(pdf, "Monthly Performance Summary")

    rows = []
    for lbl, key, is_pct in [
        ("Income", "income", False),
        ("Expenses", "expenses", False),
        ("Net Savings", "net", False),
        ("Savings Rate", "rate", True),
    ]:
        f = fp if is_pct else fc
        rows.append(
            [
                lbl,
                f(cm_s[key]),
                f(smly_s[key]),
                delta(cm_s[key], smly_s[key]),
                f(avg_s[key]),
                delta(cm_s[key], avg_s[key]),
            ]
        )

    tbl(
        pdf,
        ["Metric", "This Month", lbl_ly, "vs LY", "90-Day Avg", "vs 90d"],
        rows,
        [42, 32, 32, 22, 32, 22],
        ["L", "R", "R", "C", "R", "C"],
    )

    pdf.set_font("Arial", "I", 7)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(
        0,
        5,
        f"* 90-day avg based on {n90} complete month(s) preceding this month",
        ln=True,
    )
    pdf.set_text_color(0, 0, 0)


def page_categories(pdf, cm_df, smly_df, avg_df, n90, cat_chart, pie_chart):
    _section_gap(pdf)
    _check_space(pdf, 25)
    sec_hdr(pdf, "Category Breakdown - Expenses Comparison")

    cm_c = cat_totals(cm_df)
    sm_c = cat_totals(smly_df)
    av_c = (
        cat_totals(avg_df, n90)
        if n90
        else pd.Series({c: 0.0 for c in EXPENSE_CATEGORIES})
    )

    rows = []
    for cat in EXPENSE_CATEGORIES:
        cv, sv, av = cm_c.get(cat, 0), sm_c.get(cat, 0), av_c.get(cat, 0)
        rows.append(
            [
                cat,
                fc(cv),
                fc(sv),
                delta(cv, sv) if (cv or sv) else "-",
                fc(av),
                delta(cv, av) if (cv or av) else "-",
            ]
        )

    rows.append(
        [
            "TOTAL",
            fc(cm_c.sum()),
            fc(sm_c.sum()),
            delta(cm_c.sum(), sm_c.sum()),
            fc(av_c.sum()),
            delta(cm_c.sum(), av_c.sum()),
        ]
    )

    tbl(
        pdf,
        ["Category", "This Month", "Same Month LY", "vs LY", "90-Day Avg", "vs 90d"],
        rows,
        [42, 30, 30, 22, 30, 22],
        ["L", "R", "R", "C", "R", "C"],
    )

    if cat_chart and os.path.exists(cat_chart):
        _check_space(pdf, 70)
        pdf.image(cat_chart, x=10, w=190)

    if pie_chart and os.path.exists(pie_chart):
        _section_gap(pdf)
        _check_space(pdf, 145)
        sec_hdr(pdf, "Expense Distribution")
        pdf.image(pie_chart, x=30, w=150)


def page_person(pdf, cm_df, smly_df, avg_df, n90, person, chart_path):
    colour = {"Isa": (155, 89, 182), "Toio": (52, 152, 219)}
    r, g, b = colour.get(person, (41, 128, 185))

    # Person's own transactions + 50% of Joint (shared couple) transactions
    cm_combined = combined_person_df(cm_df, person)
    smly_combined = combined_person_df(smly_df, person)
    avg_combined = combined_person_df(avg_df, person)

    cm_s = summarize(cm_combined)
    smly_s = summarize(smly_combined)
    avg_s = summarize(avg_combined, n90) if n90 else summarize(avg_combined)

    _section_gap(pdf)
    _check_space(pdf, 25)
    sec_hdr(pdf, f"{person} - Monthly Summary", r, g, b)

    _check_space(pdf, 30)
    kpi_boxes(
        pdf,
        [
            ("Income", fc(cm_s["income"]), "Personal"),
            ("Expenses", fc(cm_s["expenses"]), "Own + 50% Shared"),
            ("Net Savings", fc(cm_s["net"]), ""),
            ("Savings Rate", fp(cm_s["rate"]), ""),
        ],
    )

    rows = []
    for lbl, key, is_pct in [
        ("Income", "income", False),
        ("Expenses", "expenses", False),
        ("Net Savings", "net", False),
        ("Savings Rate", "rate", True),
    ]:
        f = fp if is_pct else fc
        rows.append(
            [
                lbl,
                f(cm_s[key]),
                f(smly_s[key]),
                delta(cm_s[key], smly_s[key]),
                f(avg_s[key]),
                delta(cm_s[key], avg_s[key]),
            ]
        )

    tbl(
        pdf,
        ["Metric", "This Month", "Same Month LY", "vs LY", "90-Day Avg", "vs 90d"],
        rows,
        [42, 32, 32, 22, 32, 22],
        ["L", "R", "R", "C", "R", "C"],
    )

    # per-category breakdown: person's own + 50% of Joint expenses
    _check_space(pdf, 25)
    sec_hdr(pdf, f"{person} - Category Breakdown", r, g, b)
    cm_c = cat_totals(cm_combined)
    sm_c = cat_totals(smly_combined)
    av_c = cat_totals(avg_combined, n90) if n90 else cat_totals(avg_combined)

    cat_rows = []
    for cat in EXPENSE_CATEGORIES:
        cv, sv, av = cm_c.get(cat, 0), sm_c.get(cat, 0), av_c.get(cat, 0)
        if cv == 0 and sv == 0 and av == 0:
            continue
        cat_rows.append(
            [
                cat,
                fc(cv),
                fc(sv),
                delta(cv, sv) if (cv or sv) else "-",
                fc(av),
                delta(cv, av) if (cv or av) else "-",
            ]
        )

    if cat_rows:
        tbl(
            pdf,
            [
                "Category",
                "This Month",
                "Same Month LY",
                "vs LY",
                "90-Day Avg",
                "vs 90d",
            ],
            cat_rows,
            [42, 30, 30, 22, 30, 22],
            ["L", "R", "R", "C", "R", "C"],
        )

    if chart_path and os.path.exists(chart_path):
        _check_space(pdf, 80)
        pdf.image(chart_path, x=10, w=190)


def page_couple_combined(pdf, cm_df, smly_df, avg_df, n90, lbl_cm, lbl_ly):
    """Household total: Isa + Toio + Joint combined."""
    _section_gap(pdf)
    _check_space(pdf, 25)
    sec_hdr(pdf, "Combined - Household Total", 39, 174, 96)

    cm_s = summarize(cm_df)
    smly_s = summarize(smly_df)
    avg_s = summarize(avg_df, n90) if n90 else summarize(avg_df)

    _check_space(pdf, 30)
    kpi_boxes(
        pdf,
        [
            ("Household Income", fc(cm_s["income"]), "This Month"),
            ("Household Expenses", fc(cm_s["expenses"]), "This Month"),
            ("Net Savings", fc(cm_s["net"]), "This Month"),
            ("Savings Rate", fp(cm_s["rate"]), "This Month"),
        ],
    )

    rows = []
    for lbl, key, is_pct in [
        ("Income", "income", False),
        ("Expenses", "expenses", False),
        ("Net Savings", "net", False),
        ("Savings Rate", "rate", True),
    ]:
        f = fp if is_pct else fc
        rows.append(
            [
                lbl,
                f(cm_s[key]),
                f(smly_s[key]),
                delta(cm_s[key], smly_s[key]),
                f(avg_s[key]),
                delta(cm_s[key], avg_s[key]),
            ]
        )

    tbl(
        pdf,
        ["Metric", "This Month", lbl_ly, "vs LY", "90-Day Avg", "vs 90d"],
        rows,
        [42, 32, 32, 22, 32, 22],
        ["L", "R", "R", "C", "R", "C"],
    )

    # category breakdown for the whole household
    _check_space(pdf, 25)
    sec_hdr(pdf, "Household Total - Category Breakdown", 39, 174, 96)
    cm_c = cat_totals(cm_df)
    sm_c = cat_totals(smly_df)
    av_c = (
        cat_totals(avg_df, n90)
        if n90
        else pd.Series({c: 0.0 for c in EXPENSE_CATEGORIES})
    )

    cat_rows = []
    for cat in EXPENSE_CATEGORIES:
        cv, sv, av = cm_c.get(cat, 0), sm_c.get(cat, 0), av_c.get(cat, 0)
        if cv == 0 and sv == 0 and av == 0:
            continue
        cat_rows.append(
            [
                cat,
                fc(cv),
                fc(sv),
                delta(cv, sv) if (cv or sv) else "-",
                fc(av),
                delta(cv, av) if (cv or av) else "-",
            ]
        )

    if cat_rows:
        tbl(
            pdf,
            ["Category", "This Month", lbl_ly, "vs LY", "90-Day Avg", "vs 90d"],
            cat_rows,
            [42, 30, 30, 22, 30, 22],
            ["L", "R", "R", "C", "R", "C"],
        )


def page_splurge(pdf, df, cy, cm, splurge_chart, monthly_chart):
    _section_gap(pdf)
    _check_space(pdf, 25)
    sec_hdr(pdf, "Splurge Budget Tracker", 231, 76, 60)

    ytd = df[(df["_yr"] == cy) & (df["_mo"] <= cm) & df["_splurge"]]
    cm_s = df[(df["_yr"] == cy) & (df["_mo"] == cm) & df["_splurge"]]

    rows = []
    sd = {}
    for p in PERSONS:
        ytd_v = ytd[ytd["Name"] == p]["Amount"].sum()
        cm_v = cm_s[cm_s["Name"] == p]["Amount"].sum()
        rem = max(0.0, SPLURGE_LIMIT - ytd_v)
        avg = ytd_v / cm if cm else 0.0
        pct = ytd_v / SPLURGE_LIMIT * 100
        sd[p] = {"used": ytd_v}
        rows.append([p, fc(cm_v), fc(ytd_v), fc(avg), fc(rem), f"{pct:.1f}%"])

    tbl(
        pdf,
        ["Person", "This Month", "YTD Total", "Monthly Avg", "Remaining", "% Used"],
        rows,
        [35, 32, 32, 32, 32, 22],
        ["L", "R", "R", "R", "R", "C"],
    )

    pdf.set_font("Arial", "I", 7)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(
        0, 5, f"Annual limit: ${SPLURGE_LIMIT:,}/person · Months elapsed: {cm}", ln=True
    )
    pdf.set_text_color(0, 0, 0)
    pdf.ln(4)

    if splurge_chart and os.path.exists(splurge_chart):
        _check_space(pdf, 90)
        pdf.image(splurge_chart, x=30, w=150)

    if monthly_chart and os.path.exists(monthly_chart):
        _check_space(pdf, 75)
        pdf.image(monthly_chart, x=10, w=190)


def page_big_picture(
    pdf, df, cy, cm, trend_chart, income_chart, expense_chart, rate_chart
):
    _section_gap(pdf)
    _check_space(pdf, 25)
    sec_hdr(pdf, f"Big Picture - YTD {cy} vs {cy - 1}")

    ml = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ]
    rows = []
    for m in range(1, cm + 1):
        c = summarize(df[(df["_yr"] == cy) & (df["_mo"] == m)])
        p = summarize(df[(df["_yr"] == cy - 1) & (df["_mo"] == m)])
        rows.append(
            [
                ml[m - 1],
                fc(c["income"]),
                fc(c["expenses"]),
                fc(c["net"]),
                fc(p["income"]),
                fc(p["expenses"]),
                fc(p["net"]),
            ]
        )

    # YTD totals row
    c = summarize(df[(df["_yr"] == cy) & (df["_mo"] <= cm)])
    p = summarize(df[(df["_yr"] == cy - 1) & (df["_mo"] <= cm)])
    rows.append(
        [
            "YTD",
            fc(c["income"]),
            fc(c["expenses"]),
            fc(c["net"]),
            fc(p["income"]),
            fc(p["expenses"]),
            fc(p["net"]),
        ]
    )

    tbl(
        pdf,
        [
            "Month",
            f"{cy} Income",
            f"{cy} Exp",
            f"{cy} Net",
            f"{cy - 1} Income",
            f"{cy - 1} Exp",
            f"{cy - 1} Net",
        ],
        rows,
        [18, 28, 28, 28, 28, 28, 28],
        ["L", "R", "R", "R", "R", "R", "R"],
    )

    if trend_chart and os.path.exists(trend_chart):
        _check_space(pdf, 75)
        pdf.image(trend_chart, x=10, w=190)

    if income_chart and os.path.exists(income_chart):
        _section_gap(pdf)
        _check_space(pdf, 25)
        sec_hdr(pdf, f"{cy} Monthly Income - Last 3 Months vs Same Month LY")
        _check_space(pdf, 75)
        pdf.image(income_chart, x=10, w=190)

    if expense_chart and os.path.exists(expense_chart):
        _section_gap(pdf)
        _check_space(pdf, 25)
        sec_hdr(pdf, f"{cy} Monthly Expenses - Last 3 Months vs Same Month LY")
        _check_space(pdf, 75)
        pdf.image(expense_chart, x=10, w=190)

    if rate_chart and os.path.exists(rate_chart):
        _section_gap(pdf)
        _check_space(pdf, 25)
        sec_hdr(pdf, f"{cy} Monthly Savings Rate - Last 3 Months vs Same Month LY")
        _check_space(pdf, 75)
        pdf.image(rate_chart, x=10, w=190)


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="Couple Financial Report")
    ap.add_argument("--csv", default=os.path.join(SCRIPT_DIR, CSV_FILENAME))
    ap.add_argument("--month", type=int)
    ap.add_argument("--year", type=int)
    ap.add_argument("--out", default=os.path.join(SCRIPT_DIR, PDF_FILENAME))
    args = ap.parse_args()

    print(f"Loading: {args.csv}")
    if not os.path.exists(args.csv):
        print(f"ERROR: File not found: {args.csv}")
        sys.exit(1)

    df = load(args.csv)

    # Default to the most recent month that has income data, capped at last
    # complete calendar month (i.e. never pick the current in-progress month).
    today = datetime.today()
    last_complete = today.replace(day=1) - __import__("datetime").timedelta(days=1)
    cap_yr, cap_mo = last_complete.year, last_complete.month

    income_months = (
        df[df["Category"] == "Income"][["_yr", "_mo"]]
        .drop_duplicates()
        .apply(lambda r: (int(r["_yr"]), int(r["_mo"])), axis=1)
        .tolist()
    )
    # Keep only months up to the last complete month
    income_months = [(y, m) for y, m in income_months if (y, m) <= (cap_yr, cap_mo)]

    if income_months:
        default_yr, default_mo = max(income_months)
    else:
        default_yr, default_mo = cap_yr, cap_mo

    cy = args.year or default_yr
    cm = args.month or default_mo

    lbl = f"{MONTH_NAMES[cm - 1]} {cy}"
    lbl_ly = f"{MONTH_NAMES[cm - 1]} {cy - 1}"
    print(f"Report period: {lbl}")

    cm_df = period(df, cy, cm)
    smly_df = period(df, cy - 1, cm)
    avg_df, n90 = prior_3m(df, cy, cm)

    # ── Charts ──
    print("Generating charts...")
    trend_c = chart_trend(df, cy, cm, "_trend.png")
    rate_c = chart_savings_rate(df, cy, cm, "_rate.png")
    income_c = chart_income_3m(df, cy, cm, "_income.png")
    expense_c = chart_expense_3m(df, cy, cm, "_expense.png")

    ytd_sp = df[(df["_yr"] == cy) & (df["_mo"] <= cm) & df["_splurge"]]
    sd = {p: {"used": ytd_sp[ytd_sp["Name"] == p]["Amount"].sum()} for p in PERSONS}
    sp_c = chart_splurge(sd, "_splurge.png")
    spm_c = chart_person_splurge_monthly(df, cy, cm, "_splurge_monthly.png")

    isa_c = chart_person_cats(combined_person_df(cm_df, "Isa"), "Isa", "_isa.png", df, cy, cm)
    toio_c = chart_person_cats(combined_person_df(cm_df, "Toio"), "Toio", "_toio.png", df, cy, cm)

    # ── PDF ──
    print("Building PDF...")
    pdf = PDF(lbl)
    pdf.set_auto_page_break(auto=True, margin=15)

    pdf.add_page()
    page_couple_combined(pdf, cm_df, smly_df, avg_df, n90, lbl, lbl_ly)
    page_person(pdf, cm_df, smly_df, avg_df, n90, "Isa", isa_c)
    page_person(pdf, cm_df, smly_df, avg_df, n90, "Toio", toio_c)
    page_splurge(pdf, df, cy, cm, sp_c, spm_c)
    page_big_picture(pdf, df, cy, cm, trend_c, income_c, expense_c, rate_c)

    pdf.output(args.out)
    print(f"PDF saved: {args.out}")
    cleanup()
    print("Done!")


if __name__ == "__main__":
    main()

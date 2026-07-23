#!/usr/bin/env python3
"""
FINC 310 - Procter & Gamble Financial Story Dashboard
======================================================

REM - REASONING SUMMARY
This program converts audited Form 10-K financial statement data into four
complementary visualizations. It does not infer or invent financial values.
All source values are read from the standardized FINC 310 workbook.

REM - EQUATIONS
1. Gross Profit = Net Sales - Cost of Products Sold
2. NOPAT = EBIT x (1 - Effective Tax Rate)
3. Operating NOWC = Accounts Receivable + Inventory + Prepaids
                    - Accounts Payable - Accrued Operating Liabilities
4. Textbook FCF = NOPAT + D&A - Capital Expenditures - Change in NOWC
5. Cash FCF = Cash Flow from Operations - Capital Expenditures
6. Profit Margin = Net Earnings / Net Sales
7. Total Asset Turnover = Net Sales / Average Total Assets
8. Equity Multiplier = Average Total Assets / Average Total Equity
9. ROE = Profit Margin x Total Asset Turnover x Equity Multiplier

REM - METHOD
The program:
- reads the 01_Input worksheet;
- builds a dictionary using Metric_ID values;
- performs accounting and finance sanity checks;
- creates the Sankey, FCF waterfall, five-year trend, and DuPont diagram;
- creates a one-page dashboard; and
- writes a CSV audit report.

REM - LIMITATIONS
- The textbook FCF calculation is a simplified unlevered cash-flow measure.
- P&G's adjusted FCF is a company-defined non-GAAP measure.
- The bridge between textbook FCF and cash FCF is a reconciliation residual
  that captures operating accruals, noncash items, and timing items not fully
  represented in the simplified NOWC definition.
- The DuPont analysis uses consolidated net earnings and average total
  shareholders' equity for classroom consistency.
"""

from __future__ import annotations

import argparse
import math
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, Tuple

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.path import Path as MplPath
from matplotlib.patches import PathPatch, Rectangle
from matplotlib.ticker import FuncFormatter
from PIL import Image, ImageDraw, ImageFont


def load_metrics(workbook_path: Path) -> Dict[str, Dict[int, float]]:
    """Read machine-readable metrics from the standardized 01_Input sheet."""
    df = pd.read_excel(workbook_path, sheet_name="01_Input", header=4)
    required = {"Metric_ID", "Metric", "Unit", "FY2025", "FY2024", "FY2023", "FY2022", "FY2021"}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Workbook is missing expected columns: {sorted(missing)}")

    df = df[df["Metric_ID"].notna()].copy()
    metrics: Dict[str, Dict[int, float]] = {}
    for _, row in df.iterrows():
        metric_id = str(row["Metric_ID"]).strip()
        metrics[metric_id] = {}
        for year in (2025, 2024, 2023, 2022, 2021):
            value = row.get(f"FY{year}")
            if pd.notna(value):
                metrics[metric_id][year] = float(value)
    return metrics


def value(metrics: Dict[str, Dict[int, float]], metric_id: str, year: int) -> float:
    try:
        return metrics[metric_id][year]
    except KeyError as exc:
        raise KeyError(f"Missing {metric_id} for FY{year} in 01_Input") from exc


def close_enough(actual: float, expected: float, tolerance: float = 1.0) -> bool:
    return math.isclose(actual, expected, abs_tol=tolerance)


def build_model(metrics: Dict[str, Dict[int, float]]) -> Dict[str, float]:
    """Calculate the finance metrics used by all four visuals."""
    model: Dict[str, float] = {}

    model["sales_2025"] = value(metrics, "NET_SALES", 2025)
    model["cogs_2025"] = value(metrics, "COGS", 2025)
    model["gross_profit_2025"] = model["sales_2025"] - model["cogs_2025"]
    model["sga_2025"] = value(metrics, "SGA", 2025)
    model["impairment_2025"] = value(metrics, "IMPAIRMENT", 2025)
    model["ebit_2025"] = value(metrics, "OPERATING_INCOME", 2025)
    model["interest_expense_2025"] = abs(value(metrics, "INTEREST_EXPENSE", 2025))
    model["interest_income_2025"] = value(metrics, "INTEREST_INCOME", 2025)
    model["other_nonop_2025"] = value(metrics, "OTHER_NONOP", 2025)
    model["net_nonop_expense_2025"] = (
        model["interest_expense_2025"]
        - model["interest_income_2025"]
        - model["other_nonop_2025"]
    )
    model["ebt_2025"] = value(metrics, "EBT", 2025)
    model["taxes_2025"] = value(metrics, "INCOME_TAXES", 2025)
    model["net_earnings_2025"] = value(metrics, "NET_EARNINGS", 2025)
    model["net_income_pg_2025"] = value(metrics, "NET_INCOME_PG", 2025)

    model["effective_tax_rate"] = model["taxes_2025"] / model["ebt_2025"]
    model["nopat"] = model["ebit_2025"] * (1 - model["effective_tax_rate"])

    model["da_2025"] = value(metrics, "D_AND_A", 2025)
    model["capex_2025"] = abs(value(metrics, "CAPEX", 2025))
    model["cfo_2025"] = value(metrics, "CFO", 2025)
    model["tax_act_payment_2025"] = value(metrics, "TAX_ACT_PAYMENT", 2025)
    model["adjusted_fcf_2025"] = value(metrics, "ADJUSTED_FCF", 2025)

    nowc_2025 = (
        value(metrics, "AR", 2025)
        + value(metrics, "INVENTORY", 2025)
        + value(metrics, "PREPAIDS", 2025)
        - value(metrics, "AP", 2025)
        - value(metrics, "ACCRUED_LIABILITIES", 2025)
    )
    nowc_2024 = (
        value(metrics, "AR", 2024)
        + value(metrics, "INVENTORY", 2024)
        + value(metrics, "PREPAIDS", 2024)
        - value(metrics, "AP", 2024)
        - value(metrics, "ACCRUED_LIABILITIES", 2024)
    )
    model["nowc_2025"] = nowc_2025
    model["nowc_2024"] = nowc_2024
    model["delta_nowc"] = nowc_2025 - nowc_2024
    model["textbook_fcf"] = (
        model["nopat"]
        + model["da_2025"]
        - model["capex_2025"]
        - model["delta_nowc"]
    )
    model["cash_fcf"] = model["cfo_2025"] - model["capex_2025"]
    model["other_operating_bridge"] = model["cash_fcf"] - model["textbook_fcf"]

    model["avg_assets"] = (
        value(metrics, "TOTAL_ASSETS", 2025) + value(metrics, "TOTAL_ASSETS", 2024)
    ) / 2
    model["avg_equity"] = (
        value(metrics, "TOTAL_EQUITY", 2025) + value(metrics, "TOTAL_EQUITY", 2024)
    ) / 2
    model["profit_margin"] = model["net_earnings_2025"] / model["sales_2025"]
    model["asset_turnover"] = model["sales_2025"] / model["avg_assets"]
    model["equity_multiplier"] = model["avg_assets"] / model["avg_equity"]
    model["roe"] = (
        model["profit_margin"]
        * model["asset_turnover"]
        * model["equity_multiplier"]
    )

    model["years"] = [2021, 2022, 2023, 2024, 2025]
    model["sales_history"] = [value(metrics, "NET_SALES", y) for y in model["years"]]
    model["eps_history"] = [value(metrics, "DILUTED_EPS", y) for y in model["years"]]
    model["sales_cagr"] = (
        model["sales_history"][-1] / model["sales_history"][0]
    ) ** (1 / 4) - 1
    model["eps_cagr"] = (
        model["eps_history"][-1] / model["eps_history"][0]
    ) ** (1 / 4) - 1
    return model


def sanity_checks(model: Dict[str, float]) -> pd.DataFrame:
    """Return an auditable list of accounting and finance checks."""
    checks = [
        (
            "Revenue identity",
            model["gross_profit_2025"] + model["cogs_2025"],
            model["sales_2025"],
            "$mm",
        ),
        (
            "Gross profit identity",
            model["sga_2025"] + model["impairment_2025"] + model["ebit_2025"],
            model["gross_profit_2025"],
            "$mm",
        ),
        (
            "Pretax income identity",
            model["ebit_2025"] - model["net_nonop_expense_2025"],
            model["ebt_2025"],
            "$mm",
        ),
        (
            "Net earnings identity",
            model["ebt_2025"] - model["taxes_2025"],
            model["net_earnings_2025"],
            "$mm",
        ),
        (
            "Adjusted FCF identity",
            model["cash_fcf"] + model["tax_act_payment_2025"],
            model["adjusted_fcf_2025"],
            "$mm",
        ),
        (
            "DuPont ROE identity",
            model["roe"],
            model["net_earnings_2025"] / model["avg_equity"],
            "ratio",
        ),
    ]
    rows = []
    for name, actual, expected, unit in checks:
        tolerance = 0.0001 if unit == "ratio" else 1.0
        rows.append(
            {
                "Check": name,
                "Calculated": actual,
                "Expected": expected,
                "Difference": actual - expected,
                "Status": "PASS" if close_enough(actual, expected, tolerance) else "REVIEW",
            }
        )
    return pd.DataFrame(rows)



def fmt_whole(value: float) -> str:
    """Format a displayed whole number using conventional half-up rounding."""
    rounded = Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return f"{int(rounded):,}"


def fmt_b(value_mm: float) -> str:
    return f"${value_mm / 1000:,.1f}B"


def save_sankey(model: Dict[str, float], path: Path) -> None:
    """Create a multi-stage, width-proportional income-statement Sankey."""
    total = model["sales_2025"]
    scale = 0.78 / total
    x_positions = [0.04, 0.23, 0.43, 0.63, 0.83]
    node_w = 0.027

    nodes = {}
    nodes["Net Sales"] = (x_positions[0], 0.11, 0.89)

    cogs_h = model["cogs_2025"] * scale
    gp_h = model["gross_profit_2025"] * scale
    nodes["Cost of Products Sold"] = (x_positions[1], 0.11, 0.11 + cogs_h)
    nodes["Gross Profit"] = (x_positions[1], 0.89 - gp_h, 0.89)

    gp_y0, gp_y1 = nodes["Gross Profit"][1], nodes["Gross Profit"][2]
    sga_h = model["sga_2025"] * scale
    op_h = model["ebit_2025"] * scale
    nodes["SG&A"] = (x_positions[2], gp_y0, gp_y0 + sga_h)
    nodes["Operating Income"] = (x_positions[2], gp_y1 - op_h, gp_y1)

    op_y0, op_y1 = nodes["Operating Income"][1], nodes["Operating Income"][2]
    nonop_h = max(model["net_nonop_expense_2025"] * scale, 0.004)
    ebt_h = model["ebt_2025"] * scale
    nodes["Net Non-Operating Expense"] = (x_positions[3], op_y0, op_y0 + nonop_h)
    nodes["Earnings Before Income Taxes"] = (x_positions[3], op_y1 - ebt_h, op_y1)

    ebt_y0, ebt_y1 = (
        nodes["Earnings Before Income Taxes"][1],
        nodes["Earnings Before Income Taxes"][2],
    )
    tax_h = model["taxes_2025"] * scale
    ni_h = model["net_earnings_2025"] * scale
    nodes["Income Taxes"] = (x_positions[4], ebt_y0, ebt_y0 + tax_h)
    nodes["Net Earnings"] = (x_positions[4], ebt_y1 - ni_h, ebt_y1)

    flows = [
        ("Net Sales", "Cost of Products Sold", model["cogs_2025"]),
        ("Net Sales", "Gross Profit", model["gross_profit_2025"]),
        ("Gross Profit", "SG&A", model["sga_2025"]),
        ("Gross Profit", "Operating Income", model["ebit_2025"]),
        (
            "Operating Income",
            "Net Non-Operating Expense",
            model["net_nonop_expense_2025"],
        ),
        ("Operating Income", "Earnings Before Income Taxes", model["ebt_2025"]),
        ("Earnings Before Income Taxes", "Income Taxes", model["taxes_2025"]),
        ("Earnings Before Income Taxes", "Net Earnings", model["net_earnings_2025"]),
    ]

    fig, ax = plt.subplots(figsize=(18, 10))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.set_title(
        "Procter & Gamble FY2025 Income Statement Sankey\n"
        "Where Each Sales Dollar Went",
        fontsize=22,
        weight="bold",
        pad=18,
    )

    out_cursor = {key: node[1] for key, node in nodes.items()}
    in_cursor = {key: node[1] for key, node in nodes.items()}
    colors = plt.rcParams["axes.prop_cycle"].by_key()["color"]
    node_color = {key: colors[i % len(colors)] for i, key in enumerate(nodes)}

    def ribbon(
        x0: float,
        y0a: float,
        y0b: float,
        x1: float,
        y1a: float,
        y1b: float,
        color: str,
    ) -> None:
        vertices = [
            (x0, y0a),
            (x0 + (x1 - x0) * 0.45, y0a),
            (x1 - (x1 - x0) * 0.45, y1a),
            (x1, y1a),
            (x1, y1b),
            (x1 - (x1 - x0) * 0.45, y1b),
            (x0 + (x1 - x0) * 0.45, y0b),
            (x0, y0b),
            (x0, y0a),
        ]
        codes = [
            MplPath.MOVETO,
            MplPath.CURVE4,
            MplPath.CURVE4,
            MplPath.CURVE4,
            MplPath.LINETO,
            MplPath.CURVE4,
            MplPath.CURVE4,
            MplPath.CURVE4,
            MplPath.CLOSEPOLY,
        ]
        ax.add_patch(
            PathPatch(
                MplPath(vertices, codes),
                facecolor=color,
                edgecolor="none",
                alpha=0.45,
            )
        )

    for source, target, amount in flows:
        visible_height = max(amount * scale, 0.0015)
        sx, sy0, sy1 = nodes[source]
        tx, ty0, ty1 = nodes[target]
        source_start = out_cursor[source]
        source_end = min(source_start + visible_height, sy1)
        target_start = in_cursor[target]
        target_end = min(target_start + visible_height, ty1)
        ribbon(
            sx + node_w,
            source_start,
            source_end,
            tx,
            target_start,
            target_end,
            node_color[target],
        )
        out_cursor[source] = source_end
        in_cursor[target] = target_end

    labels = {
        "Net Sales": ("Net Sales", model["sales_2025"]),
        "Cost of Products Sold": ("Cost of Products Sold", model["cogs_2025"]),
        "Gross Profit": ("Gross Profit", model["gross_profit_2025"]),
        "SG&A": ("SG&A", model["sga_2025"]),
        "Operating Income": ("Operating Income (EBIT)", model["ebit_2025"]),
        "Net Non-Operating Expense": (
            "Net Non-Operating Expense",
            model["net_nonop_expense_2025"],
        ),
        "Earnings Before Income Taxes": (
            "Earnings Before Taxes",
            model["ebt_2025"],
        ),
        "Income Taxes": ("Income Taxes", model["taxes_2025"]),
        "Net Earnings": ("Net Earnings", model["net_earnings_2025"]),
    }

    bold_nodes = {"Net Sales", "Gross Profit", "Operating Income", "Net Earnings"}
    for key, (x, y0, y1) in nodes.items():
        height = y1 - y0
        ax.add_patch(
            Rectangle(
                (x, y0),
                node_w,
                height,
                facecolor=node_color[key],
                edgecolor="black",
                linewidth=0.8,
            )
        )
        title, amount = labels[key]
        percent = amount / model["sales_2025"]
        y_center = (y0 + y1) / 2
        x_text, align = (x - 0.01, "right") if x < 0.1 else (x + node_w + 0.008, "left")
        if height < 0.018:
            y_center = y1 + 0.015
        ax.text(
            x_text,
            y_center,
            f"{title}\n{fmt_b(amount)} | {percent:.1%} of sales",
            va="center",
            ha=align,
            fontsize=11,
            weight="bold" if key in bold_nodes else "normal",
        )

    ax.text(
        0.5,
        0.025,
        "Source: Procter & Gamble FY2025 Form 10-K, "
        "Consolidated Statements of Earnings. Amounts in USD millions.",
        ha="center",
        va="bottom",
        fontsize=10,
    )
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_waterfall(model: Dict[str, float], path: Path) -> None:
    labels = [
        "EBIT",
        "Tax on EBIT",
        "NOPAT",
        "D&A",
        "Capital\nExpenditures",
        "Increase in\nNOWC",
        "Textbook FCF",
        "Other Operating\nAdjustments",
        "Cash FCF\n(CFO - CapEx)",
        "2017 Tax Act\nPayment Add-back",
        "P&G Adjusted FCF",
    ]
    kinds = [
        "total",
        "change",
        "total",
        "change",
        "change",
        "change",
        "total",
        "change",
        "total",
        "change",
        "total",
    ]
    values = [
        model["ebit_2025"],
        -model["ebit_2025"] * model["effective_tax_rate"],
        model["nopat"],
        model["da_2025"],
        -model["capex_2025"],
        -model["delta_nowc"],
        model["textbook_fcf"],
        model["other_operating_bridge"],
        model["cash_fcf"],
        model["tax_act_payment_2025"],
        model["adjusted_fcf_2025"],
    ]

    starts, heights = [], []
    cumulative = 0.0
    for kind, amount in zip(kinds, values):
        if kind == "total":
            starts.append(0)
            heights.append(amount)
            cumulative = amount
        else:
            starts.append(cumulative if amount >= 0 else cumulative + amount)
            heights.append(abs(amount))
            cumulative += amount

    colors = []
    for kind, amount in zip(kinds, values):
        colors.append("C0" if kind == "total" else ("C2" if amount >= 0 else "C3"))

    fig, ax = plt.subplots(figsize=(18, 10))
    bars = ax.bar(
        range(len(labels)),
        heights,
        bottom=starts,
        color=colors,
        edgecolor="black",
        linewidth=0.7,
    )

    cumulative = 0.0
    for index, (kind, amount) in enumerate(zip(kinds, values)):
        cumulative = amount if kind == "total" else cumulative + amount
        if index < len(labels) - 1:
            ax.plot(
                [index + 0.4, index + 0.6],
                [cumulative, cumulative],
                color="gray",
                linewidth=1,
            )

    for index, (kind, amount, bar) in enumerate(zip(kinds, values, bars)):
        y = bar.get_y() + bar.get_height() + 350 if kind == "total" else bar.get_y() + bar.get_height() / 2
        label = (
            f"{amount / 1000:,.1f}B"
            if abs(amount) >= 1000
            else f"{amount:,.0f}M"
        )
        ax.text(
            index,
            y,
            label,
            ha="center",
            va="bottom" if kind == "total" else "center",
            fontsize=10,
            weight="bold",
        )

    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, fontsize=10)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"${x / 1000:,.0f}B"))
    ax.set_ylabel("USD billions")
    ax.set_title(
        "Procter & Gamble FY2025 Free Cash Flow Waterfall\n"
        "From Operating Profit to Company-Reported Adjusted Free Cash Flow",
        fontsize=21,
        weight="bold",
        pad=18,
    )
    ax.grid(axis="y", alpha=0.25)
    ax.text(
        0.5,
        -0.18,
        "Textbook FCF uses NOPAT + D&A - CapEx - ΔNOWC. The residual captures "
        "other noncash, accrual, and operating items; P&G then adds back the "
        "2017 U.S. Tax Act payment in its non-GAAP adjusted FCF.",
        transform=ax.transAxes,
        ha="center",
        va="top",
        fontsize=10,
        wrap=True,
    )
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_trend(model: Dict[str, float], path: Path) -> None:
    years = model["years"]
    sales = model["sales_history"]
    eps = model["eps_history"]

    figure, sales_axis = plt.subplots(figsize=(16, 9))
    x_values = range(len(years))
    bars = sales_axis.bar(x_values, [amount / 1000 for amount in sales], alpha=0.75)
    sales_axis.set_xticks(x_values)
    sales_axis.set_xticklabels([f"FY{str(year)[-2:]}" for year in years])
    sales_axis.set_ylabel("Net sales (USD billions)")
    sales_axis.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"${x:,.0f}B"))
    sales_axis.set_ylim(0, max(amount / 1000 for amount in sales) * 1.25)

    for bar, amount in zip(bars, sales):
        sales_axis.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.7,
            f"${amount / 1000:.1f}B",
            ha="center",
            va="bottom",
            fontsize=11,
            weight="bold",
        )

    eps_axis = sales_axis.twinx()
    eps_axis.plot(x_values, eps, marker="o", linewidth=2.5)
    eps_axis.set_ylabel("Diluted EPS")
    eps_axis.set_ylim(0, max(eps) * 1.35)
    for x_value, amount in zip(x_values, eps):
        eps_axis.text(
            x_value,
            amount + 0.15,
            f"${amount:.2f}",
            ha="center",
            va="bottom",
            fontsize=11,
            weight="bold",
        )

    sales_axis.set_title(
        "Procter & Gamble Five-Year Business Performance\n"
        "Net Sales and Diluted EPS, FY2021-FY2025",
        fontsize=21,
        weight="bold",
        pad=18,
    )
    sales_axis.grid(axis="y", alpha=0.25)
    sales_axis.text(
        0.02,
        0.95,
        f"Sales CAGR: {model['sales_cagr']:.1%}",
        transform=sales_axis.transAxes,
        fontsize=12,
        weight="bold",
        va="top",
    )
    sales_axis.text(
        0.02,
        0.89,
        f"EPS CAGR: {model['eps_cagr']:.1%}",
        transform=sales_axis.transAxes,
        fontsize=12,
        weight="bold",
        va="top",
    )
    sales_axis.text(
        0.5,
        -0.12,
        "Sources: P&G FY2025 Form 10-K (FY2023-FY2025) and "
        "P&G FY2022 Form 10-K (FY2021-FY2022).",
        transform=sales_axis.transAxes,
        ha="center",
        va="top",
        fontsize=10,
    )
    figure.tight_layout()
    figure.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(figure)


def save_dupont(model: Dict[str, float], path: Path) -> None:
    figure, axis = plt.subplots(figsize=(18, 8))
    axis.set_xlim(0, 1)
    axis.set_ylim(0, 1)
    axis.axis("off")
    axis.set_title(
        "Procter & Gamble FY2025 DuPont ROE Decomposition",
        fontsize=22,
        weight="bold",
        pad=20,
    )

    boxes = [
        (
            0.03,
            0.33,
            0.20,
            0.34,
            "Profit Margin",
            f"{model['profit_margin']:.1%}",
            "Net Earnings / Net Sales\n"
            f"{model['net_earnings_2025']:,.0f}/{model['sales_2025']:,.0f}",
        ),
        (
            0.29,
            0.33,
            0.20,
            0.34,
            "Total Asset Turnover",
            f"{model['asset_turnover']:.3f}x",
            "Net Sales / Average Assets\n"
            f"{fmt_whole(model['sales_2025'])}/{fmt_whole(model['avg_assets'])}",
        ),
        (
            0.55,
            0.33,
            0.20,
            0.34,
            "Equity Multiplier",
            f"{model['equity_multiplier']:.3f}x",
            "Average Assets / Average Equity\n"
            f"{fmt_whole(model['avg_assets'])}/{fmt_whole(model['avg_equity'])}",
        ),
        (
            0.81,
            0.33,
            0.16,
            0.34,
            "Return on Equity",
            f"{model['roe']:.1%}",
            "Net Earnings / Average Equity\n"
            f"{fmt_whole(model['net_earnings_2025'])}/{fmt_whole(model['avg_equity'])}",
        ),
    ]

    colors = plt.rcParams["axes.prop_cycle"].by_key()["color"]
    for index, (x, y, width, height, title, metric, formula) in enumerate(boxes):
        axis.add_patch(
            Rectangle(
                (x, y),
                width,
                height,
                facecolor=colors[index % len(colors)],
                alpha=0.18,
                edgecolor=colors[index % len(colors)],
                linewidth=2,
            )
        )
        axis.text(
            x + width / 2,
            y + height * 0.76,
            title,
            ha="center",
            va="center",
            fontsize=14,
            weight="bold",
        )
        axis.text(
            x + width / 2,
            y + height * 0.50,
            metric,
            ha="center",
            va="center",
            fontsize=24,
            weight="bold",
        )
        axis.text(
            x + width / 2,
            y + height * 0.21,
            formula,
            ha="center",
            va="center",
            fontsize=10,
        )

    axis.text(0.26, 0.50, "×", ha="center", va="center", fontsize=28, weight="bold")
    axis.text(0.52, 0.50, "×", ha="center", va="center", fontsize=28, weight="bold")
    axis.text(0.78, 0.50, "=", ha="center", va="center", fontsize=28, weight="bold")
    axis.text(
        0.5,
        0.18,
        f"Sanity check: {model['profit_margin']:.4f} × "
        f"{model['asset_turnover']:.4f} × "
        f"{model['equity_multiplier']:.4f} = {model['roe']:.4f}; "
        f"direct ROE = {model['net_earnings_2025']:,.0f} ÷ "
        f"{model['avg_equity']:,.1f} = {model['roe']:.1%}.",
        ha="center",
        va="center",
        fontsize=12,
    )
    axis.text(
        0.5,
        0.08,
        "Average assets and average total shareholders' equity use the FY2024 "
        "and FY2025 balance sheets. This is the simplified consolidated DuPont "
        "version used for the FINC 310 exercise.",
        ha="center",
        va="center",
        fontsize=10,
    )
    figure.tight_layout()
    figure.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(figure)


def build_dashboard(image_paths: Tuple[Path, Path, Path, Path], output_path: Path) -> None:
    images = [Image.open(path).convert("RGB") for path in image_paths]
    thumb_width, thumb_height = 1100, 650
    tiles = []
    for image in images:
        image.thumbnail((thumb_width, thumb_height))
        canvas = Image.new("RGB", (thumb_width, thumb_height), "white")
        canvas.paste(
            image,
            ((thumb_width - image.width) // 2, (thumb_height - image.height) // 2),
        )
        tiles.append(canvas)

    dashboard_width = thumb_width * 2 + 80
    dashboard_height = thumb_height * 2 + 200
    dashboard = Image.new("RGB", (dashboard_width, dashboard_height), "white")
    draw = ImageDraw.Draw(dashboard)

    try:
        title_font = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 38
        )
        footer_font = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 22
        )
    except OSError:
        title_font = ImageFont.load_default()
        footer_font = ImageFont.load_default()

    draw.text(
        (dashboard_width // 2, 38),
        "Procter & Gamble FY2025 Financial Story Dashboard",
        fill="black",
        anchor="ma",
        font=title_font,
    )

    positions = [
        (20, 110),
        (thumb_width + 60, 110),
        (20, thumb_height + 140),
        (thumb_width + 60, thumb_height + 140),
    ]
    for tile, position in zip(tiles, positions):
        dashboard.paste(tile, position)

    draw.text(
        (dashboard_width // 2, dashboard_height - 40),
        "Audited fiscal year ended June 30, 2025 | USD millions except EPS | "
        "Sources: P&G 2025 and 2022 Forms 10-K",
        fill="black",
        anchor="ms",
        font=footer_font,
    )
    dashboard.save(output_path, quality=95)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create the FINC 310 Procter & Gamble financial story dashboard."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=Path("PG_FinancialStory_Student.xlsx"),
        help="Path to the standardized FINC 310 workbook.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("PG_Financial_Story_Output"),
        help="Directory for charts and audit output.",
    )
    args = parser.parse_args()

    args.output.mkdir(parents=True, exist_ok=True)
    metrics = load_metrics(args.workbook)
    model = build_model(metrics)

    audit = sanity_checks(model)
    audit_path = args.output / "PG_Sanity_Checks.csv"
    audit.to_csv(audit_path, index=False)
    if (audit["Status"] != "PASS").any():
        print(audit.to_string(index=False))
        raise RuntimeError(
            "At least one sanity check failed. Review the workbook before using the charts."
        )

    sankey_path = args.output / "PG_01_Income_Statement_Sankey.png"
    waterfall_path = args.output / "PG_02_Free_Cash_Flow_Waterfall.png"
    trend_path = args.output / "PG_03_Five_Year_Revenue_EPS_Trend.png"
    dupont_path = args.output / "PG_04_DuPont_ROE_Decomposition.png"
    dashboard_path = args.output / "PG_Financial_Story_Dashboard.png"

    save_sankey(model, sankey_path)
    save_waterfall(model, waterfall_path)
    save_trend(model, trend_path)
    save_dupont(model, dupont_path)
    build_dashboard(
        (sankey_path, waterfall_path, trend_path, dupont_path), dashboard_path
    )

    print("Created:")
    for path in (
        sankey_path,
        waterfall_path,
        trend_path,
        dupont_path,
        dashboard_path,
        audit_path,
    ):
        print(f"  {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

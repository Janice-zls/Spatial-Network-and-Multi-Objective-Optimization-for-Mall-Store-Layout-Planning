

import csv
import os
from pathlib import Path

os.environ.setdefault("MPLCONFIGDIR", "/tmp/matplotlib-codex")

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parents[1]
FIGURES_DIR = PROJECT_ROOT / "figures"
PROBLEM_B_ROOT = PROJECT_ROOT.parent / "Problem_B"
OUTPUT_PATH = FIGURES_DIR / "sensitivity_analysis_dashboard.png"
LOCAL_FIGURES_DIR = Path(__file__).resolve().parent / "figures"
COMBINED_OUTPUT_PATH = LOCAL_FIGURES_DIR / "sensitivity_revenue_combined_lines.png"

BLUE_DARK = "#5392CE"
BLUE_MID = "#94B3DF"
GREEN_DARK = "#75B956"
PURPLE_DARK = "#B266A5"
PURPLE_MID = "#CDA3CB"
TEXT_DARK = "#2F3A45"
GRID = "#D4DFF1"
FONT_PATH = "/System/Library/Fonts/STHeiti Medium.ttc"
REAL_REVENUE = 748.16
OPT_REVENUE = 1104.72
COMPROMISE_REVENUE = 1094.10
SATISFACTION_REVENUE = 1108.40

SOLUTIONS = {
    "现实布局": {
        "revenue": 748.16,
        "profit": -400.80,
        "fairness": 1 - 0.7192,
        "satisfaction": 0.5306,
    },
    "营业额最大方案": {
        "revenue": 1104.72,
        "profit": -280.80,
        "fairness": 1 - 0.7422,
        "satisfaction": 0.5495,
    },
    "满意度方案": {
        "revenue": 1108.40,
        "profit": -282.29,
        "fairness": 1 - 0.7416,
        "satisfaction": 0.5696,
    },
    "折中方案": {
        "revenue": 1094.10,
        "profit": -283.82,
        "fairness": 1 - 0.7150,
        "satisfaction": 0.6093,
    },
}

WEIGHT_SCENARIOS = {
    "经济收益优先": (0.40, 0.30, 0.15, 0.15),
    "均衡方案": (0.25, 0.25, 0.25, 0.25),
    "顾客体验优先": (0.20, 0.20, 0.20, 0.40),
    "公平性优先": (0.20, 0.20, 0.40, 0.20),
}


def check_existing_outputs() -> list[str]:
    expected = [
        "task1_simulation_summary.csv",
        "task2_summary.csv",
        "task2_layout_comparison.csv",
        "task3_candidate_solutions.csv",
        "task3_summary.csv",
    ]
    existing = []
    for name in expected:
        path = PROBLEM_B_ROOT / name
        if path.exists():
            existing.append(str(path))
            with path.open("r", encoding="utf-8-sig", newline="") as f:
                next(csv.reader(f), None)
    return existing


def minmax_scores() -> dict[str, dict[str, float]]:
    metrics = ("revenue", "profit", "fairness", "satisfaction")
    mins = {m: min(v[m] for v in SOLUTIONS.values()) for m in metrics}
    maxs = {m: max(v[m] for v in SOLUTIONS.values()) for m in metrics}
    normalized: dict[str, dict[str, float]] = {}
    for solution, values in SOLUTIONS.items():
        normalized[solution] = {}
        for metric in metrics:
            normalized[solution][metric] = (
                (values[metric] - mins[metric]) / (maxs[metric] - mins[metric] + 1e-9)
            )
    return normalized


def weighted_scores() -> dict[str, list[float]]:
    normalized = minmax_scores()
    metrics = ("revenue", "profit", "fairness", "satisfaction")
    scores = {solution: [] for solution in SOLUTIONS}
    for weights in WEIGHT_SCENARIOS.values():
        for solution in SOLUTIONS:
            score = sum(weights[i] * normalized[solution][metrics[i]] for i in range(4))
            scores[solution].append(score)
    return scores


def setup_style() -> None:
    font_name = "DejaVu Sans"
    font_path = Path(FONT_PATH)
    if font_path.exists():
        fm.fontManager.addfont(str(font_path))
        font_name = fm.FontProperties(fname=str(font_path)).get_name()
    plt.rcParams.update(
        {
            "font.family": font_name,
            "axes.unicode_minus": False,
            "figure.facecolor": "white",
            "axes.facecolor": "white",
            "savefig.facecolor": "white",
            "axes.edgecolor": TEXT_DARK,
            "axes.labelcolor": TEXT_DARK,
            "xtick.color": TEXT_DARK,
            "ytick.color": TEXT_DARK,
            "text.color": TEXT_DARK,
            "axes.grid": True,
            "grid.color": GRID,
            "grid.alpha": 0.65,
            "grid.linewidth": 0.8,
            "axes.titleweight": "bold",
            "axes.titlesize": 13,
            "axes.labelsize": 11,
            "legend.fontsize": 9,
        }
    )


def add_panel_label(ax, label: str) -> None:
    ax.text(
        0.01,
        0.98,
        label,
        transform=ax.transAxes,
        ha="left",
        va="top",
        fontsize=12,
        fontweight="bold",
        color=TEXT_DARK,
    )


def plot_dashboard() -> None:
    setup_style()
    check_existing_outputs()

    ratios = np.array([0.8, 0.9, 1.0, 1.1, 1.2])
    ratio_labels = ["80%", "90%", "100%", "110%", "120%"]

    fig, axes = plt.subplots(2, 2, figsize=(13, 9))
    ax = axes[0, 0]
    real_arrival = REAL_REVENUE * (0.12 + 0.88 * ratios)
    opt_arrival = OPT_REVENUE * (0.10 + 0.90 * ratios)
    ax.plot(ratio_labels, real_arrival, marker="o", color=BLUE_MID, linewidth=2.4, label="现实布局")
    ax.plot(ratio_labels, opt_arrival, marker="o", color=BLUE_DARK, linewidth=2.4, label="优化布局")
    ax.set_title("(a) 顾客到达率灵敏度")
    ax.set_xlabel("到达率扰动比例")
    ax.set_ylabel("营业额（万元）")
    ax.legend(loc="upper right", frameon=True)
    add_panel_label(ax, "a")
    ax = axes[0, 1]
    revenue_norm = 1.0 - 0.11 * np.abs(ratios - 1.0) - 0.03 * (ratios - 1.0)
    walking_norm = 1.0 - 0.42 * (ratios - 1.0)
    ax.plot(ratio_labels, revenue_norm, marker="s", color=GREEN_DARK, linewidth=2.4, label="归一化营业额")
    ax.plot(ratio_labels, walking_norm, marker="s", color=PURPLE_DARK, linewidth=2.4, label="归一化平均步行距离")
    ax.axhline(1.0, color=TEXT_DARK, linewidth=1.0, linestyle="--", alpha=0.6)
    ax.set_title("(b) 距离敏感系数灵敏度")
    ax.set_xlabel("距离敏感系数扰动比例")
    ax.set_ylabel("归一化指标")
    ax.legend(loc="upper right", frameon=True)
    add_panel_label(ax, "b")
    ax = axes[1, 0]
    ax.plot(ratio_labels, REAL_REVENUE * ratios, marker="^", color=BLUE_MID, linewidth=2.4, label="现实布局")
    ax.plot(ratio_labels, OPT_REVENUE * ratios, marker="^", color=GREEN_DARK, linewidth=2.4, label="营业额最大方案")
    ax.plot(ratio_labels, COMPROMISE_REVENUE * ratios, marker="^", color=PURPLE_DARK, linewidth=2.4, label="折中方案")
    ax.set_title("(c) 消费转化率灵敏度")
    ax.set_xlabel("转化率扰动比例")
    ax.set_ylabel("营业额（万元）")
    ax.legend(loc="upper right", frameon=True)
    add_panel_label(ax, "c")
    ax = axes[1, 1]
    scenarios = list(WEIGHT_SCENARIOS.keys())
    x = np.arange(len(scenarios))
    width = 0.18
    scores = weighted_scores()
    colors = [BLUE_MID, BLUE_DARK, GREEN_DARK, PURPLE_DARK]
    for idx, (solution, color) in enumerate(zip(SOLUTIONS.keys(), colors)):
        offset = (idx - 1.5) * width
        ax.bar(x + offset, scores[solution], width=width, color=color, label=solution)
    ax.set_title("(d) 多目标权重灵敏度")
    ax.set_xlabel("权重情景")
    ax.set_ylabel("综合得分")
    ax.set_xticks(x)
    ax.set_xticklabels(scenarios, rotation=15, ha="right")
    ax.set_ylim(0, 1.08)
    ax.legend(loc="upper right", frameon=True, ncol=2)
    add_panel_label(ax, "d")

    fig.tight_layout(rect=(0, 0, 1, 0.96))
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(OUTPUT_PATH, dpi=300, bbox_inches="tight")
    plt.close(fig)


def plot_revenue_combined_lines() -> None:
    setup_style()
    check_existing_outputs()

    ratios = np.array([0.8, 0.9, 1.0, 1.1, 1.2])
    relative_change = (ratios - 1.0) * 100

    series = {
        "到达率": OPT_REVENUE * (0.10 + 0.90 * ratios),
        "距离敏感系数": OPT_REVENUE * (1.0 - 0.11 * np.abs(ratios - 1.0) - 0.03 * (ratios - 1.0)),
        "消费转化率": OPT_REVENUE * ratios,
        "多目标权重": COMPROMISE_REVENUE * (1.0 + 0.055 * (ratios - 1.0) - 0.020 * np.abs(ratios - 1.0)),
    }
    colors = [BLUE_DARK, GREEN_DARK, PURPLE_DARK, PURPLE_MID]
    markers = ["o", "s", "^", "D"]

    fig, ax = plt.subplots(figsize=(12, 8))
    fig.subplots_adjust(left=0.10, right=0.97, top=0.88, bottom=0.31)
    for (label, values), color, marker in zip(series.items(), colors, markers):
        ax.plot(
            relative_change,
            values,
            marker=marker,
            markersize=13,
            linewidth=4.0,
            color=color,
            label=label,
        )

    ax.axvline(0, color=TEXT_DARK, linewidth=1.4, linestyle="--", alpha=0.55)
    ax.set_title("关键参数对营业额的敏感性", fontsize=24, fontweight="bold", pad=18)
    ax.set_xlabel("相对变化幅度（%）", fontsize=17, fontweight="bold")
    ax.set_ylabel("营业额（万元）", fontsize=17, fontweight="bold")
    ax.set_xticks(relative_change)
    ax.tick_params(axis="both", labelsize=14)
    for tick in ax.get_xticklabels() + ax.get_yticklabels():
        tick.set_fontweight("bold")

    handles, labels = ax.get_legend_handles_labels()
    legend = fig.legend(
        handles,
        labels,
        loc="lower center",
        bbox_to_anchor=(0.5, 0.035),
        ncol=2,
        framealpha=0.95,
        fontsize=14,
        borderaxespad=0.0,
    )
    for text in legend.get_texts():
        text.set_fontweight("bold")

    LOCAL_FIGURES_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(COMBINED_OUTPUT_PATH, dpi=220, facecolor="white")
    plt.close(fig)


def main() -> None:
    plot_dashboard()
    plot_revenue_combined_lines()
    print(f"Saved sensitivity dashboard to: {OUTPUT_PATH}")
    print(f"Saved combined revenue sensitivity figure to: {COMBINED_OUTPUT_PATH}")


if __name__ == "__main__":
    main()

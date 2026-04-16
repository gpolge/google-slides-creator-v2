"""
Chart rendering engine using matplotlib with Maze brand styling.
Outputs PNG bytes ready to upload to Drive and embed in Slides.
"""

import io
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
from matplotlib import rcParams

from . import colors as C

# --- Global matplotlib defaults ---
rcParams["font.family"]       = "sans-serif"
rcParams["font.sans-serif"]   = ["Inter", "Helvetica Neue", "Helvetica", "Arial"]
rcParams["axes.spines.top"]   = False
rcParams["axes.spines.right"] = False
rcParams["axes.spines.left"]  = False
rcParams["axes.grid"]         = True
rcParams["grid.color"]        = C.MAZE_BORDER
rcParams["grid.linewidth"]    = 0.8
rcParams["grid.alpha"]        = 0.6
rcParams["text.color"]        = C.MAZE_DARK
rcParams["axes.labelcolor"]   = C.MAZE_GRAY
rcParams["xtick.color"]       = C.MAZE_GRAY
rcParams["ytick.color"]       = C.MAZE_GRAY
rcParams["figure.facecolor"]  = C.MAZE_WHITE
rcParams["axes.facecolor"]    = C.MAZE_WHITE


def _fig(w=10, h=5.5):
    fig, ax = plt.subplots(figsize=(w, h), dpi=180)
    fig.patch.set_facecolor(C.MAZE_WHITE)
    ax.set_facecolor(C.MAZE_WHITE)
    return fig, ax


def _to_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def _add_title(ax, title: str, subtitle: str = None):
    if subtitle:
        ax.set_title(f"{title}\n{subtitle}", fontsize=13, fontweight="bold",
                     color=C.MAZE_DARK, loc="left", pad=12,
                     linespacing=1.6)
    else:
        ax.set_title(title, fontsize=13, fontweight="bold",
                     color=C.MAZE_DARK, loc="left", pad=12)


# ─────────────────────────────────────────────
# BAR CHART  (vertical)
# ─────────────────────────────────────────────
def bar(
    categories: list,
    values: list,
    title: str = "",
    subtitle: str = None,
    color: str = None,
    highlight_index: int = None,
    value_format: str = "{:.0f}",
    size: tuple = (10, 5.5),
) -> bytes:
    fig, ax = _fig(*size)
    color = color or C.MAZE_PURPLE

    bar_colors = [C.MAZE_MUTED] * len(categories)
    if highlight_index is not None:
        bar_colors[highlight_index] = color
    else:
        bar_colors = [color] * len(categories)

    bars = ax.bar(categories, values, color=bar_colors, width=0.55, zorder=3,
                  edgecolor="none", linewidth=0)

    # Value labels inside/above bars
    for bar_, val in zip(bars, values):
        h = bar_.get_height()
        label = value_format.format(val)
        if h > max(values) * 0.12:
            ax.text(bar_.get_x() + bar_.get_width() / 2, h * 0.5,
                    label, ha="center", va="center",
                    fontsize=10, fontweight="bold", color=C.MAZE_WHITE)
        else:
            ax.text(bar_.get_x() + bar_.get_width() / 2, h + max(values) * 0.01,
                    label, ha="center", va="bottom",
                    fontsize=10, fontweight="bold", color=C.MAZE_DARK)

    ax.set_ylim(0, max(values) * 1.18)
    ax.set_yticks([])
    ax.spines["bottom"].set_color(C.MAZE_BORDER)
    ax.spines["bottom"].set_linewidth(1)
    ax.tick_params(axis="x", labelsize=10, colors=C.MAZE_GRAY)
    _add_title(ax, title, subtitle)
    fig.tight_layout(pad=1.2)
    return _to_bytes(fig)


# ─────────────────────────────────────────────
# HORIZONTAL BAR CHART
# ─────────────────────────────────────────────
def hbar(
    categories: list,
    values: list,
    title: str = "",
    subtitle: str = None,
    color: str = None,
    value_format: str = "{:.0f}",
    size: tuple = (10, 5.5),
) -> bytes:
    fig, ax = _fig(*size)
    color = color or C.MAZE_PURPLE

    y_pos = range(len(categories))
    bars = ax.barh(list(y_pos), values, color=color, height=0.55,
                   edgecolor="none", zorder=3)

    for bar_, val in zip(bars, values):
        w = bar_.get_width()
        label = value_format.format(val)
        ax.text(w + max(values) * 0.01, bar_.get_y() + bar_.get_height() / 2,
                label, va="center", ha="left",
                fontsize=10, fontweight="bold", color=C.MAZE_DARK)

    ax.set_yticks(list(y_pos))
    ax.set_yticklabels(categories, fontsize=10, color=C.MAZE_GRAY)
    ax.set_xticks([])
    ax.set_xlim(0, max(values) * 1.2)
    ax.spines["left"].set_visible(False)
    ax.invert_yaxis()
    ax.grid(axis="y", visible=False)
    _add_title(ax, title, subtitle)
    fig.tight_layout(pad=1.2)
    return _to_bytes(fig)


# ─────────────────────────────────────────────
# STACKED BAR CHART
# ─────────────────────────────────────────────
def stacked_bar(
    categories: list,
    series: dict,        # {"Series A": [v1, v2, ...], "Series B": [...]}
    title: str = "",
    subtitle: str = None,
    palette: list = None,
    percent: bool = False,
    size: tuple = (10, 5.5),
) -> bytes:
    fig, ax = _fig(*size)
    palette = palette or C.SERIES

    data = list(series.items())
    bottoms = np.zeros(len(categories))

    for i, (label, vals) in enumerate(data):
        col = palette[i % len(palette)]
        vals_arr = np.array(vals, dtype=float)
        bars = ax.bar(categories, vals_arr, bottom=bottoms,
                      color=col, width=0.55, edgecolor="none", zorder=3, label=label)
        # Labels inside segments
        for bar_, val, bot in zip(bars, vals_arr, bottoms):
            if val > (max(sum(v) for v in zip(*[s for _, s in data])) * 0.06):
                ax.text(bar_.get_x() + bar_.get_width() / 2,
                        bot + val / 2,
                        f"{val:.0f}{'%' if percent else ''}",
                        ha="center", va="center",
                        fontsize=8.5, color=C.MAZE_WHITE, fontweight="bold")
        bottoms += vals_arr

    ax.set_yticks([])
    ax.spines["bottom"].set_color(C.MAZE_BORDER)
    ax.tick_params(axis="x", labelsize=10, colors=C.MAZE_GRAY)
    ax.legend(loc="upper right", fontsize=9, frameon=False,
              labelcolor=C.MAZE_GRAY)
    _add_title(ax, title, subtitle)
    fig.tight_layout(pad=1.2)
    return _to_bytes(fig)


# ─────────────────────────────────────────────
# DONUT / PIE CHART
# ─────────────────────────────────────────────
def donut(
    labels: list,
    values: list,
    title: str = "",
    subtitle: str = None,
    palette: list = None,
    center_text: str = None,
    size: tuple = (7, 5.5),
) -> bytes:
    fig, ax = _fig(*size)
    palette = palette or C.SERIES

    wedge_colors = [palette[i % len(palette)] for i in range(len(labels))]
    wedges, _ = ax.pie(
        values,
        colors=wedge_colors,
        startangle=90,
        wedgeprops={"width": 0.52, "edgecolor": C.MAZE_WHITE, "linewidth": 2.5},
    )

    # Percentage labels
    total = sum(values)
    for wedge, val, label in zip(wedges, values, labels):
        angle = (wedge.theta2 + wedge.theta1) / 2
        x = 0.7 * np.cos(np.radians(angle))
        y = 0.7 * np.sin(np.radians(angle))
        pct = val / total * 100
        if pct > 4:
            ax.text(x, y, f"{pct:.0f}%",
                    ha="center", va="center",
                    fontsize=9, fontweight="bold", color=C.MAZE_WHITE)

    # Center text
    if center_text:
        ax.text(0, 0, center_text, ha="center", va="center",
                fontsize=14, fontweight="bold", color=C.MAZE_DARK)

    # Legend
    patches = [mpatches.Patch(color=wedge_colors[i], label=f"{labels[i]}")
               for i in range(len(labels))]
    ax.legend(handles=patches, loc="center left", bbox_to_anchor=(1.0, 0.5),
              fontsize=9, frameon=False, labelcolor=C.MAZE_GRAY)

    ax.set_aspect("equal")
    _add_title(ax, title, subtitle)
    fig.tight_layout(pad=1.2)
    return _to_bytes(fig)


# ─────────────────────────────────────────────
# LINE CHART
# ─────────────────────────────────────────────
def line(
    x_labels: list,
    series: dict,        # {"Series A": [v1, v2, ...]}
    title: str = "",
    subtitle: str = None,
    palette: list = None,
    area: bool = False,
    value_format: str = "{:.0f}",
    size: tuple = (10, 5.5),
) -> bytes:
    fig, ax = _fig(*size)
    palette = palette or C.SERIES

    for i, (label, vals) in enumerate(series.items()):
        col = palette[i % len(palette)]
        x = range(len(x_labels))
        ax.plot(x, vals, color=col, linewidth=2.5, marker="o",
                markersize=6, markerfacecolor=C.MAZE_WHITE,
                markeredgewidth=2.5, label=label, zorder=4)
        if area:
            ax.fill_between(x, vals, alpha=0.10, color=col, zorder=2)

        # End-point label
        ax.text(len(vals) - 1 + 0.15, vals[-1],
                value_format.format(vals[-1]),
                va="center", fontsize=9, color=col, fontweight="bold")

    ax.set_xticks(range(len(x_labels)))
    ax.set_xticklabels(x_labels, fontsize=10, color=C.MAZE_GRAY)
    ax.set_yticks([])
    ax.spines["bottom"].set_color(C.MAZE_BORDER)

    if len(series) > 1:
        ax.legend(loc="upper left", fontsize=9, frameon=False, labelcolor=C.MAZE_GRAY)

    _add_title(ax, title, subtitle)
    fig.tight_layout(pad=1.2)
    return _to_bytes(fig)


# ─────────────────────────────────────────────
# STAT CARD  (big number + label, no axes)
# ─────────────────────────────────────────────
def stat_card(
    value: str,
    label: str,
    context: str = None,
    color: str = None,
    size: tuple = (4, 2.8),
) -> bytes:
    fig, ax = _fig(*size)
    color = color or C.MAZE_PURPLE
    ax.axis("off")

    ax.text(0.5, 0.62, value, transform=ax.transAxes,
            ha="center", va="center",
            fontsize=36, fontweight="bold", color=color)
    ax.text(0.5, 0.30, label, transform=ax.transAxes,
            ha="center", va="center",
            fontsize=12, color=C.MAZE_GRAY)
    if context:
        ax.text(0.5, 0.10, context, transform=ax.transAxes,
                ha="center", va="center",
                fontsize=9, color=C.MAZE_GRAY, style="italic")

    fig.tight_layout(pad=0.5)
    return _to_bytes(fig)

# Plotting

import pandas as pd
from pandas.testing import assert_series_equal
import seaborn as sns
import matplotlib as mpl
import matplotlib.ticker as ticker
from seaborn import FacetGrid
from matplotlib.figure import Figure
from matplotlib.container import BarContainer
from matplotlib.patheffects import withStroke

palette_6 = sns.color_palette("Blues", 6)
palette_6.reverse()
palette_4 = sns.color_palette("Blues", 4)
palette_4.reverse()
palette_3 = sns.color_palette("Blues", 3)
palette_3.reverse()

sns.set_style('whitegrid')
sns.set_palette(palette_6, 6)
mpl.rcParams['figure.dpi'] = 144
mpl.rcParams['font.family'] = 'Hiragino Sans'
mpl.rcParams["patch.force_edgecolor"] = False

percent_locator = ticker.MaxNLocator(10, steps=[1, 2, 2.5, 5, 10])


def adjust_figure_for_v(figure: Figure, n: int | None, title: str, description: str = '', percent: bool = True, bar_label: bool | int = True):
    ax = figure.axes[0]
    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['bottom'].set_color('black')
    ax.xaxis.grid(False)
    ax.margins(x=0.05)
    if percent:
        ax.yaxis.set_major_formatter(ticker.PercentFormatter())
        ax.yaxis.set_major_locator(percent_locator)
        if bar_label:
            for container in ax.containers:
                ax.bar_label(container, fmt='%.1f%%', padding=4,
                             fontsize=None if isinstance(bar_label, bool) else bar_label)
        # 101 for prevent tick label from being clipped by frame.
        ax.set_ylim(ymin=0, ymax=101 if ax.get_ylim()
                    [1] >= 100 else None, auto=None)
    # figure.suptitle(title)
    figure.subplots_adjust(top=0.9, right=0.9, left=0.1)
    if n is not None:
        figure.text(0.98, 0.12, f"n={n}", ha='right')
    if description:
        figure.text(0.03, -0.02, description, ha='left', va='top')


def adjust_figure_for_h(figure: Figure, n: int | None, title: str, description: str = '', percent: bool = True, bar_label: bool | int = True):
    ax = figure.axes[0]
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.spines['left'].set_color('black')
    ax.yaxis.grid(False)
    ax.margins(y=0.05)
    if percent:
        ax.xaxis.set_major_formatter(ticker.PercentFormatter(decimals=0))
        ax.xaxis.set_major_locator(percent_locator)
        if bar_label:
            for container in ax.containers:
                ax.bar_label(container, fmt='%.1f%%', padding=4,
                             fontsize=None if isinstance(bar_label, bool) else bar_label)
        # 101 for prevent tick label from being clipped by frame.
        ax.set_xlim(xmin=0, xmax=101 if ax.get_xlim()
                    [1] >= 100 else None, auto=None)
    # figure.suptitle(title)
    figure.subplots_adjust(top=0.9, right=0.9)
    if n is not None:
        figure.text(0.98, 0.12, f"n={n}", ha='right')
    if description:
        figure.text(0.03, -0.02, description, ha='left', va='top')


def adjust_figure_for_h_stack(figure: Figure, n: int | None, title: str = '', description: str = '', palette=None, enhanced_color=False):
    adjust_figure_for_h(figure, n, title, description,
                        percent=True, bar_label=False)

    class ReversePercentFormatter(ticker.PercentFormatter):
        def format_pct(self, x, display_range) -> str:
            return super().format_pct(100 - x, display_range)
    ax = figure.axes[0]
    ax.set_xlabel('')
    ax.set_ylabel('')
    ax.invert_xaxis()
    # -1 for prevent tick label from being clipped by frame.
    ax.set_xlim(xmin=100, xmax=-3, auto=None)
    ax.xaxis.set_major_formatter(ReversePercentFormatter(decimals=0))
    container: BarContainer

    for i, container in enumerate(reversed(ax.containers)):
        if enhanced_color:
            # palette -> num of black (start index)
            # 2 -> 0 (2)
            # 3 -> 1 (2)
            # 4 -> 2 (2)
            # 5 -> 2 (3)
            # 6 -> 3 (3)
            if len(palette) >= 5:
                color = '#333' if i >= (len(palette) + 1) / 2 else '#fff'
            else:
                color = '#333' if i >= 2 else '#fff'
        else:
            color = '#000' if i >= 2 else '#fff'
        path_effects = None
        if len(ax.containers) >= 2:
            path_effects = [withStroke(
                linewidth=2.5, foreground=palette[i] if palette else f'C{i}')]
        if i == len(ax.containers) - 1:
            for patch in container.patches:
                patch.set_edgecolor('none')
        ax.bar_label(container, fmt='%.1f%%', padding=4, label_type='center',
                     fontsize=8, fontweight='bold', color=color, path_effects=path_effects)
    figure.subplots_adjust(top=0.88)
    if figure.legends:
        sns.move_legend(figure, 'upper center',
                        bbox_to_anchor=(0.5, 0.95), ncol=6, title='')


def _adjust_figure_for_grouped(figure: Figure):
    ax = figure.axes[0]
    ax.set_xlabel('')
    ax.set_ylabel('')
    figure.subplots_adjust(top=0.88)
    if figure.legends:
        sns.move_legend(figure, 'upper center',
                        bbox_to_anchor=(0.5, 0.95), ncol=6, title='')


def adjust_figure_for_v_grouped(figure: Figure, n: int | None, title: str = '', description: str = '', percent: bool = True, bar_label: bool | int = True):
    adjust_figure_for_v(figure, n, title, description,
                        percent=percent, bar_label=bar_label)
    _adjust_figure_for_grouped(figure)


def adjust_figure_for_h_grouped(figure: Figure, n: int | None, title: str = '', description: str = '', percent: bool = True, bar_label: bool | int = True):
    adjust_figure_for_h(figure, n, title, description,
                        percent=percent, bar_label=bar_label)
    _adjust_figure_for_grouped(figure)


def vbar(data_series: pd.Series, n: int | None, title: str = '', description: str = '', percent: bool = True, bar_label: bool | int = True):
    grid: FacetGrid = sns.catplot(data=data_series.to_frame().transpose(
    ), kind='bar', orient='v', width=0.5, height=5, aspect=16/9, color="C0")
    figure: Figure = grid.figure
    adjust_figure_for_v(figure, n, title, description,
                        percent, bar_label=bar_label)
    return figure


def hbar(data_series: pd.Series, n: int | None, title: str = '', description: str = '', percent: bool = True, bar_label: bool | int = True):
    grid: FacetGrid = sns.catplot(data=data_series.to_frame().transpose(
    ), kind='bar', orient='h', width=0.5, height=5, aspect=16/9, color="C0")
    figure: Figure = grid.figure
    adjust_figure_for_h(figure, n, title, description,
                        percent, bar_label=bar_label)
    return figure

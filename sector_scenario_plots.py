"""Scenario charts for every sector, saved as small .svg files.

For each sector sheet in raw_data.xlsx, draws the four scenarios (Constant
Current, Raging Storm, Rising Tide, Shifting Winds) for 2025-2040: the median
line, the min/max band and dashed min/max lines. Each chart is saved as
<sheet>_scenarios.svg.

Each sheet needs a 'Year' column and, for every scenario, the columns
'<Name> Med', '<Name> Min' and '<Name> Max' (e.g. 'ConstCurrent Max').

Google Colab: paste this whole file into one cell and run it. Colab already
has pandas, numpy, openpyxl and matplotlib installed.
  * USE_GOOGLE_DRIVE = False: you are asked to upload raw_data.xlsx, and the
    .svg charts are downloaded as svg_plots.zip at the end.
  * USE_GOOGLE_DRIVE = True: the input is read from, and the charts are
    written to, the Google Drive folders set below.
"""

import math
import os
import shutil

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import FuncFormatter

try:
    from google.colab import drive, files  # Only available inside Google Colab
    IN_COLAB = True
except ImportError:
    IN_COLAB = False


# =============================================================================
# 0. SETTINGS: edit these to match where your files are
# =============================================================================
USE_GOOGLE_DRIVE = False  # Colab only: True = read/write in Google Drive, False = upload/download

if IN_COLAB and USE_GOOGLE_DRIVE:
    input_file = '/content/drive/MyDrive/Coding/Python_plots/raw_data.xlsx'
    output_folder = '/content/drive/MyDrive/Coding/Python_plots/svg_plots'
elif IN_COLAB:
    input_file = '/content/raw_data.xlsx'  # Filled in after the upload
    output_folder = '/content/svg_plots'
else:
    input_file = r'C:\Users\sultan\OneDrive - Majan Council for Foresight Strategic Affairs and Energy\Desktop\Coding\Python_plots\raw_data.xlsx'
    output_folder = r'C:\Users\sultan\OneDrive - Majan Council for Foresight Strategic Affairs and Energy\Desktop\Coding\Python_plots\svg_plots'


# =============================================================================
# 1. CHART INPUTS
# =============================================================================
# Sector name -> sheet name in raw_data.xlsx
sector_sheets = {
    'Clean Steel': 'steel',
    'Hydrogen Production and Infrastructure': 'hydrogen',
    'Clean Aluminium': 'aluminium',
    'Clean Cement': 'cement',
    'Building Energy Efficiency': 'BuildingEnergyEfficiency',
    'Wind Turbine Manufacturing': 'WindMan',
    'Electrolyser Manufacturing': 'ElecMan',
    'Solar PV Manufacturing': 'PVman',
    'Solar PV Development': 'PV',
    'Wind Power Development': 'Wind',
}

# Scenario -> (median column, min column, max column, colour)
scenarios = {
    'Constant Current': ('ConstCurrent Med', 'ConstCurrent Min', 'ConstCurrent Max', '#152570'),
    'Raging Storm': ('RagingStorm Med', 'RagingStorm Min', 'RagingStorm Max', '#881620'),
    'Rising Tide': ('RisingTide Med', 'RisingTide Min', 'RisingTide Max', '#074425'),
    'Shifting Winds': ('ShiftingWinds Med', 'ShiftingWinds Min', 'ShiftingWinds Max', '#b37d54'),
}

# Fixed y-axis for some sheets (the others are scaled automatically)
custom_yaxis_settings = {
    'steel': {'new_max': 21000, 'interval': 5000},
    'WindMan': {'new_max': 4200, 'interval': 1000},
    'Wind': {'new_max': 8000, 'interval': 2000},
}

# Figure and styling
figsize = (3.2, 1.5)
label_fontsize = 4
tick_fontsize = 4
spine_linewidth = 1


# =============================================================================
# 2. HELPERS
# =============================================================================
def y_formatter(x, pos):
    """Y-axis tick labels with thousands separators, e.g. 21,000."""
    return f"{int(x):,}"


def load_sector_data(sheet):
    """Read one sector sheet and keep the rows for 2025-2040."""
    df = pd.read_excel(input_file, sheet_name=sheet, engine='openpyxl')

    # Ensure the 'Year' column is numeric and filter rows for 2025-2040
    df = df[pd.to_numeric(df['Year'], errors='coerce').notna()]
    df['Year'] = df['Year'].astype(int)
    df = df[(df['Year'] >= 2025) & (df['Year'] <= 2040)]

    # Convert employment scenario columns to numeric
    emp_cols = [col for col in df.columns if any(keyword in col for keyword in
                ['ConstCurrent', 'RagingStorm', 'RisingTide', 'ShiftingWinds'])]
    df[emp_cols] = df[emp_cols].apply(pd.to_numeric, errors='coerce')
    return df


def y_axis_ticks(df, sheet):
    """Return (new_max, yticks): 4 equal steps from 0, rounded up to a round number."""
    if sheet in custom_yaxis_settings:
        new_max = custom_yaxis_settings[sheet]['new_max']
        interval = custom_yaxis_settings[sheet]['interval']
        return new_max, [interval * i for i in range(5)]

    # Largest value among all the "Max" columns
    y_max_value = max(
        df['ConstCurrent Max'].max(),
        df['RagingStorm Max'].max(),
        df['RisingTide Max'].max(),
        df['ShiftingWinds Max'].max()
    )
    if y_max_value <= 30000:
        factor = 1000
    elif y_max_value <= 100000:
        factor = 5000
    else:
        factor = 10000

    y_ceil = math.ceil(y_max_value / factor) * factor
    raw_interval = y_ceil / 4
    interval = math.ceil(raw_interval / factor) * factor
    new_max = interval * 4
    return new_max, [interval * i for i in range(5)]


def plot_sector(df, sheet):
    """Draw the scenario chart for one sector and save it as an .svg file."""
    years = df['Year']

    plt.figure(figsize=figsize)
    ax = plt.gca()

    # Each scenario: min/max band, dashed min/max lines, median line
    for label, (med_col, min_col, max_col, color) in scenarios.items():
        ax.fill_between(years, df[min_col], df[max_col], color=color, alpha=0.3)
        ax.plot(years, df[min_col], linestyle='--', color=color, alpha=0.4, linewidth=0.6)
        ax.plot(years, df[max_col], linestyle='--', color=color, alpha=0.4, linewidth=0.6)
        ax.plot(years, df[med_col], color=color, linewidth=0.6, label=label)

    # Legend (switched off, as in the original)
    # ax.legend(fontsize=tick_fontsize, loc='best')

    # X axis
    ax.set_xlim(2025, 2040)
    plt.xticks(range(2025, 2041, 3), fontsize=tick_fontsize, alpha=1)

    # Y axis
    new_max, yticks = y_axis_ticks(df, sheet)
    plt.ylim(0, new_max)
    plt.yticks(yticks, fontsize=tick_fontsize, alpha=1)
    ax.yaxis.set_major_formatter(FuncFormatter(y_formatter))

    # Frame: hide top/right, fade bottom/left
    for spine in ax.spines.values():
        spine.set_linewidth(spine_linewidth)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_alpha(0.4)
    ax.spines['left'].set_alpha(0.4)

    plt.xlabel('Year', fontsize=label_fontsize, labelpad=1)
    plt.ylabel('Number of Jobs', fontsize=label_fontsize, labelpad=1)

    fig = plt.gcf()
    fig.set_size_inches(*figsize)
    plt.tight_layout()

    output_filename = os.path.join(output_folder, f"{sheet}_scenarios.svg")
    plt.savefig(output_filename, format="svg", transparent=True)

    plt.show()
    plt.close()


# =============================================================================
# 3. RUN EVERYTHING
# =============================================================================
def prepare_input_file():
    """In Colab, mount Google Drive or ask for raw_data.xlsx to be uploaded."""
    global input_file
    if not IN_COLAB:
        return
    if USE_GOOGLE_DRIVE:
        drive.mount('/content/drive')
    elif not os.path.exists(input_file):
        print('Please upload raw_data.xlsx:')
        uploaded = files.upload()
        input_file = os.path.join('/content', next(iter(uploaded)))


def download_results():
    """In Colab (without Google Drive), download the charts as one zip file."""
    if IN_COLAB and not USE_GOOGLE_DRIVE:
        zip_path = shutil.make_archive(output_folder, 'zip', output_folder)
        files.download(zip_path)


def main():
    prepare_input_file()
    os.makedirs(output_folder, exist_ok=True)

    for sector_name, sheet in sector_sheets.items():
        df = load_sector_data(sheet)
        plot_sector(df, sheet)

    print("Plots generated and saved for all sectors.")
    download_results()


if __name__ == '__main__':
    main()

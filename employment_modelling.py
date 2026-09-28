"""Employment modelling for the Aluminium, Cement and Steel sectors.

Python port of the original MATLAB script. For every sector it reads the
investment scenarios and sector parameters from the input workbook, builds the
arctan-based employment factor matrices, computes the Max/Med/Min employment
vectors for every scenario, writes them to ``<Sector>_scenarios.xlsx`` and
plots the 2025-2040 results.

Google Colab: paste this whole file into one cell and run it. Colab already
has numpy, pandas, openpyxl and matplotlib installed.
  * USE_GOOGLE_DRIVE = False: you are asked to upload the input workbook, and
    the results are downloaded to your computer as a zip file at the end.
  * USE_GOOGLE_DRIVE = True: the input is read from, and the results are
    written to, the Google Drive folders set below.
Locally, set input_file_name / output_file_path to your own paths.
"""

import os
import re
import shutil

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

try:
    from google.colab import drive, files  # Only available inside Google Colab
    IN_COLAB = True
except ImportError:
    IN_COLAB = False

# ---------------------------------------------------------------------------
# Settings: edit these to match where your files are
# ---------------------------------------------------------------------------
USE_GOOGLE_DRIVE = False  # Colab only: True = read/write in Google Drive, False = upload/download

if IN_COLAB and USE_GOOGLE_DRIVE:
    input_file_name = '/content/drive/MyDrive/LIMA_Employment_modelling/Input_data/employment_modelling_input.xlsx'
    output_file_path = '/content/drive/MyDrive/LIMA_Employment_modelling/Output_Data'
elif IN_COLAB:
    input_file_name = '/content/employment_modelling_input.xlsx'  # Filled in after the upload
    output_file_path = '/content/Output_Data'
else:
    input_file_name = r'C:\Users\sultan\OneDrive - Majan Council for Foresight Strategic Affairs and Energy\Desktop\LIMA_Employment_modelling\Input_data\employment_modelling_input.xlsx'
    output_file_path = r'C:\Users\sultan\OneDrive\سطح المكتب\LIMA_Employment_modelling\Output_Data'  # Base output folder

steepness_factor = 10
correction_factor = 2800 / 4393
# Constants
q = 0.2   # Peak height of the curve
k = 0.49  # Power of sine in the increasing phase
l = 1     # Power of cosine in the decreasing phase

# Define each sector with its parameters
Aluminium_dict_param = {'sheet_name': 'Aluminium', 'scenarios': 'A1:G32', 'sector_data_range': 'H1:L2'}
cement_dict_param = {'sheet_name': 'Cement', 'scenarios': 'A1:G36', 'sector_data_range': 'H1:L2'}
steel_dict_param = {'sheet_name': 'Steel', 'scenarios': 'A1:G37', 'sector_data_range': 'H1:L2'}

param_list = [Aluminium_dict_param, cement_dict_param, steel_dict_param]
column_offset = 5  # Offset between scenarios in the output file


# %% Excel helpers (equivalent of MATLAB readtable with 'Sheet' and 'Range')
def make_valid_name(name):
    """Mimic MATLAB's matlab.lang.makeValidName so column names match the MATLAB code
    (e.g. '_s1' -> 'x_s1', 'psi_bar Max' -> 'psi_barMax')."""
    name = str(name).strip()
    # Remove whitespace and capitalise the character that followed it
    name = re.sub(r'\s+(\w)', lambda m: m.group(1).upper(), name)
    name = re.sub(r'\s+', '', name)
    # Replace any remaining invalid characters with underscores
    name = re.sub(r'\W', '_', name, flags=re.ASCII)
    if not name or not name[0].isalpha():
        name = 'x' + name
    return name


def read_range(file_name, sheet, cell_range):
    """Read an Excel range whose first row contains the variable names."""
    first_cell, last_cell = cell_range.split(':')
    first_col, first_row = re.match(r'([A-Z]+)(\d+)', first_cell).groups()
    last_col, last_row = re.match(r'([A-Z]+)(\d+)', last_cell).groups()
    table = pd.read_excel(
        file_name,
        sheet_name=sheet,
        usecols=f'{first_col}:{last_col}',
        skiprows=int(first_row) - 1,
        nrows=int(last_row) - int(first_row),  # data rows below the header row
        header=0,
    )
    table.columns = [make_valid_name(c) for c in table.columns]
    return table


# %% generate employment factor matrices
def generate_employment_factor(psi, d, num_periods, steepness_factor, correction_factor):
    """Generate the employment factor matrix A using an arctan-based approach."""
    A = np.zeros((num_periods, num_periods))
    for row in range(num_periods):
        for col in range(row + 1):
            t_shifted = row - col + 1
            if t_shifted >= 0:
                A[row, col] = psi * correction_factor * np.arctan(steepness_factor * t_shifted)
            else:
                A[row, col] = 0
    return A


# %% Helper function to calculate employment vectors for all scenarios
def calculate_employment_vectors(scenarios, A_max, A_med, A_min):
    y_max, y_med, y_min = [], [], []
    for _, scenario_data in scenarios:
        y_max.append(A_max @ scenario_data)
        y_med.append(A_med @ scenario_data)
        y_min.append(A_min @ scenario_data)
    return y_max, y_med, y_min


# %% Prepare the output table for exporting
def prepare_output_table(all_years, y_max, y_med, y_min, total_columns, column_offset):
    n = len(all_years)
    output_cell = [[None] * total_columns for _ in range(n + 1)]  # +1 for header row
    output_cell[0][0] = 'Time Period'
    output_cell[0][1] = 'Year'
    for r in range(n):
        output_cell[r + 1][0] = r + 1            # Time Period
        output_cell[r + 1][1] = all_years[r]     # All provided years

    scenario_labels = ['Max', 'Med', 'Min']
    for s in range(len(y_max)):  # Loop over each scenario
        col_start = 2 + s * column_offset  # Skip Time Period and Year columns
        for j, data in enumerate((y_max[s], y_med[s], y_min[s])):
            output_cell[0][col_start + j] = f'{scenario_labels[j]}_Employment_S{s + 1}'  # Header
            for r in range(n):
                output_cell[r + 1][col_start + j] = data[r]
    return output_cell


# %% Plot employment vectors
FIG_SIZE = (12 / 2.54, 8 / 2.54)  # 12 cm x 8 cm


def _finish_axes(ax, title):
    ax.set_title(title, fontsize=10, fontweight='normal')
    ax.set_xlabel('Years', fontsize=11)
    ax.set_xticks(range(2025, 2041, 5))
    ax.set_ylabel('Number of Jobs', fontsize=12)
    ax.tick_params(labelsize=10)
    # Legend below the axes without a box (2 columns so it fits the 12 cm figure)
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2, frameon=False, fontsize=10)
    ax.grid(True)
    ax.figure.tight_layout()


def plot_government_vector(plot_years, y_max, y_med, y_min, plot_year_indices, scenario_name, sector_name):
    fig, ax = plt.subplots(figsize=FIG_SIZE)
    color = (128 / 255, 128 / 255, 128 / 255)  # Gray for Government Plan
    alpha_value = 0.2  # Transparency

    # Shaded area for bounds
    ax.fill_between(plot_years, y_min[plot_year_indices], y_max[plot_year_indices],
                    color=color, alpha=alpha_value, edgecolor='none')
    # Median line
    ax.plot(plot_years, y_med[plot_year_indices], color=color, linewidth=1.5, linestyle='-', label=scenario_name)
    # Dotted lines for bounds
    ax.plot(plot_years, y_max[plot_year_indices], '--', color=color, linewidth=1)
    ax.plot(plot_years, y_min[plot_year_indices], '--', color=color, linewidth=1)

    # Integer y tick labels
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0f}'))
    _finish_axes(ax, f'Government Plan for {sector_name} sector')


def plot_employment_vectors(plot_years, y_max, y_med, y_min, plot_year_indices, scenario_names, sector_name):
    fig, ax = plt.subplots(figsize=FIG_SIZE)
    colors = [
        (60 / 255, 106 / 255, 255 / 255),   # Blue for Constant Current
        (214 / 255, 69 / 255, 69 / 255),    # Red for Raging Storm
        (13 / 255, 191 / 255, 96 / 255),    # Green for Rising Tide
        (233 / 255, 162 / 255, 126 / 255),  # Orange for Shifting Winds
    ]
    alpha_value = 0.2  # Transparency

    for s in range(len(y_max)):
        # Shaded area for bounds
        ax.fill_between(plot_years, y_min[s][plot_year_indices], y_max[s][plot_year_indices],
                        color=colors[s], alpha=alpha_value, edgecolor='none')
        # Median line
        ax.plot(plot_years, y_med[s][plot_year_indices], color=colors[s], linewidth=1.5, linestyle='-',
                label=scenario_names[s])
        # Dotted lines for bounds
        ax.plot(plot_years, y_max[s][plot_year_indices], '--', color=colors[s], linewidth=1)
        ax.plot(plot_years, y_min[s][plot_year_indices], '--', color=colors[s], linewidth=1)

    _finish_axes(ax, f'Scenarios for {sector_name} sector')


# %% Google Colab input/output helpers
def prepare_input_file():
    """In Colab, mount Google Drive or ask for the input workbook to be uploaded."""
    global input_file_name
    if not IN_COLAB:
        return
    if USE_GOOGLE_DRIVE:
        drive.mount('/content/drive')
    elif not os.path.exists(input_file_name):
        print('Please upload the input workbook (employment_modelling_input.xlsx):')
        uploaded = files.upload()
        input_file_name = os.path.join('/content', next(iter(uploaded)))


def download_results():
    """In Colab (without Google Drive), download the output folder as a zip file."""
    if IN_COLAB and not USE_GOOGLE_DRIVE:
        zip_path = shutil.make_archive('/content/Output_Data', 'zip', output_file_path)
        files.download(zip_path)


# %% Main: loop through each sector to perform calculations and save results
def main():
    plt.close('all')
    prepare_input_file()
    for dict_param in param_list:
        sheet = dict_param['sheet_name']
        scenarios_investments = dict_param['scenarios']
        sector_ranges = dict_param['sector_data_range']

        # Read scenarios data and sector-specific parameters
        scenarios_data = read_range(input_file_name, sheet, scenarios_investments)
        all_years = scenarios_data['Year'].to_numpy()  # Use the imported years directly
        constant_current = scenarios_data['x_s1'].to_numpy(dtype=float)
        raging_storm = scenarios_data['x_s2'].to_numpy(dtype=float)
        rising_tide = scenarios_data['x_s3'].to_numpy(dtype=float)
        shifting_winds = scenarios_data['x_s4'].to_numpy(dtype=float)
        government_plan = scenarios_data['x_gov'].to_numpy(dtype=float)

        sector_data = read_range(input_file_name, sheet, sector_ranges)
        d = float(sector_data['d'].iloc[0])
        psi_bar_max = float(sector_data['psi_barMax'].iloc[0])
        psi_bar_med = float(sector_data['psi_barMed'].iloc[0])
        psi_bar_min = float(sector_data['psi_barMin'].iloc[0])
        theta = float(sector_data['teta'].iloc[0])
        num_periods = len(all_years)

        # Generate employment factor matrices
        A_max = generate_employment_factor(psi_bar_max, d, num_periods, steepness_factor, correction_factor)
        A_med = generate_employment_factor(psi_bar_med, d, num_periods, steepness_factor, correction_factor)
        A_min = generate_employment_factor(psi_bar_min, d, num_periods, steepness_factor, correction_factor)

        # Calculate employment vectors
        scenarios = [
            ('Constant Current', constant_current),
            ('Raging Storm', raging_storm),
            ('Rising Tide', rising_tide),
            ('Shifting Winds', shifting_winds),
            ('Government Plan', government_plan),
        ]
        y_max, y_med, y_min = calculate_employment_vectors(scenarios, A_max, A_med, A_min)

        # Prepare data for exporting with all years
        total_columns = 2 + (5 * column_offset)
        output_cell = prepare_output_table(all_years, y_max, y_med, y_min, total_columns, column_offset)

        # Write the result to an Excel file with the sector name in the file name
        os.makedirs(output_file_path, exist_ok=True)
        sector_output_file_path = os.path.join(output_file_path, f'{sheet}_scenarios.xlsx')
        pd.DataFrame(output_cell).to_excel(sector_output_file_path, header=False, index=False)

        # Filter data for plotting (only 2025-2040) and plot
        plot_year_indices = np.where((all_years >= 2025) & (all_years <= 2040))[0]
        plot_years = all_years[plot_year_indices]

        # Plot scenarios together (excluding government)
        plot_employment_vectors(plot_years, y_max[:4], y_med[:4], y_min[:4], plot_year_indices,
                                [name for name, _ in scenarios[:4]], sheet)

        # Plot government scenario separately
        if len(y_max) >= 5:  # Check that government data exists
            plot_government_vector(plot_years, y_max[4], y_med[4], y_min[4], plot_year_indices,
                                   'Government Plan', sheet)

    plt.show()
    download_results()


if __name__ == '__main__':
    main()

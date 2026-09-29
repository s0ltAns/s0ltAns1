# =============================================================================
# STEP 2 - STEEL, ALUMINIUM, CEMENT (from the aluminium/cement/steel MATLAB script)
# Arctan employment factor, Max/Med/Min psi, 4 scenarios + government plan.
# Writes: Output_data/Sectors/<Sector>_scenarios.xlsx
#         raw_data.xlsx sheets: aluminium, cement, steel
# =============================================================================
plt.close('all')

sector_steepness_factor = 10
sector_correction_factor = 2800 / 4393
sector_column_offset = 5  # Offset between scenarios in the output file
SECTORS_OUTPUT_DIR = os.path.join(OUTPUT_DIR, 'Sectors')

# (input sheet, scenario range, sector parameter range, raw_data sheet)
BASIC_SECTORS = [
    ('Aluminium', 'A1:G32', 'H1:L2', 'aluminium'),
    ('Cement',    'A1:G36', 'H1:L2', 'cement'),
    ('Steel',     'A1:G37', 'H1:L2', 'steel'),
]

SCENARIO_COLORS_4 = [
    (60 / 255, 106 / 255, 255 / 255),   # Blue for Constant Current
    (214 / 255, 69 / 255, 69 / 255),    # Red for Raging Storm
    (13 / 255, 191 / 255, 96 / 255),    # Green for Rising Tide
    (233 / 255, 162 / 255, 126 / 255),  # Orange for Shifting Winds
]
SCENARIO_NAMES_4 = ['Constant Current', 'Raging Storm', 'Rising Tide', 'Shifting Winds']


def arctan_employment_factor(psi, num_periods):
    """A(row, col) = psi * correction_factor * atan(steepness * (row - col + 1)) on and below the diagonal."""
    A = np.zeros((num_periods, num_periods))
    for row in range(1, num_periods + 1):
        for col in range(1, row + 1):
            t_shifted = row - col + 1
            A[row - 1, col - 1] = psi * sector_correction_factor * np.arctan(sector_steepness_factor * t_shifted)
    return A


def scenario_output_table(all_years, y_max, y_med, y_min, total_columns):
    """Time Period | Year | Max/Med/Min_Employment_S1 | 2 blank | ..._S2 | ... (as the MATLAB export)."""
    n = len(all_years)
    table = [[None] * total_columns for _ in range(n + 1)]
    table[0][0], table[0][1] = 'Time Period', 'Year'
    for r in range(n):
        table[r + 1][0] = r + 1
        table[r + 1][1] = all_years[r]
    for s in range(len(y_max)):
        col_start = 2 + s * sector_column_offset
        for j, (label, data) in enumerate(zip(['Max', 'Med', 'Min'], (y_max[s], y_med[s], y_min[s]))):
            table[0][col_start + j] = f'{label}_Employment_S{s + 1}'
            for r in range(n):
                table[r + 1][col_start + j] = data[r]
    return table


for sheet, scenario_range, parameter_range, raw_sheet in BASIC_SECTORS:
    scenarios_data = read_table(INPUT_FILE, sheet, scenario_range)
    all_years = scenarios_data['Year'].to_numpy()
    scenario_vectors = [scenarios_data[c].to_numpy() for c in ['x_s1', 'x_s2', 'x_s3', 'x_s4', 'x_gov']]

    sector_data = read_table(INPUT_FILE, sheet, parameter_range)
    psi_levels = [sector_data[c].iloc[0] for c in ['psi_barMax', 'psi_barMed', 'psi_barMin']]
    num_periods = len(all_years)

    A_max, A_med, A_min = (arctan_employment_factor(psi, num_periods) for psi in psi_levels)
    y_max = [A_max @ x for x in scenario_vectors]
    y_med = [A_med @ x for x in scenario_vectors]
    y_min = [A_min @ x for x in scenario_vectors]

    # Export (all years), then add the sheet to raw_data.xlsx
    table = scenario_output_table(all_years, y_max, y_med, y_min, 2 + 5 * sector_column_offset)
    write_cells(os.path.join(SECTORS_OUTPUT_DIR, f'{sheet}_scenarios.xlsx'), 'Sheet1', table)
    write_raw_data_sheet(raw_sheet, raw_rows_from_output_table(table))

    # Charts for 2025-2040
    plot_year_indices = np.where((all_years >= 2025) & (all_years <= 2040))[0]
    plot_years = all_years[plot_year_indices]
    plot_scenario_bands(plot_years, y_max[:4], y_med[:4], y_min[:4], plot_year_indices, SCENARIO_NAMES_4,
                        SCENARIO_COLORS_4, f'Scenarios for {sheet} sector', xticks=range(2025, 2041, 5))
    plot_scenario_bands(plot_years, y_max[4:], y_med[4:], y_min[4:], plot_year_indices, ['Government Plan'],
                        [(128 / 255, 128 / 255, 128 / 255)], f'Government Plan for {sheet} sector',
                        integer_ticks=True, xticks=range(2025, 2041, 5))

plt.show()

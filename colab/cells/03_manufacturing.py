# =============================================================================
# STEP 3 - WIND TURBINE, PV AND ELECTROLYSER MANUFACTURING
# (from PV_Wind_Elec_Manufacturing.m)
# Arctan employment factor, Max/Med/Min psi, 4 scenarios + 100% / 50% government plan.
# Writes: raw_data.xlsx sheets: WindMan, PVman, ElecMan
#         (the MATLAB script's own Excel export is switched off, so there is no other file)
# Uses arctan_employment_factor and scenario_output_table from STEP 2 - run STEP 2 first.
# =============================================================================
plt.close('all')

# (input sheet, scenario range, sector parameter range, raw_data sheet)
# Note: the first two sheet names end with a space, as in the MATLAB script.
MANUFACTURING_SECTORS = [
    ('Wind turbine Manufacturing ', 'A1:H37', 'I1:M2', 'WindMan'),
    ('PV Manufacturing ',           'A1:H36', 'I1:M2', 'PVman'),
    ('Hydrogen Production',         'A1:H38', 'I1:M2', 'ElecMan'),  # Electrolyser manufacturing
]
MANUFACTURING_COLUMN_COUNT = 2 + 6 * sector_column_offset
GOVERNMENT_COLORS = [(128 / 255, 128 / 255, 128 / 255), (192 / 255, 192 / 255, 192 / 255)]  # Gray, light gray

for sheet, scenario_range, parameter_range, raw_sheet in MANUFACTURING_SECTORS:
    scenarios_data = read_table(INPUT_FILE, sheet, scenario_range)
    all_years = scenarios_data['Year'].to_numpy()
    scenario_vectors = [scenarios_data[c].to_numpy() for c in ['x_s1', 'x_s2', 'x_s3', 'x_s4']]
    government_vectors = [scenarios_data[c].to_numpy() for c in ['x100_gov', 'x50_gov']]

    sector_data = read_table(INPUT_FILE, sheet, parameter_range)
    psi_levels = [sector_data[c].iloc[0] for c in ['psi_barMax', 'psi_barMed', 'psi_barMin']]
    num_periods = len(all_years)

    A_max, A_med, A_min = (arctan_employment_factor(psi, num_periods) for psi in psi_levels)
    y_max, y_med, y_min = ([A @ x for x in scenario_vectors] for A in (A_max, A_med, A_min))
    gov_max, gov_med, gov_min = ([A @ x for x in government_vectors] for A in (A_max, A_med, A_min))

    # The 4 scenarios go to raw_data.xlsx (government plans are not exported, as in MATLAB)
    table = scenario_output_table(all_years, y_max, y_med, y_min, MANUFACTURING_COLUMN_COUNT)
    write_raw_data_sheet(raw_sheet, raw_rows_from_output_table(table))

    # Charts for 2025-2040
    plot_year_indices = np.where((all_years >= 2025) & (all_years <= 2040))[0]
    plot_years = all_years[plot_year_indices]
    plot_scenario_bands(plot_years, y_max, y_med, y_min, plot_year_indices, SCENARIO_NAMES_4,
                        SCENARIO_COLORS_4, f'Employment Projections for {sheet}', integer_ticks=True)
    plot_scenario_bands(plot_years, gov_max, gov_med, gov_min, plot_year_indices,
                        ['100% Government Plan', '50% Government Plan'], GOVERNMENT_COLORS,
                        f'Government Employment Projections for {sheet}')

plt.show()

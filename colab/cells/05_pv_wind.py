# =============================================================================
# STEP 5 - SOLAR PV AND WIND POWER
# (Planning.m, Construction.m, OperationAndMaintenance.m, Solar_plots.m, Wind_plots.m)
# Part A: employment by job role -> PV_Wind_Power/employment_per_jobroles/Output_data
# Part B: employment by phase + charts -> PV_Wind_Power/step1 and PV_Wind_Power/scenarios_data
# Writes raw_data.xlsx sheets: PV (Solar), Wind
# =============================================================================
plt.close('all')

PV_WIND_DIR = os.path.join(OUTPUT_DIR, 'PV_Wind_Power')
JOBROLES_DIR = os.path.join(PV_WIND_DIR, 'employment_per_jobroles', 'Output_data')  # Part A
STEP1_DIR = os.path.join(PV_WIND_DIR, 'step1')                                      # Part B, steps B1-B2
SCENARIOS_DATA_DIR = os.path.join(PV_WIND_DIR, 'scenarios_data')                    # Part B, steps B3-B4


# =============================================================================
# Model inputs
# =============================================================================

# --- Planning & construction curve: sine rise -> flat peak -> cosine decline ---
q = 0.0        # The height of the curve's peak
k = 0.49       # Power of sine in the first part (rate of increase)
l = 1          # Power of cosine in the third part (rate of decrease)
theta = 1.009  # Scaling parameter

# --- Operation & maintenance curve: arctan (rises quickly, then stays flat) ---
steepness_factor = 10
correction_factor = 2800 / 4393

DURATION_CELL = 'I2:I2'  # Project duration d, on the Solar and Wind sheets

# --- Part A: employment by job role -----------------------------------------
# (scenario name, input column); rows 2-38
JOBROLE_SCENARIOS = [
    ('x_s1', 'C'), ('x_s2', 'D'), ('x_s3', 'E'), ('x_s4', 'F'),
    ('x_gov', 'G'),
    ('x_gov_new', 'G'),  # As in the MATLAB scripts, this reads column G again (same as x_gov)
]
JOBROLE_ROWS = (2, 38)
JOBROLE_YEARS_RANGE = 'A2:A38'
JOBROLE_OM_OFFSET = -0.0011  # Only OperationAndMaintenance.m subtracts this

# (output sheet name, curve type, job-role table range per technology)
JOBROLE_PHASES = [
    ('Planning',                  'plc', {'Solar': 'M2:Q25',  'Wind': 'M2:Q13'}),   # Planning.m
    ('Construction',              'plc', {'Solar': 'M26:Q64', 'Wind': 'M14:Q38'}),  # Construction.m
    ('Operation and Maintenance', 'om',  {'Solar': 'M65:Q77', 'Wind': 'M39:Q54'}),  # OperationAndMaintenance.m
]

# --- Part B: employment by phase + charts -----------------------------------
# (scenario name, input column); rows 2 to the technology's last_row
PHASE_SCENARIOS = [
    ('x_s1', 'C'), ('x_s2', 'D'), ('x_s3', 'E'), ('x_s4', 'F'),
    ('x_gov_sm', 'G'),  # Government plan, smoothed
    ('x_gov_sc', 'H'),  # Government plan, scattered
]
PLOT_SCENARIOS = ['x_s1', 'x_s2', 'x_s3', 'x_s4', 'x_gov_sc', 'x_gov_sm']
LEGEND_NAMES = ['Constant Current', 'Raging Storm', 'Rising Tide', 'Shifting Winds', 'Scattered', 'Smoothed']
STAGE_NAMES = ['Project Planning and Development', 'Construction', 'Operation and Maintenance']
STAGE_COLORS = [(74 / 255, 111 / 255, 165 / 255), (209 / 255, 86 / 255, 86 / 255), (214 / 255, 160 / 255, 118 / 255)]


def _rgb(values):
    return [tuple(v / 255 for v in rgb) for rgb in values]


# Where Solar_plots.m and Wind_plots.m differ
TECHNOLOGIES = {
    'Solar': dict(                       # Solar_plots.m
        raw_sheet='PV',                  # Sheet in raw_data.xlsx
        last_row=37,                     # Investment in C2:C37 ... H2:H37 (36 half-year periods)
        plc_job_range='T2:X3',           # GOV ('T9:X10' for the scenarios)
        om_job_range='T4:X4',            # GOV ('T11:X11' for the scenarios)
        # Solar_plots.m names the planning & construction sheets with hyphens
        plc_sheet_names=['x-s1', 'x-s2', 'x-s3', 'x-s4', 'x-gov-sm', 'x-gov-sc'],
        om_range_for_totals='B7:C37',    # O&M rows used for the scenario charts (B3)
        stretch_to_even_years=True,      # B3: 2025:0.5:2040 made 32 long; data stretched by interpolation
        scenario_colors=_rgb([(85, 168, 104), (244, 123, 32), (72, 133, 237),
                              (219, 68, 55), (171, 71, 188), (66, 133, 244)]),
        title_main='Employment Over Time for Scenarios x-s1 to x-s4',
        title_gov='Employment Over Time for Scenarios x-gov-sc and x-gov-sm',
    ),
    'Wind': dict(                        # Wind_plots.m
        raw_sheet='Wind',                # Sheet in raw_data.xlsx
        last_row=32,                     # Investment in C2:C32 ... H2:H32 (31 half-year periods)
        plc_job_range='M55:Q56',         # gov ('T2:X3' for the scenarios)
        om_job_range='M57:Q57',          # Gov ('T3:X3' for the scenarios)
        plc_sheet_names=['x_s1', 'x_s2', 'x_s3', 'x_s4', 'x_gov_sm', 'x_gov_sc'],
        om_range_for_totals='B2:C32',
        stretch_to_even_years=False,
        scenario_colors=_rgb([(85, 168, 104), (244, 123, 32), (72, 133, 237),
                              (219, 68, 55), (85, 168, 104), (66, 133, 244)]),
        title_main='Employment Over Time for Scenarios x_s1 to x_s4',
        title_gov='Employment Over Time for Scenarios x_gov_sc and x_gov_sm',
    ),
}
PLC_RANGE_FOR_CHARTS = 'B7:C37'  # Planning (column B) and construction (column C) rows used in B3 and B4
OM_RANGE_FOR_STACKED = 'B2:B32'  # O&M rows used in B4


# =============================================================================
# Model formulas
# =============================================================================
def convert_start_length(start, length, d, early_share):
    """Turn a job's start (B/M/E) and length (S/M/L) into a time window.

    early_share is 0.2 in Planning.m / Construction.m and 0.25 in Solar_plots.m / Wind_plots.m.
    """
    start_time = end_time = None
    if start == 'B':    # Beginning
        start_time = 0
        end_time = {'S': early_share * d, 'M': 0.8 * d, 'L': d}.get(length)
    elif start == 'M':  # Middle
        start_time = early_share * d
        end_time = {'S': 0.8 * d, 'M': d}.get(length)
    elif start == 'E':  # End
        start_time = 0.8 * d
        end_time = {'S': d}.get(length)
    if start_time is None or end_time is None:
        # MATLAB stops with "Output argument not assigned" for these combinations
        raise ValueError(f'Unsupported start/length combination: {start!r}/{length!r}')
    return start_time, end_time


def plc_factor_matrix(psi, start_time, end_time, d, num_periods, formula):
    """Employment factor matrix A for planning & construction jobs.

    formula='jobroles' is the version in Planning.m / Construction.m and
    formula='plots' the version in Solar_plots.m / Wind_plots.m. They differ
    only in the peak and decline pieces; with q = 0 and l = 1 both give the same numbers.
    """
    A = np.zeros((num_periods, num_periods))
    for row in range(1, num_periods + 1):
        for col in range(1, row + 1):
            t = row - col + 1  # Periods since the investment was made
            if not (start_time <= t <= end_time):
                continue
            if t <= d / 3:  # Rise
                value = theta * psi * np.sin((3 * np.pi / (2 * d)) * t) ** k
            elif t <= 2 * d / 3:  # Peak
                bump = q * np.sin((3 * np.pi / d) * (t - d / 3))
                if formula == 'jobroles':
                    bump = abs(bump)
                value = bump + theta * psi
            elif t <= d:  # Decline
                cosine = np.cos((3 * np.pi / d) * (t - 2 * d / 3))
                if formula == 'jobroles':
                    value = (psi * (theta - 1) / 2) * cosine ** l + psi * (1 + theta) / 2
                else:
                    value = (psi * (theta - 1) / 2 * cosine) ** l + psi * (1 + theta) / 2
            else:
                continue
            A[row - 1, col - 1] = value
    return A


def om_factor_matrix(psi, start_time, num_periods, offset):
    """Employment factor matrix A for operation & maintenance jobs (arctan curve).

    offset is -0.0011 in OperationAndMaintenance.m and 0 in Solar_plots.m / Wind_plots.m.
    """
    A = np.zeros((num_periods, num_periods))
    for row in range(1, num_periods + 1):
        for col in range(1, row + 1):
            t = row - col + 1 - start_time
            if t >= 0:
                A[row - 1, col - 1] = psi * correction_factor * np.arctan(steepness_factor * t) + offset
    return A


# =============================================================================
# Part A: employment by job role
#    (Planning.m, Construction.m, OperationAndMaintenance.m)
#    Output: <Tech>_JobRoles_Scenarios.xlsx with sheets Planning, Construction,
#    Operation and Maintenance. In each sheet the six scenarios are stacked
#    vertically, one column per job role.
# =============================================================================
def run_jobrole_phase(tech, phase_name, curve, job_range):
    jobs = read_job_table(INPUT_FILE, tech, job_range)
    first_row, last_row = JOBROLE_ROWS
    scenarios = [(name, read_column(INPUT_FILE, tech, f'{col}{first_row}:{col}{last_row}'))
                 for name, col in JOBROLE_SCENARIOS]

    if curve == 'plc':
        d = read_numbers(INPUT_FILE, tech, DURATION_CELL)[0, 0]
        years = list(read_column(INPUT_FILE, tech, JOBROLE_YEARS_RANGE))  # Years from the input file
    else:
        # O&M: half-year periods labelled 2025, blank, 2026, blank, ...
        num_periods = len(scenarios[0][1])
        years = [2025 + i // 2 if i % 2 == 0 else None for i in range(num_periods)]
    num_periods = len(years)

    all_rows = []
    for scenario_name, x in scenarios:
        # Header row, then one row per period: Scenario | Time Period | Year | job 1 | job 2 | ...
        block = [['Scenario', 'Time Period', 'Year'] + [job.name for job in jobs]]
        block += [[scenario_name, p + 1, years[p]] + [None] * len(jobs) for p in range(num_periods)]

        for j, job in enumerate(jobs):
            if curve == 'plc':
                if job.psi == 0:
                    continue  # Left blank
                start_time, end_time = convert_start_length(job.start, job.length, d, early_share=0.2)
                A = plc_factor_matrix(job.psi, start_time, end_time, d, num_periods, formula='jobroles')
            else:
                A = om_factor_matrix(job.psi, 0, num_periods, offset=JOBROLE_OM_OFFSET)
            y = A @ x
            for p in range(num_periods):
                block[p + 1][3 + j] = y[p]

        all_rows += block

    output_file = os.path.join(JOBROLES_DIR, f'{tech}_JobRoles_Scenarios.xlsx')
    write_cells(output_file, phase_name, all_rows)
    print(f'{tech} - {phase_name}: employment vectors for all scenarios exported to Excel.')


# =============================================================================
# Part B: employment by phase + charts (Solar_plots.m, Wind_plots.m)
#    B1  Planning & construction totals -> step1/<Tech>_Employment_Output_PLC.xlsx
#    B2  Operation & maintenance totals -> step1/<Tech>_Employment_Output_OM.xlsx
#    B3  Scenario charts with +/-20% bands -> scenarios_data/<Tech>_Scenarios_Plot_Data.xlsx
#    B4  Stacked area chart per scenario  -> scenarios_data/<Tech>_Stacked_Plot_Data.xlsx
# =============================================================================
def step_b1_planning_construction(tech, cfg, x, plc_file):
    jobs = read_job_table(INPUT_FILE, tech, cfg['plc_job_range'])
    d = read_numbers(INPUT_FILE, tech, DURATION_CELL)[0, 0]
    num_periods = len(x['x_s1'])
    print(f'{tech} PLC - Number of Jobs: {len(jobs)}')
    print(f'{tech} PLC - Number of Periods: {num_periods}')

    Y = {name: np.zeros((num_periods, len(jobs))) for name in x}
    A_matrices = [None] * len(jobs)
    for i, job in enumerate(jobs):
        if job.psi == 0:
            continue  # Skip if no employment factor
        start_time, end_time = convert_start_length(job.start, job.length, d, early_share=0.25)
        A = plc_factor_matrix(job.psi, start_time, end_time, d, num_periods, formula='plots')
        A_matrices[i] = A
        for name in Y:
            Y[name][:, i] = A @ x[name]

    # One sheet per scenario: period numbers in column A, one column per job
    header = [None] + [job.name for job in jobs]
    for (name, _), sheet_name in zip(PHASE_SCENARIOS, cfg['plc_sheet_names']):
        rows = [header] + [[p + 1] + list(Y[name][p]) for p in range(num_periods)]
        write_cells(plc_file, sheet_name, rows)

    # Each job's A matrix on its own sheet (sheet names cleaned, max 31 characters)
    for job, A in zip(jobs, A_matrices):
        if A is not None:
            sheet_name = re.sub(r'[:\\/*?\[\]]', '', job.name)[:31]
            write_cells(plc_file, sheet_name, A.tolist())
    print(f'{tech} PLC employment vectors and factor matrices exported.')


def step_b2_operation_maintenance(tech, cfg, x, om_file):
    jobs = read_job_table(INPUT_FILE, tech, cfg['om_job_range'])
    num_periods = len(x['x_gov_sm'])
    print(f'{tech} OM - Number of Jobs: {len(jobs)}')
    print(f'{tech} OM - Number of Periods: {num_periods}')

    Y = {name: np.zeros((num_periods, len(jobs))) for name in x}
    for i, job in enumerate(jobs):
        A = om_factor_matrix(job.psi, 0, num_periods, offset=0.0)  # For OM, start time is zero
        for name in Y:
            Y[name][:, i] = A @ x[name]

    header = [None] + [job.name for job in jobs]
    for name, _ in PHASE_SCENARIOS:
        rows = [header] + [[p + 1] + list(Y[name][p]) for p in range(num_periods)]
        write_cells(om_file, name, rows)
    print(f'{tech} OM employment vectors exported.')


def _fill_missing_with_zero(values):
    return np.where(np.isnan(values), 0.0, values)  # MATLAB fillmissing(..., 'constant', 0)


def _stretch(values, length):
    """MATLAB interp1(linspace(1, length, numel(values)), values, 1:length, 'linear', 'extrap')."""
    if len(values) == length:
        return values
    return np.interp(np.arange(1, length + 1), np.linspace(1, length, len(values)), values)


def step_b3_scenario_charts(tech, cfg, plc_file, om_file):
    years = np.arange(2025, 2040.25, 0.5)  # 2025:0.5:2040 -> 31 values
    if cfg['stretch_to_even_years'] and len(years) % 2 != 0:
        years = np.append(years, years[-1] + 0.5)

    plc_sheet = dict(zip([name for name, _ in PHASE_SCENARIOS], cfg['plc_sheet_names']))
    total_employment = np.zeros((len(years), len(PLOT_SCENARIOS)))
    for i, scenario in enumerate(PLOT_SCENARIOS):
        data_plc = _fill_missing_with_zero(read_numbers(plc_file, plc_sheet[scenario], PLC_RANGE_FOR_CHARTS))
        data_om = _fill_missing_with_zero(read_numbers(om_file, scenario, cfg['om_range_for_totals']))
        planning, construction, op_maint = data_plc[:, 0], data_plc[:, 1], data_om[:, 0]
        if cfg['stretch_to_even_years']:
            planning, construction, op_maint = (_stretch(v, len(years)) for v in (planning, construction, op_maint))
        total_employment[:, i] = planning + construction + op_maint

    if cfg['stretch_to_even_years'] and np.isnan(total_employment).any():
        raise ValueError('total_employment contains NaN values. Check the input data files or ranges.')

    upper_bounds = total_employment * 1.2
    lower_bounds = total_employment * 0.8

    # raw_data.xlsx: the four scenarios (Max = upper bound, Med = total, Min = lower bound)
    write_raw_data_sheet(cfg['raw_sheet'], raw_rows_by_year(
        years, [(upper_bounds[:, i], total_employment[:, i], lower_bounds[:, i]) for i in range(4)]))

    plot_file = os.path.join(SCENARIOS_DATA_DIR, f'{tech}_Scenarios_Plot_Data.xlsx')
    for indices, title, sheet in [([0, 1, 2, 3], cfg['title_main'], 'Scenarios'),
                                  ([4, 5], cfg['title_gov'], 'GOV_Scenarios')]:
        plot_scenario_bands(years, total_employment, upper_bounds, lower_bounds, indices,
                            cfg['scenario_colors'], title)
        # Export plot data: Year, then Max / Med / Min for each scenario
        header = ['Year']
        for i in indices:
            header += [f'Max_{PLOT_SCENARIOS[i]}', f'Med_{PLOT_SCENARIOS[i]}', f'Min_{PLOT_SCENARIOS[i]}']
        rows = [[years[r]] + [v for i in indices
                              for v in (upper_bounds[r, i], total_employment[r, i], lower_bounds[r, i])]
                for r in range(len(years))]
        write_cells(plot_file, sheet, [header] + rows)


def step_b4_stacked_charts(tech, cfg, plc_file, om_file):
    years = 2025 + np.arange(31) * 0.5  # 2025:0.5:2040
    plc_sheet = dict(zip([name for name, _ in PHASE_SCENARIOS], cfg['plc_sheet_names']))
    stack_file = os.path.join(SCENARIOS_DATA_DIR, f'{tech}_Stacked_Plot_Data.xlsx')

    for scenario in PLOT_SCENARIOS:
        data_plc = read_numbers(plc_file, plc_sheet[scenario], PLC_RANGE_FOR_CHARTS)
        data_om = read_numbers(om_file, scenario, OM_RANGE_FOR_STACKED)
        planning, construction, op_maint = data_plc[:, 0], data_plc[:, 1], data_om[:, 0]
        if not (len(planning) == len(construction) == len(op_maint) == len(years)):
            raise ValueError(f'Data length mismatch in scenario {scenario}. Verify data ranges.')

        plot_stacked(years, [planning, construction, op_maint], f'Employment Over Time for Scenario: {scenario}')

        header = ['Year', 'Project_Planning_Development', 'Construction', 'Operation_Maintenance']
        rows = [[years[r], planning[r], construction[r], op_maint[r]] for r in range(len(years))]
        write_cells(stack_file, scenario, [header] + rows)


def run_technology(tech, cfg):
    plc_file = os.path.join(STEP1_DIR, f'{tech}_Employment_Output_PLC.xlsx')
    om_file = os.path.join(STEP1_DIR, f'{tech}_Employment_Output_OM.xlsx')
    # Investment vectors for each scenario
    x = {name: read_column(INPUT_FILE, tech, f'{col}2:{col}{cfg["last_row"]}') for name, col in PHASE_SCENARIOS}

    step_b1_planning_construction(tech, cfg, x, plc_file)
    step_b2_operation_maintenance(tech, cfg, x, om_file)
    step_b3_scenario_charts(tech, cfg, plc_file, om_file)
    step_b4_stacked_charts(tech, cfg, plc_file, om_file)
    print(f'{tech}: plotting complete. All data exported.')


# --- Charts (MATLAB default figure size 560 x 420 px, axes at [0.13 0.3 0.8 0.6]) ---
def _new_figure():
    fig = plt.figure(figsize=(5.6, 4.2))
    ax = fig.add_axes([0.13, 0.3, 0.8, 0.6])
    return fig, ax


def _finish_axes(fig, ax, title, handles):
    ax.set_xlabel('Year')
    ax.set_ylabel('Number of Jobs')
    ax.set_title(title, fontsize=10, fontweight='bold')
    ax.set_xticks(range(2025, 2041, 5))
    ax.set_xlim(2025, 2040)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f'{v:.0f}'))  # No 10^x exponent
    ax.grid(True)
    # Legend below the axes without a box (2 columns so it fits the figure width)
    fig.legend(handles=handles, loc='upper center', bbox_to_anchor=(0.5, 0.17), ncol=2, frameon=False, fontsize=9)


def plot_scenario_bands(years, total, upper, lower, indices, colors, title):
    fig, ax = _new_figure()
    handles = []
    for i in indices:
        ax.fill_between(years, lower[:, i], upper[:, i], color=colors[i], alpha=0.3, linewidth=0)
        line, = ax.plot(years, total[:, i], color=colors[i], linewidth=1.5, label=LEGEND_NAMES[i])
        ax.plot(years, upper[:, i], '--', color=colors[i], linewidth=1)
        ax.plot(years, lower[:, i], '--', color=colors[i], linewidth=1)
        handles.append(line)
    _finish_axes(fig, ax, title, handles)


def plot_stacked(years, stages, title):
    fig, ax = _new_figure()
    areas = ax.stackplot(years, *stages, colors=STAGE_COLORS, labels=STAGE_NAMES, linewidth=0)
    _finish_axes(fig, ax, title, areas)


# =============================================================================
# Run Part A and Part B
# =============================================================================
# Part A: employment by job role (Planning, Construction, O&M) for Solar and Wind
for phase_name, curve, job_ranges in JOBROLE_PHASES:
    for tech in ('Solar', 'Wind'):
        run_jobrole_phase(tech, phase_name, curve, job_ranges[tech])

# Part B: employment by phase and charts, Solar then Wind
for tech, cfg in TECHNOLOGIES.items():
    run_technology(tech, cfg)
    plt.show()

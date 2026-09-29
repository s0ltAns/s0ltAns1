# =============================================================================
# STEP 4 - HYDROGEN SECTOR (from H2.m)
# Four subsectors (Solar, Wind, Electrolyser, Ammonia) with median/upper/lower psi.
# Writes: Output_data/Hydrogen/H2_Sector_New.xlsx and three .svg charts
#         raw_data.xlsx sheet: hydrogen
# =============================================================================
plt.close('all')

H2_OUTPUT_DIR = os.path.join(OUTPUT_DIR, 'Hydrogen')
H2_OUTPUT_FILE = os.path.join(H2_OUTPUT_DIR, 'H2_Sector_New.xlsx')
os.makedirs(H2_OUTPUT_DIR, exist_ok=True)

h2_sheet = 'Hydrogen As one Sector'
h2_time_range = 'B2:B33'   # Sets the number of periods
h2_start_year = 2024.5     # Each period is half a year
H2_INVESTMENT_RANGES = {
    'x_s1': 'C2:C33', 'x_s2': 'D2:D33', 'x_s3': 'E2:E33', 'x_s4': 'F2:F33',
    'x_gov': 'G2:G33',      # Government investment (smoothed)
    'x_gov_cur': 'H2:H33',  # Current government investment (scattered)
}
H2_SUBSECTORS = ['Solar', 'Wind', 'Electrolyser', 'Ammonia']
# psi cells: median (column U), upper (column V), lower (column T). GOV psi (P2, P3, P5, P6) not used.
H2_PSI_CELLS = {
    'median': {'Solar': 'U6:U6', 'Wind': 'U7:U7', 'Electrolyser': 'U9:U9', 'Ammonia': 'U10:U10'},
    'upper':  {'Solar': 'V6:V6', 'Wind': 'V7:V7', 'Electrolyser': 'V9:V9', 'Ammonia': 'V10:V10'},
    'lower':  {'Solar': 'T6:T6', 'Wind': 'T7:T7', 'Electrolyser': 'T9:T9', 'Ammonia': 'T10:T10'},
}
h2_d_map = {'Solar': 1, 'Wind': 1, 'Electrolyser': 1, 'Ammonia': 1}      # Example values, as in H2.m
h2_theta_map = {'Solar': 1, 'Wind': 1, 'Electrolyser': 1, 'Ammonia': 1}  # Example values, as in H2.m
h2_q, h2_k, h2_l = 0.2, 0.49, 1  # Peak height, power of sine (rise), power of cosine (decline)

H2_AREA_COLORS = [(209 / 255, 86 / 255, 86 / 255), (76 / 255, 175 / 255, 80 / 255),
                  (214 / 255, 160 / 255, 118 / 255), (165 / 255, 105 / 255, 189 / 255)]
H2_SCENARIOS = [  # (investment vector, name, colour, export name)
    ('x_s1', 'Constant Current', (0, 0.4470, 0.7410), 'ConstCurrent'),
    ('x_s2', 'Raging Storm', (0.8500, 0.3250, 0.0980), 'RagingStorm'),
    ('x_s3', 'Rising Tide', (0.9290, 0.6940, 0.1250), 'RisingTide'),
    ('x_s4', 'Shifting Winds', (0.4940, 0.1840, 0.5560), 'ShiftingWinds'),
]


def h2_employment_factor(psi_bar, d, num_periods, q, k, l, theta):
    """Sine rise, peak, cosine decline, then psi_bar after d periods."""
    A = np.zeros((num_periods, num_periods))
    for row in range(1, num_periods + 1):
        for col in range(1, row + 1):
            t_diff = row - col + 1
            if t_diff < d / 3:
                value = theta * psi_bar * abs(np.sin((3 * np.pi / (2 * d)) * t_diff)) ** k
            elif t_diff < 2 * d / 3:
                value = q * abs(np.sin((3 * np.pi / d) * (t_diff - d / 3))) + theta * psi_bar
            elif t_diff < d:
                value = (psi_bar * (theta - 1) / 2) * np.cos((3 * np.pi / d) * (t_diff - 2 * d / 3)) ** l \
                    + psi_bar * (1 + theta) / 2
            else:
                value = psi_bar  # After d steps, it settles at psi_bar
            A[row - 1, col - 1] = value
    return A


def h2_axes(ax, title, label_size):
    ax.set_title(title, fontsize=11, fontweight='bold')
    ax.set_xlabel('Years', fontsize=label_size)
    ax.set_ylabel('Number of Jobs', fontsize=label_size)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f'{v:.0f}'))
    ax.grid(True)


# --- Read data ---
h2_x = {name: read_column(INPUT_FILE, h2_sheet, rng) for name, rng in H2_INVESTMENT_RANGES.items()}
h2_num_periods = len(read_column(INPUT_FILE, h2_sheet, h2_time_range))
h2_psi = {level: {s: read_value(INPUT_FILE, h2_sheet, cell) for s, cell in cells.items()}
          for level, cells in H2_PSI_CELLS.items()}

h2_years = h2_start_year + np.arange(h2_num_periods) * 0.5
h2_full_year_indices = np.where(h2_years >= 2025)[0]  # Only show data from 2025 onward

h2_A = {level: {s: h2_employment_factor(h2_psi[level][s], h2_d_map[s], h2_num_periods,
                                        h2_q, h2_k, h2_l, h2_theta_map[s])
                for s in H2_SUBSECTORS}
        for level in H2_PSI_CELLS}


def h2_jobs_by_subsector(level, investment):
    return np.column_stack([h2_A[level][s] @ investment for s in H2_SUBSECTORS])


# --- Stacked area charts for the two government plans (median psi) ---
h2_gov_data = h2_jobs_by_subsector('median', h2_x['x_gov'])[h2_full_year_indices, :]
h2_gov_cur_data = h2_jobs_by_subsector('median', h2_x['x_gov_cur'])[h2_full_year_indices, :]
for data, title, svg_name in [(h2_gov_data, 'Hydrogen sector (smoothed inv.)', 'Employment_Distribution_x_gov.svg'),
                              (h2_gov_cur_data, 'Hydrogen sector (scattered inv.)',
                               'Employment_Distribution_x_gov_cur.svg')]:
    fig, ax = plt.subplots(figsize=(5.6, 4.2))
    areas = ax.stackplot(h2_years[h2_full_year_indices], data.T, colors=H2_AREA_COLORS, linewidth=0)
    ax.set_xlim(h2_years[h2_full_year_indices][0], h2_years[-1])
    h2_axes(ax, title, 10)
    ax.legend(areas, H2_SUBSECTORS, loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2,
              frameon=False, fontsize=10)
    fig.tight_layout()
    fig.savefig(os.path.join(H2_OUTPUT_DIR, svg_name))

# --- Total jobs per scenario with upper and lower bounds ---
h2_totals = {name: {level: h2_jobs_by_subsector(level, h2_x[name]).sum(axis=1)
                    for level in ('median', 'upper', 'lower')}
             for name, _, _, _ in H2_SCENARIOS}

fig, ax = plt.subplots(figsize=(5.6, 4.2))
handles = []
for name, label, color, _ in H2_SCENARIOS:
    t = h2_totals[name]
    ax.fill_between(h2_years, t['lower'], t['upper'], color=color, alpha=0.1, linewidth=0)
    line, = ax.plot(h2_years, t['median'], linewidth=2, color=color, label=label)
    ax.plot(h2_years, t['upper'], '--', linewidth=1, color=color)
    ax.plot(h2_years, t['lower'], '--', linewidth=1, color=color)
    handles.append(line)
h2_axes(ax, 'Hydrogen sector (scenarios)', 12)
ax.legend(handles=handles, loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2, frameon=False, fontsize=10)
ax.set_xlim(2025, 2040)
fig.tight_layout()
fig.savefig(os.path.join(H2_OUTPUT_DIR, 'Scenario_Line_Plot_with_Bounds.svg'))

# --- Export the chart data to H2_Sector_New.xlsx ---
for sheet_name, data in [('Raw_x_gov', h2_gov_data), ('Raw_x_gov_cur', h2_gov_cur_data)]:
    rows = [[year] + list(values) for year, values in zip(h2_years[h2_full_year_indices], data)]
    write_cells(H2_OUTPUT_FILE, sheet_name, [['Year'] + H2_SUBSECTORS] + rows)

header = ['Year']
for _, _, _, export_name in H2_SCENARIOS:
    header += [f'{export_name} Max', f'{export_name} Med', f'{export_name} Min']
rows = [[h2_years[p]] + [v for name, _, _, _ in H2_SCENARIOS
                         for v in (h2_totals[name]['upper'][p], h2_totals[name]['median'][p],
                                   h2_totals[name]['lower'][p])]
        for p in range(h2_num_periods)]
write_cells(H2_OUTPUT_FILE, 'Raw_Scenarios', [header] + rows)

# --- raw_data.xlsx: the same scenario data, on the 2022.5-2040 half-year grid ---
write_raw_data_sheet('hydrogen', raw_rows_by_year(
    h2_years, [(h2_totals[name]['upper'], h2_totals[name]['median'], h2_totals[name]['lower'])
               for name, _, _, _ in H2_SCENARIOS]))

plt.show()

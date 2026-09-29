# =============================================================================
# STEP 7 - FINAL REPORT FIGURES (from the "Python code for EF-plots" scripts)
# Uses the files written by Steps 2, 4 and 5, and ems.xlsx if it was uploaded.
# Writes to Output_data/Employment_Line_Charts:
#   Solar_Stacked_Smoothed.svg, Solar_Stacked_Scattered.svg        (solar_stacked.py)
#   Wind_Stacked_Smoothed_x_gov_sm.svg, ..._x_gov_sc.svg           (wind_stacked.py)
#   Hydrogen_Stacked_Smooth.svg, Hydrogen_Stacked_Scattered.svg    (Stacked_plots.py)
#   plots/steel_scenarios.svg                                      (Steel_*colours_scenarios.py)
#   ems.svg, ems_area_chart.svg                                    (BEE.py, stacked_ems.py)
# =============================================================================
from scipy.interpolate import make_interp_spline

plt.close('all')
os.makedirs(REPORT_DIR, exist_ok=True)

STEEL_COLOURS = 'bright'  # 'bright' (Steel_brightcolours_scenarios.py) or 'dark' (Steel_darkcolours_scenarios.py)

# Per-figure y-axis settings: {"new_max": top of the axis, "interval": tick spacing}
REPORT_YAXIS = {
    'Solar_Stacked_Smoothed': {"new_max": 16000, "interval": 4000},
    'Solar_Stacked_Scattered': {"new_max": 28000, "interval": 7000},
    'Wind_Stacked_Smoothed_x_gov_sm': {"new_max": 3000, "interval": 1000},
    'Wind_Stacked_Smoothed_x_gov_sc': {"new_max": 0, "interval": 1000},
    'Hydrogen Stacked (Smooth)': {"new_max": 30000, "interval": 8000},
    'Hydrogen Stacked (Scattered)': {"new_max": 3000, "interval": 8000},
    'ems': {"new_max": 16000, "interval": 4000},
}


def hundred_formatter(x, pos):
    """Tick labels floored to the nearest hundred, with thousands separators."""
    return f'{(int(x) // 100) * 100:,}'


def plot_report_stackplot(years, series, labels, colors, title, smooth, num_points, custom_yaxis):
    """Stacked area chart in the report format (3.2 x 1.5 in), saved as <title>.svg.

    Same as plot_stackplot in solar_stacked.py / wind_stacked.py / Stacked_plots.py.
    """
    series = [np.asarray(s) for s in series]
    if smooth:
        # Remove duplicate x-values, then cubic-spline onto a finer x-axis
        unique_years, unique_indices = np.unique(years, return_index=True)
        years = unique_years
        series = [s[unique_indices] for s in series]
        xnew = np.linspace(np.min(years), np.max(years), num_points)
        series = [make_interp_spline(years, s, k=3)(xnew) for s in series]
        years = xnew

    # Total jobs, rounded up to the next hundred (automatic y-axis)
    y_max_tick = np.ceil(np.max(np.sum(series, axis=0)) / 100) * 100

    plt.figure(figsize=(3.2, 1.5))
    ax = plt.gca()
    label_fontsize = 4
    tick_fontsize = 4

    ax.stackplot(years, *series, colors=colors, labels=labels)
    # ax.legend(loc='upper left', fontsize=label_fontsize)
    plt.xlabel('Year', fontsize=label_fontsize, labelpad=1)
    plt.ylabel('Number of Jobs', fontsize=label_fontsize, labelpad=1)

    ax.set_xlim(2025, 2040)
    plt.xticks(np.arange(2025, 2041, 5), fontsize=tick_fontsize)

    if custom_yaxis is not None and custom_yaxis.get("new_max") is not None \
            and custom_yaxis.get("interval") is not None:
        new_max = custom_yaxis["new_max"]
        yticks = [custom_yaxis["interval"] * i for i in range(5)]
    else:
        new_max = y_max_tick
        yticks = np.linspace(0, new_max, 5)
    ax.set_ylim(0, new_max)
    plt.yticks(yticks, fontsize=tick_fontsize)
    ax.yaxis.set_major_formatter(FuncFormatter(hundred_formatter))

    for spine in ax.spines.values():
        spine.set_linewidth(1.5)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_alpha(0.4)
    ax.spines['left'].set_alpha(0.4)

    plt.tight_layout()
    safe_title = title.replace(" ", "_").replace("(", "").replace(")", "").replace("-", "")
    plt.savefig(os.path.join(REPORT_DIR, f"{safe_title}.svg"), format='svg', bbox_inches='tight', transparent=True)
    plt.show()


def read_stacked_sheet(path, sheet, columns, group_by_year):
    """Read a stacked-chart sheet; years become whole numbers (half-years share a year)."""
    df = pd.read_excel(path, sheet_name=sheet)
    df = df.dropna(subset=['Year'])
    if group_by_year:
        df = df.groupby('Year', as_index=False).mean()
    df['Year'] = df['Year'].astype(int)
    # Empty cells count as 0 jobs. (On a fresh run the Wind planning & construction data has no
    # values after mid-2037 - see Step 5 - and the smoothing cannot handle empty cells.)
    return df['Year'].to_numpy(), [df[c].fillna(0).to_numpy() for c in columns]


# -----------------------------------------------------------------------------
# 1. Solar and wind: jobs by phase, government plans (solar_stacked.py, wind_stacked.py)
# -----------------------------------------------------------------------------
PHASE_COLUMNS = ['Project_Planning_Development', 'Construction', 'Operation_Maintenance']
PHASE_LABELS = ['Planning and development', 'Construction and installation', 'Operation and maintenance']
PHASE_COLORS = ['#008080', '#7495AB', '#E6E0D4']

for tech, sm_title, sc_title in [('Solar', 'Solar_Stacked_Smoothed', 'Solar_Stacked_Scattered'),
                                 ('Wind', 'Wind_Stacked_Smoothed_x_gov_sm', 'Wind_Stacked_Smoothed_x_gov_sc')]:
    stacked_file = os.path.join(SCENARIOS_DATA_DIR, f'{tech}_Stacked_Plot_Data.xlsx')
    for sheet, title, group_by_year in [('x_gov_sm', sm_title, True), ('x_gov_sc', sc_title, False)]:
        years, series = read_stacked_sheet(stacked_file, sheet, PHASE_COLUMNS, group_by_year)
        plot_report_stackplot(years, series, PHASE_LABELS, PHASE_COLORS, title,
                              smooth=True, num_points=35, custom_yaxis=REPORT_YAXIS[title])

# -----------------------------------------------------------------------------
# 2. Hydrogen: jobs by subsector, government plans (Stacked_plots.py)
# -----------------------------------------------------------------------------
H2_COLUMNS = ['Solar', 'Wind', 'Electrolyser', 'Ammonia']
H2_REPORT_COLORS = ['#E6E0D4', '#7C1B19', '#36454F', '#7C5B84']  # From bottom to top

for sheet, title, smooth, group_by_year in [('Raw_x_gov', 'Hydrogen Stacked (Smooth)', True, True),
                                            ('Raw_x_gov_cur', 'Hydrogen Stacked (Scattered)', False, False)]:
    years, series = read_stacked_sheet(H2_OUTPUT_FILE, sheet, H2_COLUMNS, group_by_year)
    plot_report_stackplot(years, series, H2_COLUMNS, H2_REPORT_COLORS, title,
                          smooth=smooth, num_points=300, custom_yaxis=REPORT_YAXIS[title])

# -----------------------------------------------------------------------------
# 3. Steel: four scenarios, large figure (Steel_brightcolours / Steel_darkcolours_scenarios.py)
# -----------------------------------------------------------------------------
df = pd.read_excel(os.path.join(SECTORS_OUTPUT_DIR, 'Steel_scenarios.xlsx'), skiprows=6, header=None,
                   engine='openpyxl')
df.columns = ['Time Period', 'Year',
              'Max_Employment_S1', 'Med_Employment_S1', 'Min_Employment_S1', 'Empty1', 'Empty2',
              'Max_Employment_S2', 'Med_Employment_S2', 'Min_Employment_S2', 'Empty3', 'Empty4',
              'Max_Employment_S3', 'Med_Employment_S3', 'Min_Employment_S3', 'Empty5', 'Empty6',
              'Max_Employment_S4', 'Med_Employment_S4', 'Min_Employment_S4', 'Empty7', 'Empty8',
              'Check1', 'Check2', 'Check3']
df = df[pd.to_numeric(df['Year'], errors='coerce').notna()]
df['Year'] = df['Year'].astype(int)
df = df[(df['Year'] >= 2025) & (df['Year'] <= 2040)]
cols_to_convert = [col for col in df.columns if 'Employment' in col]
df[cols_to_convert] = df[cols_to_convert].apply(pd.to_numeric, errors='coerce')
df = df.dropna(subset=cols_to_convert)
years = df['Year']

if STEEL_COLOURS == 'bright':
    steel_colors = ['#2C4FE8', '#BA3E43', '#0FA349', '#DD8956']
else:
    steel_colors = ['#152570', '#881620', '#074425', '#b37d54']

plt.figure(figsize=(12, 7))
ax = plt.gca()
for s, color in enumerate(steel_colors, start=1):
    med, min_, max_ = f'Med_Employment_S{s}', f'Min_Employment_S{s}', f'Max_Employment_S{s}'
    ax.fill_between(years, df[min_], df[max_], color=color, alpha=0.3)
    ax.plot(years, df[min_], linestyle='--', color=color, alpha=0.4, linewidth=1.1)
    ax.plot(years, df[max_], linestyle='--', color=color, alpha=0.4, linewidth=1.1)
    ax.plot(years, df[med], color=color, linewidth=1.5)

ax.set_xlim(2025, 2040)
ax.margins(x=0)
plt.xticks(range(2025, 2041, 3))

y_max_value = max(df[f'Max_Employment_S{s}'].max() for s in range(1, 5))
if STEEL_COLOURS == 'bright':
    # Start from zero with 5 ticks only
    plt.ylim(0, y_max_value)
    plt.yticks(np.linspace(0, y_max_value, 5))
    steel_ylabel = 'Number of Jobs Created'
else:
    # Ticks in multiples of 5,000
    y_max_tick = np.ceil(y_max_value / 5000) * 5000
    plt.ylim(0, y_max_tick)
    plt.yticks(np.arange(0, y_max_tick + 1, 5000))
    steel_ylabel = 'Number of Jobs'
ax.yaxis.set_major_formatter(FuncFormatter(lambda x, pos: f'{int(x):,}'))

ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['bottom'].set_alpha(0.4)
ax.spines['left'].set_alpha(0.4)
plt.xlabel('Year')
plt.ylabel(steel_ylabel)
plt.tight_layout()
plt.gcf().set_size_inches(7.39, 3.7)
os.makedirs(os.path.join(REPORT_DIR, 'plots'), exist_ok=True)  # Separate from Step 6's steel_scenarios.svg
plt.savefig(os.path.join(REPORT_DIR, 'plots', 'steel_scenarios.svg'), format="svg")
plt.show()

# -----------------------------------------------------------------------------
# 4. Building energy efficiency, government plan (BEE.py, stacked_ems.py) - needs ems.xlsx
# -----------------------------------------------------------------------------
if not os.path.exists(EMS_FILE):
    print('Skipped the Building Energy Efficiency figures: ems.xlsx was not uploaded in Step 1.')
else:
    df = pd.read_excel(EMS_FILE, sheet_name='Gov', engine='openpyxl')
    df = df[pd.to_numeric(df['Years'], errors='coerce').notna()]
    df['Years'] = df['Years'].astype(int)
    df[['Max_Gov', 'Med_Gov', 'Min_Gov']] = df[['Max_Gov', 'Med_Gov', 'Min_Gov']].apply(pd.to_numeric, errors='coerce')
    years = df['Years']
    ems_yaxis = REPORT_YAXIS['ems']
    ems_yticks = [ems_yaxis['interval'] * i for i in range(5)]

    def ems_axes(ax, tight_labels):
        ax.set_xlim(years.min(), years.max())
        plt.xticks(range(years.min(), years.max() + 1, 2), fontsize=4)
        plt.ylim(0, ems_yaxis['new_max'])
        plt.yticks(ems_yticks, fontsize=4)
        ax.yaxis.set_major_formatter(FuncFormatter(hundred_formatter))
        for spine in ax.spines.values():
            spine.set_linewidth(1)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_alpha(0.4)
        ax.spines['left'].set_alpha(0.4)
        labelpad = 1 if tight_labels else None
        plt.xlabel('Year', fontsize=4, labelpad=labelpad)
        plt.ylabel('Number of Jobs', fontsize=4, labelpad=labelpad)

    # BEE.py: min-max band with median line
    color = '#152570'
    plt.figure(figsize=(3.2, 1.5))
    ax = plt.gca()
    ax.fill_between(years, df['Min_Gov'], df['Max_Gov'], color=color, alpha=0.3)
    ax.plot(years, df['Min_Gov'], linestyle='--', color=color, alpha=0.4, linewidth=0.6)
    ax.plot(years, df['Max_Gov'], linestyle='--', color=color, alpha=0.4, linewidth=0.6)
    ax.plot(years, df['Med_Gov'], color=color, linewidth=0.6)
    ems_axes(ax, tight_labels=True)
    plt.tight_layout()
    plt.savefig(os.path.join(REPORT_DIR, 'ems.svg'), format="svg", transparent=True)
    plt.show()

    # stacked_ems.py: min, min-to-median and median-to-max as stacked layers
    plt.figure(figsize=(3.2, 1.5))
    ax = plt.gca()
    ax.stackplot(years, df['Min_Gov'], df['Med_Gov'] - df['Min_Gov'], df['Max_Gov'] - df['Med_Gov'],
                 colors=['#a1b5d9', '#7082b0', '#152570'])
    ems_axes(ax, tight_labels=False)
    plt.savefig(os.path.join(REPORT_DIR, 'ems_area_chart.svg'), format="svg")
    plt.show()

print(f'Report figures saved to {REPORT_DIR}')

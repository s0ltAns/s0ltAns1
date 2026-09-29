# =============================================================================
# STEP 6 - SCENARIO CHARTS FOR ALL SECTORS (reads raw_data.xlsx)
# One small chart per sector: median line, min/max band, dashed min/max lines.
# Writes: Output_data/svg_plots/<sheet>_scenarios.svg
# =============================================================================
plt.close('all')

SVG_DIR = os.path.join(OUTPUT_DIR, 'svg_plots')

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
chart_scenarios = {
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
chart_figsize = (3.2, 1.5)
label_fontsize = 4
tick_fontsize = 4
spine_linewidth = 1


def y_formatter(x, pos):
    """Y-axis tick labels with thousands separators, e.g. 21,000."""
    return f"{int(x):,}"


def load_sector_data(sheet):
    """Read one sector sheet and keep the rows for 2025-2040."""
    df = pd.read_excel(RAW_DATA_FILE, sheet_name=sheet, engine='openpyxl')

    # Ensure the 'Year' column is numeric and filter rows for 2025-2040
    df = df[pd.to_numeric(df['Year'], errors='coerce').notna()]
    df['Year'] = df['Year'].astype(int)
    df = df[(df['Year'] >= 2025) & (df['Year'] <= 2040)]

    # Convert employment scenario columns to numeric
    emp_cols = [col for col in df.columns if any(keyword in str(col) for keyword in
                ['ConstCurrent', 'RagingStorm', 'RisingTide', 'ShiftingWinds'])]
    df[emp_cols] = df[emp_cols].apply(pd.to_numeric, errors='coerce')
    return df


def y_axis_ticks(df, sheet):
    """Return (new_max, yticks): 4 equal steps from 0, rounded up to a round number."""
    if sheet in custom_yaxis_settings:
        new_max = custom_yaxis_settings[sheet]['new_max']
        interval = custom_yaxis_settings[sheet]['interval']
        return new_max, [interval * i for i in range(5)]

    y_max_value = max(df['ConstCurrent Max'].max(), df['RagingStorm Max'].max(),
                      df['RisingTide Max'].max(), df['ShiftingWinds Max'].max())
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

    plt.figure(figsize=chart_figsize)
    ax = plt.gca()

    for label, (med_col, min_col, max_col, color) in chart_scenarios.items():
        ax.fill_between(years, df[min_col], df[max_col], color=color, alpha=0.3)
        ax.plot(years, df[min_col], linestyle='--', color=color, alpha=0.4, linewidth=0.6)
        ax.plot(years, df[max_col], linestyle='--', color=color, alpha=0.4, linewidth=0.6)
        ax.plot(years, df[med_col], color=color, linewidth=0.6, label=label)

    # Legend (switched off, as in the original)
    # ax.legend(fontsize=tick_fontsize, loc='best')

    ax.set_xlim(2025, 2040)
    plt.xticks(range(2025, 2041, 3), fontsize=tick_fontsize, alpha=1)

    new_max, yticks = y_axis_ticks(df, sheet)
    plt.ylim(0, new_max)
    plt.yticks(yticks, fontsize=tick_fontsize, alpha=1)
    ax.yaxis.set_major_formatter(FuncFormatter(y_formatter))

    for spine in ax.spines.values():
        spine.set_linewidth(spine_linewidth)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_alpha(0.4)
    ax.spines['left'].set_alpha(0.4)

    plt.xlabel('Year', fontsize=label_fontsize, labelpad=1)
    plt.ylabel('Number of Jobs', fontsize=label_fontsize, labelpad=1)

    fig = plt.gcf()
    fig.set_size_inches(*chart_figsize)
    plt.tight_layout()
    plt.savefig(os.path.join(SVG_DIR, f"{sheet}_scenarios.svg"), format="svg", transparent=True)
    plt.show()
    plt.close()


os.makedirs(SVG_DIR, exist_ok=True)
available_sheets = load_workbook(RAW_DATA_FILE, read_only=True).sheetnames
for sector_name, sheet in sector_sheets.items():
    if sheet not in available_sheets:
        print(f'Skipped {sector_name}: raw_data.xlsx has no sheet "{sheet}"')
        continue
    plot_sector(load_sector_data(sheet), sheet)

print("Plots generated and saved for all sectors.")

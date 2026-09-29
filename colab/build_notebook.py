"""Build LIMA_employment_modelling.ipynb from the cell files in colab/cells.

Run after editing any cell file:  python colab/build_notebook.py
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))

CELLS = [
    ('01_setup.py',
     '## Step 1 - Setup\n'
     'Run this cell first. It asks you to upload **employment_modelling_input.xlsx**. '
     'Optionally select your previous **raw_data.xlsx** at the same time: any sheet the cells below do '
     'not recreate (e.g. *BuildingEnergyEfficiency*) is then kept.\n\n'
     'To work in Google Drive instead, set `USE_GOOGLE_DRIVE = True` and put the input workbook in '
     '`MyDrive/LIMA_Employment_modelling/Input_data/`.'),
    ('02_steel_aluminium_cement.py',
     '## Step 2 - Steel, aluminium and cement\n'
     'Adds the sheets *steel*, *aluminium* and *cement* to raw_data.xlsx.'),
    ('03_manufacturing.py',
     '## Step 3 - Wind turbine, PV and electrolyser manufacturing\n'
     'Adds the sheets *WindMan*, *PVman* and *ElecMan* to raw_data.xlsx. Run Step 2 first.'),
    ('04_hydrogen.py',
     '## Step 4 - Hydrogen sector\n'
     'Adds the sheet *hydrogen* to raw_data.xlsx.'),
    ('05_pv_wind.py',
     '## Step 5 - Solar PV and wind power\n'
     'Job roles (planning, construction, O&M) and phase totals. Adds the sheets *PV* and *Wind* to raw_data.xlsx.'),
    ('06_sector_charts.py',
     '## Step 6 - Scenario charts for all sectors\n'
     'Reads raw_data.xlsx and saves one .svg chart per sector.'),
    ('07_download.py',
     '## Step 7 - Download the results\n'
     'Downloads Output_data.zip with raw_data.xlsx, all Excel outputs and the charts.'),
]

INTRO = (
    '# LIMA employment modelling\n'
    'Run the cells from top to bottom (**Runtime > Run all** also works). '
    'Steps 2-5 each add their sectors to **raw_data.xlsx**; Step 6 draws the charts from it.'
)


def markdown(text):
    return {'cell_type': 'markdown', 'metadata': {}, 'source': text.splitlines(keepends=True)}


def code(text):
    return {'cell_type': 'code', 'execution_count': None, 'metadata': {}, 'outputs': [],
            'source': text.rstrip('\n').splitlines(keepends=True)}


cells = [markdown(INTRO)]
for filename, description in CELLS:
    with open(os.path.join(HERE, 'cells', filename), encoding='utf-8') as f:
        cells += [markdown(description), code(f.read())]

notebook = {
    'cells': cells,
    'metadata': {'colab': {'provenance': []},
                 'kernelspec': {'display_name': 'Python 3', 'name': 'python3'},
                 'language_info': {'name': 'python'}},
    'nbformat': 4,
    'nbformat_minor': 0,
}
with open(os.path.join(HERE, 'LIMA_employment_modelling.ipynb'), 'w', encoding='utf-8') as f:
    json.dump(notebook, f, indent=1, ensure_ascii=False)
    f.write('\n')
print('Wrote colab/LIMA_employment_modelling.ipynb')

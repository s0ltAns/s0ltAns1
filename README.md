# LIMA employment modelling

Python (Google Colab) version of the LIMA employment models.

## Run it in Google Colab
1. Download `colab/LIMA_employment_modelling.ipynb`.
2. Open https://colab.research.google.com, choose **File > Upload notebook** and select it.
3. Choose **Runtime > Run all**.
4. When Step 1 asks for files, select `employment_modelling_input.xlsx` (and, optionally,
   your previous `raw_data.xlsx` to keep sheets that no step recreates, such as BuildingEnergyEfficiency).
5. At the end, `Output_data.zip` is downloaded with `raw_data.xlsx`, all Excel outputs and the charts.

## Steps in the notebook
| Step | Model | raw_data.xlsx sheets |
|---|---|---|
| 1 | Setup: settings, upload, shared helpers | - |
| 2 | Steel, aluminium, cement | steel, aluminium, cement |
| 3 | Wind turbine, PV and electrolyser manufacturing | WindMan, PVman, ElecMan |
| 4 | Hydrogen sector (H2.m) | hydrogen |
| 5 | Solar PV and wind power (job roles, phases, charts) | PV, Wind |
| 6 | Scenario charts for all sectors, read from raw_data.xlsx | - |
| 7 | Download the results | - |

## Editing
The notebook is built from the files in `colab/cells/`. After editing a cell file, run
`python colab/build_notebook.py` to rebuild the notebook.

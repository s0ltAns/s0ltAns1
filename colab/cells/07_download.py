# =============================================================================
# STEP 7 - DOWNLOAD ALL RESULTS
# Zips the whole Output_data folder (raw_data.xlsx, all Excel outputs and charts).
# With USE_GOOGLE_DRIVE = True the results are already in your Drive.
# =============================================================================
if IN_COLAB and not USE_GOOGLE_DRIVE:
    zip_path = shutil.make_archive(os.path.join(WORK_DIR, 'Output_data'), 'zip', OUTPUT_DIR)
    files.download(zip_path)
else:
    print(f'Results are in {OUTPUT_DIR}')

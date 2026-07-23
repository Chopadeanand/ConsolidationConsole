# ConsolidationConsole

Streamlit front end for `merge_engine.py` (your original merge script, untouched),
styled to match the ConsolidationConsole mockup.

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```

## Files
- `app.py` — the UI: top bar, step indicator, and the 3 cards (master file,
  raw files, run & download), styled to match the mockup's colors/badges/dropzones.
- `merge_engine.py` — your original merge_month_and_fill script, called as-is via its `main()`.
- `.streamlit/config.toml` — sets the 1 GB upload limit and the blue/white theme used
  by native widgets (buttons, segmented control) so they match the brand colors.
- `requirements.txt` — dependencies.

## Notes
- Step 2's "Individual files" / "Zip archive" toggle is a real segmented control —
  pick one, and the matching uploader (multi-file for individual, single for zip)
  appears underneath, same as the mockup.
- The step badges and checklist in Step 3 update live as you complete each step.
- The "Run log" panel shows everything the script normally prints to the console
  (unmapped project codes, Vendor/Inhouse entries, date-parsing warnings, etc.).
- Uploads and the merged output live in a temp directory that's deleted after each run.
- To deploy on Streamlit Community Cloud: push this folder (including `.streamlit/`)
  to a GitHub repo and point Streamlit Cloud at `app.py`.

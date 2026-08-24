"""
merge_month_and_fill_v5.py
---------------------------
Drop-in replacement for merge_month_and_fill_v4.py. Same job as v4 --
merge a month's raw rating files into the existing consolidated master
workbook (Sheet1) and statically compute the derived columns so nothing
needs to auto-recalculate on open -- but now ALSO computes:

    AW  "Vendor/Inhouse"    -- looked up from Sheet5 (Asset Type +
                               Parameter -> Vendor/Inhouse), same table
                               used in the one-off manual pass.
    AX  "Invit/Non-Invit"   -- looked up from the PROJECT_INVIT_MAP
                               table below (Project Name -> status).
    AY  "Regional Head"     -- looked up from the
                               PROJECT_REGIONAL_HEAD_MAP table below
                               (Project Name -> Regional Head), hardcoded
                               the same way as Invit/Non-Invit.

for EVERY row -- both the rows already in the master and the newly
appended month's rows -- every single time you run this script. There
is no live formula anywhere for AW/AX/AY (or AA:AV): they are plain
values, computed once in Python, so the file opens instantly with
correct numbers already in place, no matter how many months you've
accumulated.

WHY THIS EXISTS
----------------
Doing this by hand each month (open file, write formula, wait for
481k+ rows to recalc, repeat for the next new column) is slow and
error-prone. This script makes each month's merge a single command.

USAGE
-----
    python merge_month_and_fill_v5.py <master.xlsx> <raw_month_folder> [output.xlsx]

<master.xlsx>        the consolidated workbook as of last month (e.g.
                      containing April+May+June data, with Sheet1 columns
                      A:AY already -- or even just A:AV/A:AX; see note below).
<raw_month_folder>   EITHER a folder containing that month's raw .csv/.xlsx
                      files (e.g. the 27 July files), OR a .zip archive
                      of them (they may be nested inside a sub-folder
                      within the zip, e.g. "July final/ALL_XXX...csv" --
                      that's handled automatically). Each raw file is
                      shaped like the originals (Project Name, Roadaid
                      ID, Category, Asset Type, Direction, Road Type,
                      Chainage, Parameter, Concern Raised, SPV Rating,
                      HO Rating, HO Remarks, Structure Name, Method,
                      Signboard Type, Signboard Classification,
                      Placement, Date Created, ...).
[output.xlsx]        optional; defaults to <master>_with_new_month.xlsx.

To just re-flatten/re-fix an existing file with NO new rows to add
(e.g. you edited Sheet5's Vendor/Inhouse table and want AW recomputed
everywhere), point <raw_month_folder> at an empty directory and it will
still rebuild AA:AY for every existing row.

NEXT MONTH
----------
Just point <master.xlsx> at THIS script's own output and give it the
new month's folder:

    python merge_month_and_fill_v5.py Consolidated_Apr_May_Jun_Jul.xlsx aug_files/ Consolidated_Apr_thru_Aug.xlsx

Each run's output becomes next month's input. The chain of months
keeps growing; nothing has to be recomputed by hand.

IF YOUR MASTER ONLY HAS COLUMNS A:AV OR A:AX (NO AY YET, OR NO AW/AX/AY YET)
-----------------------------------------------------------------------------
That's fine -- the script detects which of the header cells AW1/AX1/AY1
are missing and adds only those labels itself before writing data, so
you can point this straight at the original June-and-earlier (A:AV)
master, or at last month's (A:AX) master, and it will backfill
whichever of AW/AX/AY are missing for all the old rows too, in the same
run. The new header cells are given the same style as the other
headers, so they come out colored blue like the rest of row 1 instead
of plain/unstyled.

MAINTAINING THE PROJECT -> INVIT/NON-INVIT AND REGIONAL HEAD TABLES
----------------------------------------------------------------------
Unlike Vendor/Inhouse (which lives in the workbook's Sheet5 and so
updates automatically if you edit that sheet), Invit/Non-Invit and
Regional Head have no sheet of their own in the master -- each is a
fixed list of 27 project codes. Edit PROJECT_INVIT_MAP or
PROJECT_REGIONAL_HEAD_MAP below when a new project code shows up that
isn't in the relevant list yet; the script will warn you loudly (rather
than silently guessing) if it hits a project code with no entry in
either table.
"""

import sys, os, re, glob, zipfile, html, time, tempfile, shutil, atexit
import numpy as np
import pandas as pd


# ──────────────────────────────────────────────────────────────────────────
# UNIVERSAL Vendor/Inhouse reference table -- extracted from and verified
# against Consolidated_June26.xlsx (Sheet5 + cross-checked against the
# actual Vendor/Inhouse column of its Sheet1, 481,447 rows, 0 mismatches).
# This is the fallback/seed table used whenever the master workbook's own
# Sheet5 tab does not look like a genuine Asset Type/Parameter/Vendor-
# Inhouse table (e.g. the first time this script is pointed at a master
# that predates this table, or if Sheet5 ever gets overwritten with
# unrelated data again). Edit VENDOR_MAP_FALLBACK_ROWS below if a brand
# new Asset Type + Parameter combination shows up that isn't covered.
# ──────────────────────────────────────────────────────────────────────────
VENDOR_MAP_FALLBACK_ROWS = [
    ('Pavement', 'Cracks', 'Inhouse'),
    ('Kerb', 'Physical Condition', 'Vendor'),
    ('Kerb', 'Painting', 'Vendor'),
    ('Pavement', 'Pothole', 'Inhouse'),
    ('Pavement', 'Rutting', 'Vendor'),
    ('Kerb', 'Cleanliness', 'Inhouse'),
    ('Lightings', 'Cleanliness', 'Inhouse'),
    ('Lightings', 'Functional Condition', 'Inhouse'),
    ('Condition Of Clearance Of Vent', 'Clearance of vent', 'Vendor'),
    ('Lightings', 'Physical Condition', 'Inhouse'),
    ('ROW', 'Cleanliness of ROW', 'Inhouse'),
    ('PGR-Pedestrain Guardrail (PGR)', 'Painting and Cleanliness', 'Inhouse'),
    ('PGR-Pedestrain Guardrail (PGR)', 'Physical Condition', 'Vendor'),
    ('Hectometer Stones', 'Physical condition and functioning', 'Inhouse'),
    ('Hectometer Stones', 'Cleanliness', 'Inhouse'),
    ('MBCB-Semi Rigid Barrier', 'Physical Condition', 'Vendor'),
    ('MBCB-Semi Rigid Barrier', 'Missing end treatments and connections between barriers', 'Vendor'),
    ('MBCB-Semi Rigid Barrier', 'Cleanliness', 'Inhouse'),
    ('Shoulder', 'Edge Drop', 'Vendor'),
    ('Shoulder', 'Vegetation Growth', 'Inhouse'),
    ('Shoulder', 'Unevenness', 'Vendor'),
    ('PTZ', 'Functional Condition', 'Vendor'),
    ('Wearing Coat On Deck Slab', 'Potholes', 'Inhouse'),
    ('Wearing Coat On Deck Slab', 'Rutting', 'Vendor'),
    ('Wearing Coat On Deck Slab', 'Cracks (Alligator Cracks)', 'Inhouse'),
    ('Approach Settlements', 'Settlement', 'Vendor'),
    ('Rigid Crash Barriers', 'Cracks', 'Vendor'),
    ('Rigid Crash Barriers', 'Painting of RC crash barrier inner face - faded/not faded', 'Vendor'),
    ('Stagnation Of Rain Water', 'Stagnation of rain water on bridge deck', 'Inhouse'),
    ('Rigid Crash Barriers', 'Chipped concrete due to accidents', 'Vendor'),
    ('Quadrant Pitching', 'Physical condition of quadrant pitching', 'Vendor'),
    ('Operator Customized Keyboard', 'Functional Condition', 'Vendor'),
    ('Operator Customized Keyboard', 'Cleanliness', 'Inhouse'),
    ('Operator Customized Keyboard', 'Physical Condition', 'Vendor'),
    ('Incident Camera', 'Cleanliness', 'Inhouse'),
    ('Incident Camera', 'Functional Condition', 'Vendor'),
    ('Incident Camera', 'Physical Condition', 'Vendor'),
    ('Static Weigh Bridge (SWB)', 'Physical Condition', 'Vendor'),
    ('Static Weigh Bridge (SWB)', 'Cleanliness', 'Vendor'),
    ('Static Weigh Bridge (SWB)', 'Functional Condition', 'Vendor'),
    ('User Fare Display (UFD)', 'Cleanliness', 'Inhouse'),
    ('User Fare Display (UFD)', 'Physical Condition', 'Vendor'),
    ('Overhead Lane Status Light (OHLS)', 'Physical Condition', 'Vendor'),
    ('User Fare Display (UFD)', 'Functional Condition', 'Vendor'),
    ('Overhead Lane Status Light (OHLS)', 'Functional Condition', 'Vendor'),
    ('Traffic Lights', 'Physical Condition', 'Vendor'),
    ('Traffic Lights', 'Functional Condition', 'Vendor'),
    ('Traffic Lights', 'Cleanliness', 'Inhouse'),
    ('Automatic Boom Barrier', 'Cleanliness', 'Inhouse'),
    ('Automatic Boom Barrier', 'Functional Condition', 'Vendor'),
    ('Automatic Boom Barrier', 'Physical Condition', 'Vendor'),
    ('Overhead Lane Status Light (OHLS)', 'Cleanliness', 'Inhouse'),
    ('Automatic Vehicle Classification and Counting system (AVCC)', 'Cleanliness', 'Inhouse'),
    ('Weigh in Motion (WIM)', 'Physical Condition', 'Vendor'),
    ('Operator Monitor', 'Cleanliness', 'Inhouse'),
    ('Operator Monitor', 'Physical Condition', 'Vendor'),
    ('Operator Monitor', 'Functional Condition', 'Vendor'),
    ('Automatic Vehicle Classification and Counting system (AVCC)', 'Functional Condition', 'Vendor'),
    ('Automatic Vehicle Classification and Counting system (AVCC)', 'Physical Condition', 'Vendor'),
    ('Weigh in Motion (WIM)', 'Functional Condition', 'Vendor'),
    ('Weigh in Motion (WIM)', 'Cleanliness', 'Inhouse'),
    ('Quadrant Pitching', 'Presence of bushes and shrubs under/infront of the structure in the stream', 'Vendor'),
    ('Bus Bay', 'Cleanliness of Bus Bay', 'Inhouse'),
    ('Bus Bay', 'Physical Condition of Seating Arrangement', 'Vendor'),
    ('Bus Bay', 'Painting of Bus Shelter', 'Vendor'),
    ('Bus Bay', 'Physical Condition of Platform', 'Vendor'),
    ('Bus Bay', 'Physical Condition of Bus Shelter', 'Vendor'),
    ('Bus Bay', 'Condition of Bus Bay Area', 'Vendor'),
    ('Signages', 'Retro Reflectivity', 'Inhouse'),
    ('Signages', 'Cleanliness', 'Inhouse'),
    ('Signages', 'Physical Condition', 'Inhouse'),
    ('License Plate Indicatory Camera (LPIC)', 'Cleanliness', 'Vendor'),
    ('License Plate Indicatory Camera (LPIC)', 'Functional Condition', 'Vendor'),
    ('License Plate Indicatory Camera (LPIC)', 'Physical Condition', 'Vendor'),
    ('PTZ', 'Cleanliness', 'Inhouse'),
    ('PTZ', 'Physical Condition', 'Vendor'),
    ('Rigid Crash Barriers', 'Hand rail pipe - missing/rusted/damaged', 'Vendor'),
    ('Bus Bay', 'Physical Condition of Pavement Marking', 'Vendor'),
    ('Bus Bay', 'Functional Condition of Pavement Marking', 'Vendor'),
    ('Structure Numbering', 'Structure Numbering', 'Vendor'),
    ('Object Hazard Marker', 'Object hazard marking infront of crash barrier', 'Inhouse'),
    ('Pavement Markings', 'Shyness line Marking', 'Vendor'),
    ('Median', 'Cleanliness of median grass', 'Inhouse'),
    ('Median', 'Physical condition of median plants', 'Vendor'),
    ('Pavement Markings', 'Lane line Marking', 'Vendor'),
    ('Pavement Markings', 'Edge line Marking', 'Vendor'),
    ('Pavement Markings', 'Lane line Marking Night Visibility', 'Vendor'),
    ('Delineators', 'Functional Condition (Reflectivity)', 'Inhouse'),
    ('Delineators', 'Physical condition and functioning', 'Vendor'),
    ('Delineators', 'Cleanliness', 'Inhouse'),
    ('Pavement Markings', 'Edge line Marking Night Visibility', 'Vendor'),
    ('Pavement Markings', 'Shyness line Marking Night Visibility', 'Vendor'),
    ('Kilometer Stones', 'Physical condition and functioning', 'Inhouse'),
    ('Kilometer Stones', 'Cleanliness', 'Inhouse'),
    ('Traffic Blinkers and Signals', 'Functional Condition', 'Vendor'),
    ('Traffic Blinkers and Signals', 'Physical Condition', 'Vendor'),
    ('Traffic Blinkers and Signals', 'Cleanliness', 'Inhouse'),
    ('Truck Lay', 'Physical condition of truck lay bye area', 'Vendor'),
    ('Truck Lay', 'Cleanliness of truck lay bye', 'Inhouse'),
    ('Truck Lay', 'Painting of truck lay bye buliding', 'Vendor'),
    ('Truck Lay', 'Physical condition of truck lay bye buliding', 'Vendor'),
    ('Embankment', 'Raincuts', 'Vendor'),
    ('Embankment', 'Embankment Protection', 'Vendor'),
    ('Embankment', 'Slope Condition', 'Vendor'),
    ('Drainage', 'Functional condition of drain', 'Vendor'),
    ('Drainage', 'Cleanliness of drain', 'Inhouse'),
    ('Drainage', 'Physical condition of drain', 'Vendor'),
    ('Drainage Spouts', 'Missing grating or not', 'Vendor'),
    ('Drainage Spouts', 'Filled with dust and debris or chocked', 'Vendor'),
    ('Variable Message Sign', 'Physical Condition', 'Vendor'),
    ('Variable Message Sign', 'Functional Condition', 'Vendor'),
    ('Variable Message Sign', 'Cleanliness', 'Inhouse'),
    ('Truck Lay', 'Functional condition of pavement marking', 'Vendor'),
    ('Truck Lay', 'Physical condition of pavement marking', 'Vendor'),
    ('Toilet Block', 'Physical Condition of Wash basin', 'Vendor'),
    ('Toilet Block', 'Physical Condition of Doors', 'Vendor'),
    ('Toilet Block', 'Cleanliness of PC Toilet', 'Inhouse'),
    ('Toilet Block', 'Cleanliness of Mirror', 'Inhouse'),
    ('Toilet Block', 'Physical Condition of Windows', 'Vendor'),
    ('Toilet Block', 'Cleanliness of IWC', 'Inhouse'),
    ('Toilet Block', 'Cleanliness of Wash basin', 'Inhouse'),
    ('Toilet Block', 'Physical Condition of building', 'Vendor'),
    ('Toilet Block', 'Physical Condition of Urinals', 'Vendor'),
    ('Toilet Block', 'Physical Condition of Mirror', 'Vendor'),
    ('Toilet Block', 'Cleanliness of Urinals', 'Inhouse'),
    ('Toilet Block', 'Physical Condition of IWC', 'Vendor'),
    ('Toilet Block', 'Physical Condition of EWC', 'Vendor'),
    ('Toilet Block', 'Physical Condition of PC Toilet', 'Vendor'),
    ('Toilet Block', 'Physical Condition of Taps', 'Vendor'),
    ('Toilet Block', 'Cleanliness of EWC', 'Inhouse'),
    ('Non Buried Expansion Joint', 'Level difference of expansion joint from BC', 'Vendor'),
    ('Non Buried Expansion Joint', 'Puncturing/missing of expansion joint sealant', 'Vendor'),
    ('Non Buried Expansion Joint', 'Strip seal expansion joint buried under BC', 'Vendor'),
    ('Non Buried Expansion Joint', 'Condition of concrete on either side of edge angles', 'Vendor'),
    ('Non Buried Expansion Joint', 'Damages to edge angles', 'Vendor'),
    ('Non Buried Expansion Joint', 'Accumulation of debris and dust', 'Inhouse'),
    ('AntiGlazer', 'Physical condition and functioning', 'Vendor'),
    ('AntiGlazer', 'Functional Condition', 'Vendor'),
    ('AntiGlazer', 'Cleanliness', 'Inhouse'),
    ('Anti Glazer', 'Cleanliness', 'Inhouse'),
    ('Anti Glazer', 'Functional Condition', 'Vendor'),
    ('Anti Glazer', 'Physical condition and functioning', 'Vendor'),
    ('VASD', 'Cleanliness', 'Inhouse'),
    ('VASD', 'Physical Condition', 'Vendor'),
    ('VASD', 'Functional Condition', 'Vendor'),
    ('MET', 'Functional Condition', 'Vendor'),
    ('MET', 'Cleanliness', 'Inhouse'),
    ('MET', 'Physical Condition', 'Vendor'),
]


# ──────────────────────────────────────────────────────────────────────────
# Project Name -> Invit/Non-Invit. NOTE: BFHL is always "Non-Invit" (fixed
# exception -- do not change to "Invit" even if that seems inconsistent
# with other entries). No sheet in the workbook holds this
# mapping (unlike Vendor/Inhouse, which comes from Sheet5), so it's kept
# here as a plain table. Add new project codes here as they show up.
#
# NOTE: DHMEPL was changed from "Invit" to "Non-Invit" per explicit request.
# ──────────────────────────────────────────────────────────────────────────
PROJECT_INVIT_MAP = {
    "ADTPL": "Non-Invit", "APEL": "Invit",     "BFHL": "Non-Invit", "BWHPL": "Invit",
    "DATL": "Invit",      "DHMEPL": "Non-Invit", "FRHL": "Invit",    "GAEPL": "Invit",
    "JMTPL": "Invit",     "JUHPL": "Invit",    "KETPL": "Invit",    "KHEPL": "Non-Invit",
    "KMTPL": "Invit",     "KTIPL": "Invit",    "MBEL": "Invit",     "MHPL": "Invit",
    "MKTPL": "Invit",     "MSHP": "Invit",     "NAM": "Invit",      "NDEPL": "Invit",
    "NKTPL": "Invit",     "SIPL": "Invit",     "SMTPL": "Invit",    "SPPL": "Invit",
    "WMPTL": "Non-Invit", "WUPTL": "Invit",    "WVEL": "Invit",
}


# ──────────────────────────────────────────────────────────────────────────
# Project Name -> Regional Head. Fixed, hardcoded list (no sheet in the
# workbook holds this mapping), same style as PROJECT_INVIT_MAP above. Edit
# PROJECT_REGIONAL_HEAD_MAP below when a new project code shows up that
# isn't in the list yet; the script will warn loudly (rather than silently
# guessing) if it hits a project code with no entry.
# ──────────────────────────────────────────────────────────────────────────
PROJECT_REGIONAL_HEAD_MAP = {
    "ADTPL": "Mr KK Rao",     "APEL": "Ms Vasundhara", "BFHL": "Mr Sanjay",     "BWHPL": "Ms Vasundhara",
    "DATL": "Mr Shreedhar",   "DHMEPL": "Mr Shreedhar","FRHL": "Mr Sanjay",     "GAEPL": "Mr Shreedhar",
    "JMTPL": "Mr Shreedhar",  "JUHPL": "Mr Sanjay",    "KETPL": "Mr KK Rao",    "KHEPL": "Mr Shreedhar",
    "KMTPL": "Mr Sanjay",     "KTIPL": "Ms Vasundhara","MBEL": "Mr Shreedhar",  "MHPL": "Ms Vasundhara",
    "MKTPL": "Mr KK Rao",     "MSHP": "Ms Vasundhara", "NAM": "Mr KK Rao",      "NDEPL": "Mr KK Rao",
    "NKTPL": "Mr KK Rao",     "SIPL": "Ms Vasundhara", "SMTPL": "Mr KK Rao",    "SPPL": "Ms Vasundhara",
    "WMPTL": "Mr Shreedhar",  "WUPTL": "Mr Sanjay",    "WVEL": "Mr KK Rao",
}


# ──────────────────────────────────────────────────────────────────────────
# Raw-file column names -> Sheet1 column names (unchanged from v1/v2)
# ──────────────────────────────────────────────────────────────────────────
DIRECT_MAP = {
    "project name": "Project Name",
    "roadaid id": "Roadaid ID",
    "category": "Category",
    "asset type": "Asset Type",
    "direction": "Direction",
    "road type": "Road Type",
    "chainage": "Chainage",
    "parameter": "Parameter",
    "concern raised": "Concern Raised",
    "spv rating": "SPV Rating",
    "ho rating": "HO Rating",
    "ho remarks": "HO Remarks",
    "structure name": "Structure Name",
    "method": "Method",
    "signboard type": "Signboard type",
    "signboard classification": "Signboard Classification",
    "placement": "Placement",
}
DATE_CREATED_KEYS = {"date created", "datecreated", "created date"}

SHEET1_COLS = [
    "Project Name", "Roadaid ID", "Category", "Asset Type", "Direction",
    "Road Type", "Chainage", "Parameter", "Concern Raised", "SPV Rating",
    "HO Rating", "HO Remarks", "Structure Name", "Method", "Signboard type",
    "Signboard Classification", "Placement",
]  # A..Q, in order (17 columns)

METHOD_CODE = {"digital": "DL", "conventional": "CL"}


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def excel_col_letters(n: int):
    out = []
    for i in range(1, n + 1):
        s = ""
        x = i
        while x > 0:
            x, rem = divmod(x - 1, 26)
            s = chr(65 + rem) + s
        out.append(s)
    return out


COL_LETTERS = excel_col_letters(50)  # A..AX
COL_LETTER = {i: COL_LETTERS[i] for i in range(50)}  # 0-based idx -> letter


# ──────────────────────────────────────────────────────────────────────────
# Step 1: load lookups from the master workbook -- Sheet3 code tables,
# PC/FC/CC dominant-class map learned from existing Sheet1 data, and now
# also Sheet5 (Asset Type + Parameter -> Vendor/Inhouse).
# ──────────────────────────────────────────────────────────────────────────

def load_master_lookups(master_path: str):
    print("Loading Sheet3 code tables (Category/Asset Type/Parameter)...")
    s3 = pd.read_excel(master_path, sheet_name="Sheet3", header=0)
    cols = list(s3.columns)

    def pair_map(name_col, code_col):
        d = {}
        for name, code in zip(s3[name_col], s3[code_col]):
            if pd.notna(name) and pd.notna(code):
                d[_norm(name)] = str(code).strip()
        return d

    category_map = pair_map(cols[0], cols[1])
    asset_map = pair_map(cols[3], cols[4])
    parameter_map = pair_map(cols[6], cols[7])

    print("Learning PC/FC/CC dominant class per Parameter from existing rows...")
    t = time.time()
    df_hy = pd.read_excel(master_path, sheet_name="Sheet1", usecols=[7, 24], header=0)
    df_hy.columns = ["Parameter", "PCFCCC"]
    mode_map = (
        df_hy.dropna()
        .groupby(df_hy["Parameter"].map(_norm))["PCFCCC"]
        .agg(lambda s: s.value_counts().idxmax())
        .to_dict()
    )
    overall_default = df_hy["PCFCCC"].value_counts().idxmax()
    print(f"  learned {len(mode_map)} parameter -> class mappings in {time.time()-t:.1f}s")

    print("Loading Sheet5 Vendor/Inhouse table (Asset Type + Parameter -> Vendor/Inhouse)...")
    vendor_map = {}
    vendor_rows = []  # display-case (Asset Type, Parameter, Vendor/Inhouse) tuples
    sheet5_is_valid = False
    try:
        s5 = pd.read_excel(master_path, sheet_name="Sheet5", header=0)
        s5cols = list(s5.columns)
        # Validate this actually looks like the Asset Type/Parameter/Verdict
        # table and not some unrelated data that happens to be sitting on a
        # tab also named "Sheet5" (this has happened before -- a master's
        # Sheet5 tab held leftover raw project data instead of the lookup
        # table, which silently produced a garbage vendor_map with no
        # matches). We require: exactly the expected header names, a
        # reasonable row count, and a verdict column containing only
        # Vendor/Inhouse-like values.
        header_ok = (
            len(s5cols) >= 3
            and _norm(s5cols[0]) == "asset type"
            and _norm(s5cols[1]) == "parameter"
        )
        if header_ok:
            candidate = {}
            candidate_rows = []
            for asset, param, verdict in zip(s5[s5cols[0]], s5[s5cols[1]], s5[s5cols[2]]):
                if pd.notna(asset) and pd.notna(param) and pd.notna(verdict):
                    verdict = str(verdict).strip()
                    candidate[(_norm(asset), _norm(param))] = verdict
                    candidate_rows.append((str(asset).strip(), str(param).strip(), verdict))
            verdict_values = {v.lower() for v in candidate.values()}
            if candidate and verdict_values <= {"vendor", "inhouse"}:
                vendor_map = candidate
                vendor_rows = candidate_rows
                sheet5_is_valid = True
    except Exception as e:
        print(f"  NOTE: could not read master's Sheet5 tab ({e}); will use the "
              f"built-in universal reference table instead.")

    if sheet5_is_valid:
        print(f"  loaded {len(vendor_map)} Asset Type/Parameter -> Vendor/Inhouse "
              f"mapping(s) from the master's own Sheet5 tab (looks valid).")
    else:
        print(f"  *** Master's Sheet5 tab does not look like a genuine Asset Type/"
              f"Parameter/Vendor-Inhouse table (wrong headers, wrong sheet content, "
              f"or missing) -- falling back to the built-in UNIVERSAL reference "
              f"table (seeded from Consolidated_June26.xlsx, {len(VENDOR_MAP_FALLBACK_ROWS)} "
              f"entries, verified against 481,447 real rows with 0 mismatches). ***")
        vendor_map = {}
        vendor_rows = list(VENDOR_MAP_FALLBACK_ROWS)
        for asset, param, verdict in VENDOR_MAP_FALLBACK_ROWS:
            vendor_map[(_norm(asset), _norm(param))] = verdict
        print(f"  loaded {len(vendor_map)} Asset Type/Parameter -> Vendor/Inhouse "
              f"mapping(s) from the built-in universal reference table")

    return category_map, asset_map, parameter_map, mode_map, overall_default, vendor_map, vendor_rows


# ──────────────────────────────────────────────────────────────────────────
# Step 2: load & harmonize the raw monthly files. Unchanged from v1/v2.
# ──────────────────────────────────────────────────────────────────────────

def _read_excel_data_sheet(path: str) -> tuple:
    """
    Picks which sheet of a raw .xlsx holds the actual row-level data, and
    flags whether that file's SPV Rating was ever calculated at all.

    Two shapes show up in practice:
      1. Plain ratings-list file: a single sheet (or the row-level sheet
         is the only one that looks like real data) -- SPV Rating is
         genuinely filled in per-row here, so the normal SPV/HO fallback
         applies.
      2. "...Ratings_List_REPORT.xlsx" shape: a 'Sheet1' tab holding a
         pivoted/aggregated summary (Category, Total Audited, No of
         Issues, ...) sitting ALONGSIDE a second sheet with the real
         row-level data (Project Name, Roadaid ID, ... SPV Rating, HO
         Rating, ...). In this shape SPV Rating is never calculated --
         every value is "-" -- so SPV Final Rating must NOT fall back to
         HO Rating for these rows.

    Detection: if the workbook has a tab literally named "Sheet1" AND at
    least one other tab, treat it as shape 2 -- read the other (non-
    "Sheet1") tab as the data, and flag spv_calculated=False. Otherwise
    read the first/only sheet normally and flag spv_calculated=True.
    """
    xl = pd.ExcelFile(path)
    sheet_names = xl.sheet_names
    if len(sheet_names) > 1 and any(_norm(s) == "sheet1" for s in sheet_names):
        data_sheet = next(s for s in sheet_names if _norm(s) != "sheet1")
        print(f"    (detected 'Sheet1' summary tab alongside '{data_sheet}' -- "
              f"reading '{data_sheet}' as the data sheet; SPV Rating in this "
              f"file is treated as not-calculated)")
        return pd.read_excel(xl, sheet_name=data_sheet, dtype=str), False
    return pd.read_excel(xl, sheet_name=sheet_names[0], dtype=str), True


def load_raw_file(path: str) -> pd.DataFrame:
    if path.lower().endswith(".csv"):
        df = pd.read_csv(path, dtype=str, keep_default_na=True)
        spv_calculated = True
    else:
        df, spv_calculated = _read_excel_data_sheet(path)

    rename = {}
    date_created_col = None
    for c in df.columns:
        key = _norm(c)
        if key in DIRECT_MAP:
            rename[c] = DIRECT_MAP[key]
        elif key in DATE_CREATED_KEYS:
            date_created_col = c
    df = df.rename(columns=rename)

    for col in SHEET1_COLS:
        if col not in df.columns:
            df[col] = None

    if date_created_col is None:
        raise ValueError(f"{path}: could not find a 'Date Created' column")
    df["_DateCreated"] = df[date_created_col]
    df["_SourceFile"] = os.path.basename(path)
    df["_SPVCalculated"] = spv_calculated

    return df[SHEET1_COLS + ["_DateCreated", "_SourceFile", "_SPVCalculated"]]


_TEMP_DIRS_TO_CLEAN = []


def _cleanup_temp_dirs():
    for d in _TEMP_DIRS_TO_CLEAN:
        shutil.rmtree(d, ignore_errors=True)


atexit.register(_cleanup_temp_dirs)


def resolve_month_input(month_input: str) -> str:
    """
    Accepts EITHER a folder of raw .csv/.xlsx files OR a .zip archive
    containing them (e.g. the "June_final.zip" you get emailed with the
    27 project files inside, possibly nested one level down in a
    sub-folder such as "June final/"). If given a .zip, extracts it to a
    fresh temp directory and returns that path so the rest of the
    pipeline can treat it exactly like a plain folder. This is what was
    missing before -- pointing the script straight at a .zip silently
    matched zero files (glob only looks for *.csv/*.xlsx, never *.zip),
    so the run fell into "recompute-only mode" and no new rows ever got
    appended, no matter how many project files were inside the zip.
    """
    if os.path.isdir(month_input):
        return month_input

    if zipfile.is_zipfile(month_input):
        tmp_dir = tempfile.mkdtemp(prefix="merge_month_zip_")
        _TEMP_DIRS_TO_CLEAN.append(tmp_dir)
        print(f"  '{month_input}' is a zip archive -- extracting to {tmp_dir} ...")
        with zipfile.ZipFile(month_input, "r") as zf:
            zf.extractall(tmp_dir)
        return tmp_dir

    raise ValueError(
        f"'{month_input}' is neither a folder nor a .zip file -- point this "
        f"script at the raw month's folder of .csv/.xlsx files, or a .zip "
        f"of them."
    )


def load_all_raw_files(folder: str) -> pd.DataFrame:
    # Recursive (**) so files nested inside a sub-folder of the zip/folder
    # (e.g. "June final/ALL_XXX_Ratings_List.csv") are still found -- a
    # plain top-level glob missed these entirely.
    paths = sorted(glob.glob(os.path.join(folder, "**", "*.csv"), recursive=True)) + \
            sorted(glob.glob(os.path.join(folder, "**", "*.xlsx"), recursive=True))
    # Drop Excel's transient lock files (~$Foo.xlsx) and macOS zip junk,
    # and de-dupe in case a path matched both patterns.
    paths = sorted({
        p for p in paths
        if not os.path.basename(p).startswith("~$")
        and "__MACOSX" not in p.split(os.sep)
    })
    if not paths:
        print(f"  No .csv/.xlsx files found in {folder} -- recompute-only mode "
              f"(no new rows will be added, existing rows will just be re-fixed).")
        return pd.DataFrame(columns=SHEET1_COLS + ["_DateCreated", "_SourceFile", "_SPVCalculated"])

    frames = []
    for p in paths:
        print(f"  Loading {os.path.relpath(p, folder)} ...")
        frames.append(load_raw_file(p))
    df = pd.concat(frames, ignore_index=True)
    print(f"  Loaded {len(df):,} raw rows from {len(paths)} files")
    return df


# ──────────────────────────────────────────────────────────────────────────
# Step 3: compute derived columns R:Z for the new rows. Unchanged from v2.
# ──────────────────────────────────────────────────────────────────────────

def compute_new_rows(df: pd.DataFrame, category_map, asset_map, parameter_map,
                      mode_map, overall_default) -> pd.DataFrame:
    if df.empty:
        for col in ("SPV Final Rating", "HO Final Rating", "Month", "Asset ST",
                    "Division CT", "Parameter CT", "method CT", "PC/FC/CC", "Z"):
            df[col] = pd.Series(dtype=object)
        return df

    print("Computing derived columns (R:Z) for new rows...")

    spv = pd.to_numeric(df["SPV Rating"], errors="coerce")
    ho = pd.to_numeric(df["HO Rating"], errors="coerce")
    chainage = pd.to_numeric(df["Chainage"], errors="coerce")

    for label, raw, coerced in (
        ("SPV Rating", df["SPV Rating"], spv),
        ("HO Rating", df["HO Rating"], ho),
        ("Chainage", df["Chainage"], chainage),
    ):
        bad_mask = raw.notna() & coerced.isna() & (raw.astype(str).str.strip() != "")
        if bad_mask.any():
            bad_vals = sorted(set(raw[bad_mask].astype(str).str.strip()))
            print(f"  NOTE: {bad_mask.sum()} non-numeric {label} value(s) kept as text: "
                  f"{bad_vals[:10]}{' ...' if len(bad_vals) > 10 else ''}")

    # SPV Final Rating: normally falls back to HO Rating whenever SPV
    # Rating is missing. BUT some raw files (the "Sheet1 summary tab +
    # ratings-list sheet" shape -- see _read_excel_data_sheet) never
    # calculate SPV Rating at all; every value is "-". For rows from
    # those files, falling back to HO would show a number that was never
    # actually an SPV rating, so SPV Final Rating is left as "-" instead.
    # Rows from normal (single-sheet / CSV) files keep the old fallback.
    # HO Final Rating is unaffected either way -- it still falls back to
    # SPV when HO itself is missing.
    spv_final = spv.where(spv.notna(), ho).astype(object)
    if "_SPVCalculated" in df.columns:
        not_calc_mask = (~df["_SPVCalculated"].fillna(True).astype(bool)) & spv.isna()
        spv_final[not_calc_mask.to_numpy()] = "-"
    df["SPV Final Rating"] = spv_final
    df["HO Final Rating"] = ho.where(ho.notna(), spv)

    # IMPORTANT: format="mixed" is required here. When 27 files are
    # concatenated into one _DateCreated column, pandas' default
    # to_datetime() infers a SINGLE date format from the whole column and
    # applies it to every row. If even one project's export uses a
    # different date format than the rest (e.g. JMTPL's raw file uses
    # "08-Jun-26 6:00:23 AM" while most others use "08-06-2026 06:00"),
    # every row from the minority-format file silently becomes NaT --
    # even though that same file parses perfectly fine on its own. This
    # is exactly what caused JMTPL's June rows to all get a blank Month
    # despite the raw file itself having valid dates. format="mixed"
    # parses each value independently by whatever format matches it,
    # instead of forcing one format across the whole batch.
    dt = pd.to_datetime(df["_DateCreated"], dayfirst=True, errors="coerce", format="mixed")
    df["Month"] = dt.values.astype("datetime64[M]")

    def _file_project_breakdown(mask: pd.Series) -> str:
        """Group affected rows by source file + project so the warning
        points straight at the culprit instead of just a bare count --
        this is what you'd otherwise have to reverse-engineer from the
        merged master afterward."""
        cols = [c for c in ("_SourceFile", "Project Name") if c in df.columns]
        if not cols:
            return ""
        sub = df.loc[mask, cols]
        counts = sub.value_counts().sort_values(ascending=False)
        lines = [f"        {' / '.join(str(x) for x in idx)}: {n:,} row(s)"
                 for idx, n in counts.head(10).items()]
        return "\n" + "\n".join(lines) if lines else ""

    bad_date_mask = df["_DateCreated"].notna() & dt.isna()
    n_bad_date = int(bad_date_mask.sum())
    if n_bad_date:
        print(f"  *** WARNING: {n_bad_date:,} row(s) have a 'Date Created' value that "
              f"could not be parsed -> Month will be BLANK for these rows. ***")
        print(f"      This is a raw-data issue, not a formula bug -- these rows will be "
              f"grouped correctly (blank groups with blank) so they will NOT cause any "
              f"#DIV/0!, but their Month-based figures will be incomplete. Sample bad "
              f"values: {sorted(set(df.loc[bad_date_mask, '_DateCreated'].astype(str)))[:10]}")
        print(f"      By file/project:{_file_project_breakdown(bad_date_mask)}")
    n_missing_date = int(df["_DateCreated"].isna().sum())
    if n_missing_date:
        missing_mask = df["_DateCreated"].isna()
        print(f"  *** WARNING: {n_missing_date:,} row(s) have NO 'Date Created' value at all. ***")
        print(f"      By file/project:{_file_project_breakdown(missing_mask)}")

    def lookup(series, mapping, default=""):
        missing = set()
        out = []
        for v in series:
            key = _norm(v) if pd.notna(v) else ""
            code = mapping.get(key)
            if code is None:
                missing.add(v)
                code = default
            out.append(code)
        return out, missing

    df["Asset ST"], miss_a = lookup(df["Asset Type"], asset_map)
    df["Division CT"], miss_c = lookup(df["Category"], category_map)
    df["Parameter CT"], miss_p = lookup(df["Parameter"], parameter_map)

    for label, miss in (("Asset Type", miss_a), ("Category", miss_c), ("Parameter", miss_p)):
        if miss:
            print(f"  WARNING: {len(miss)} unmapped {label} value(s) (code left blank): "
                  f"{sorted(x for x in miss if x)[:10]}{' ...' if len(miss) > 10 else ''}")

    df["method CT"] = df["Method"].map(lambda v: METHOD_CODE.get(_norm(v), ""))

    pcfc, miss_pcfc = lookup(df["Parameter"], mode_map, default=overall_default)
    df["PC/FC/CC"] = pcfc
    if miss_pcfc:
        print(f"  NOTE: {len(miss_pcfc)} Parameter value(s) not seen before in master; "
              f"defaulted to '{overall_default}': {sorted(x for x in miss_pcfc if x)[:10]}")

    key = list(zip(df["Project Name"], df["Month"]))
    tmp = pd.DataFrame({"key": key, "ho": df["HO Final Rating"]})
    is1 = (tmp["ho"] == 1).astype(int)
    is5 = (tmp["ho"] == 5).astype(int)
    c1 = is1.groupby(tmp["key"]).transform("sum")
    c5 = is5.groupby(tmp["key"]).transform("sum")
    total = tmp.groupby("key")["ho"].transform("count")
    with np.errstate(divide="ignore", invalid="ignore"):
        pw = (c1 * 5 + c5) / total
        df["Z"] = np.round(10 * (1 - pw), 4)

    print(f"  {len(df):,} new rows ready")
    return df


def check_month_consistency(df: pd.DataFrame, source_label: str) -> None:
    """
    Guard rail for the zip/folder merge: every row's Month is derived
    independently from its own 'Date Created' value, so if the zip/folder
    actually contains files from more than one month (e.g. a stray
    leftover file from last month, or someone accidentally zipped two
    months together), that would silently merge mismatched months into
    the master as if they were one batch.

    This checks the whole batch AFTER Month has been computed and refuses
    to proceed if more than one distinct month is present, printing a
    per-file breakdown so you can see exactly which file(s) are the
    problem. Rows with an unparseable/missing date (Month is blank) are
    ignored here -- those are already flagged separately.
    """
    if df.empty or "Month" not in df.columns:
        return
    valid = df.dropna(subset=["Month"])
    if valid.empty:
        return

    months = sorted(valid["Month"].unique())
    if len(months) <= 1:
        return

    def fmt(m):
        return pd.Timestamp(m).strftime("%b %Y")

    print(f"\n*** ERROR: raw files in '{source_label}' span MORE THAN ONE MONTH -- "
          f"refusing to merge. Months found: {[fmt(m) for m in months]}")
    if "_SourceFile" in valid.columns:
        breakdown = valid.groupby("_SourceFile")["Month"].apply(
            lambda s: sorted({fmt(m) for m in s.unique()}))
        for fname, months_in_file in breakdown.items():
            flag = "  <-- mixed dates WITHIN this file" if len(months_in_file) > 1 else ""
            print(f"    {fname}: {', '.join(months_in_file)}{flag}")
    print("    Fix the zip/folder so it contains only one month's files "
          "(remove the stray file(s) above, or split them into separate "
          "runs), then re-run.\n")
    sys.exit(1)


def check_duplicate_project_files(df: pd.DataFrame, source_label: str) -> None:
    """
    Warns if the same Project Name shows up across more than one raw file
    in this batch. This is a common cause of exactly the kind of bug seen
    with JMTPL's June data: an old/duplicate copy of a project's file sits
    in the zip alongside a freshly re-sent one, both get loaded, and the
    stale copy's differently-formatted (or blank) Date Created values
    silently blank out Month for a chunk of that project's rows -- while
    a fresh copy of the same file, checked in isolation later, looks
    perfectly fine and doesn't explain the discrepancy on its own.
    Does not block the run (unlike check_month_consistency) since this
    can legitimately happen (e.g. one file per Category for a project);
    it just flags it loudly so it can be eyeballed before trusting the
    merge.
    """
    if df.empty or "_SourceFile" not in df.columns or "Project Name" not in df.columns:
        return
    proj_files = df.groupby("Project Name")["_SourceFile"].unique()
    dupes = {p: list(files) for p, files in proj_files.items() if len(files) > 1}
    if not dupes:
        return
    print(f"\n  *** NOTE: the following project(s) appear in MORE THAN ONE raw file "
          f"within '{source_label}'. If this isn't intentional (e.g. a stray old "
          f"file left in the zip alongside a re-sent one), it can double-count "
          f"rows or blend a good file's dates with a bad file's: ***")
    for p, files in dupes.items():
        print(f"      {p}: {files}")


NUMERIC_FIELDS = {"Chainage", "SPV Rating", "HO Rating", "SPV Final Rating",
                  "HO Final Rating", "Z"}
DATE_FIELDS = {"Month"}
AZ_FIELDS = [
    "Project Name", "Roadaid ID", "Category", "Asset Type", "Direction",
    "Road Type", "Chainage", "Parameter", "Concern Raised", "SPV Rating",
    "HO Rating", "HO Remarks", "Structure Name", "Method", "Signboard type",
    "Signboard Classification", "Placement", "SPV Final Rating",
    "HO Final Rating", "Month", "Asset ST", "Division CT", "Parameter CT",
    "method CT", "PC/FC/CC", "Z",
]
EXCEL_EPOCH = pd.Timestamp("1899-12-30")
DATE_STYLE = "15"


def fmt_num(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, str) and v.strip() == "":
        return None
    try:
        fv = float(v)
    except (TypeError, ValueError):
        return None
    if fv == int(fv):
        return str(int(fv))
    return repr(fv)


class SharedStringInterner:
    """De-duplicates text against the workbook's existing sharedStrings.xml
    table and hands out compact <c t="s"><v>IDX</v></c> cells instead of
    repeating the full text inline every time (t="inlineStr"). This is the
    fix for the file-size blowup: with ~190,000 new rows x ~15 text columns,
    inlineStr repeats full strings (many highly repetitive -- project codes,
    categories, parameters, "-", etc.) every single time, which previously
    inflated a single month's worth of new rows to ~1GB of XML and caused
    the process to be OOM-killed while assembling the final .xlsx. Shared
    strings store each distinct string once and reference it by a small
    integer, which is what the rest of the workbook already does for its
    existing rows.
    """
    def __init__(self, existing_strings: list):
        self.strings = list(existing_strings)
        self.index = {s: i for i, s in enumerate(self.strings)}
        self.new_ref_count = 0  # how many *references* (not unique strings) we hand out
        self._orig_len = len(self.strings)

    def intern(self, text: str) -> int:
        self.new_ref_count += 1
        idx = self.index.get(text)
        if idx is not None:
            return idx
        idx = len(self.strings)
        self.strings.append(text)
        self.index[text] = idx
        return idx

    @property
    def newly_added(self) -> list:
        """Strings appended beyond what load_shared_strings() originally loaded."""
        return self.strings[self._orig_len:] if hasattr(self, "_orig_len") else []


def build_new_row_az_xml(row_num: int, vals: dict, ss: "SharedStringInterner") -> str:
    """Build the A:Z portion (raw + directly-derived columns) for one NEW row.
    AA:AX are appended separately once the full-sheet recompute is done."""
    parts = [f'<row r="{row_num}" spans="1:50">']
    for i, field in enumerate(AZ_FIELDS):
        col = COL_LETTER[i]
        v = vals.get(field)
        if field in DATE_FIELDS:
            if v is None or pd.isna(v):
                continue
            serial = (pd.Timestamp(v) - EXCEL_EPOCH).days
            parts.append(f'<c r="{col}{row_num}" s="{DATE_STYLE}"><v>{serial}</v></c>')
        elif field in NUMERIC_FIELDS:
            n = fmt_num(v)
            if n is not None:
                parts.append(f'<c r="{col}{row_num}"><v>{n}</v></c>')
            elif not (v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() == ""):
                idx = ss.intern(str(v).strip())
                parts.append(f'<c r="{col}{row_num}" t="s"><v>{idx}</v></c>')
        else:
            if v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() == "":
                continue
            idx = ss.intern(str(v).strip())
            parts.append(f'<c r="{col}{row_num}" t="s"><v>{idx}</v></c>')
    # NOTE: no </row> here -- AA:AX cells + </row> get appended by the caller.
    return "".join(parts)


# ──────────────────────────────────────────────────────────────────────────
# Sheet resolution helpers (tab name -> physical part). Unchanged from v1/v2;
# this is what correctly resolves "Sheet2" to whichever sheetN.xml it
# actually lives in, regardless of tab reordering.
# ──────────────────────────────────────────────────────────────────────────

def find_real_sheet_file(workbook_xml: bytes, workbook_rels_xml: bytes, sheet_name: str) -> str:
    m = re.search(rb'<sheet name="' + re.escape(sheet_name.encode()) +
                  rb'"[^>]*r:id="(rId\d+)"', workbook_xml)
    if not m:
        return None
    rid = m.group(1).decode()
    m2 = re.search(rb'<Relationship Id="' + re.escape(rid.encode()) +
                   rb'"[^>]*Target="([^"]+)"', workbook_rels_xml)
    if not m2:
        return None
    return "xl/" + m2.group(1).decode()


def load_master_last_row(sheet1_xml: bytes) -> int:
    m = re.search(rb'<dimension ref="[A-Z]+1:[A-Z]+(\d+)"', sheet1_xml)
    if not m:
        raise RuntimeError("Could not find <dimension> in sheet1.xml")
    return int(m.group(1))


# ──────────────────────────────────────────────────────────────────────────
# Shared strings + generic cell text resolution (needed because rows in
# the file can be encoded either as shared-string references (older rows)
# or inline strings (rows appended by v1/v2/v3 of this script) -- both
# must resolve to the same text for grouping to be correct).
# ──────────────────────────────────────────────────────────────────────────

def load_shared_strings(zin) -> list:
    try:
        data = zin.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    si_blocks = re.findall(rb"<si>(.*?)</si>", data, re.S)
    out = []
    for b in si_blocks:
        texts = re.findall(rb"<t[^>]*>(.*?)</t>", b, re.S)
        out.append(html.unescape(b"".join(texts).decode("utf-8")))
    return out


_CELL_RE = re.compile(rb'<c r="([A-Z]+)\d+"([^>]*)>(.*?)</c>', re.S)
_INLINE_T_RE = re.compile(rb"<t[^>]*>(.*?)</t>", re.S)
_V_RE = re.compile(rb"<v>([^<]*)</v>")
_V_RE_DOTALL = re.compile(rb"<v>(.*?)</v>", re.S)


def resolve_text_cell(attrs: bytes, content: bytes, shared_strings: list):
    if b't="s"' in attrs:
        m = _V_RE.search(content)
        return shared_strings[int(m.group(1))] if m else None
    elif b't="inlineStr"' in attrs:
        m = _INLINE_T_RE.search(content)
        return html.unescape(m.group(1).decode("utf-8")) if m else None
    elif b't="str"' in attrs:
        m = _V_RE_DOTALL.search(content)
        return html.unescape(m.group(1).decode("utf-8")) if m else None
    m = _V_RE.search(content)
    return m.group(1).decode() if m else None


def resolve_numeric_cell(content: bytes):
    m = _V_RE.search(content)
    if not m:
        return np.nan
    try:
        return float(m.group(1))
    except ValueError:
        return np.nan


# ──────────────────────────────────────────────────────────────────────────
# Extract A,C,D,H,S,T for every existing row in the master's Sheet1, plus
# byte offsets so we can copy each row's A:Z cells through unchanged and
# splice in freshly computed AA:AX cells.
# ──────────────────────────────────────────────────────────────────────────

def scan_row_positions(sheet1_xml: bytes):
    row_re = re.compile(rb'<row r="(\d+)"[^>]*>')
    row_starts = list(row_re.finditer(sheet1_xml))
    aa_starts = [m.start() for m in re.finditer(rb'<c r="AA\d+"', sheet1_xml)]
    row_ends = [m.end() for m in re.finditer(rb'</row>', sheet1_xml)]
    if not (len(row_starts) == len(aa_starts) == len(row_ends)):
        raise RuntimeError(
            f"Row/AA-cell/row-end counts don't line up "
            f"({len(row_starts)}/{len(aa_starts)}/{len(row_ends)}) -- master "
            f"file structure is not what this script expects (every row must "
            f"have an AA cell)."
        )
    return {
        "row_nums": [int(m.group(1)) for m in row_starts],
        "row_tag_start": [m.start() for m in row_starts],
        "row_tag_end": [m.end() for m in row_starts],
        "aa_start": aa_starts,
        "row_end": row_ends,
    }


def extract_old_rows(sheet1_xml: bytes, pos: dict, shared_strings: list) -> pd.DataFrame:
    print("Extracting existing rows' Project/Category/Asset Type/Parameter/Rating/Month...")
    t0 = time.time()
    n = len(pos["row_nums"])
    row_tag_end = pos["row_tag_end"]
    aa_start = pos["aa_start"]
    WANT = {"A", "C", "D", "H", "S", "T"}

    row_out, a_out, c_out, d_out, h_out = [], [], [], [], []
    s_out, t_out = [], []

    for i in range(1, n):  # skip header row (index 0)
        seg = sheet1_xml[row_tag_end[i]:aa_start[i]]
        a = c = d = h = None
        s = t = np.nan
        for m in _CELL_RE.finditer(seg):
            col = m.group(1).decode()
            if col not in WANT:
                continue
            attrs, content = m.group(2), m.group(3)
            if col == "S":
                s = resolve_numeric_cell(content)
            elif col == "T":
                t = resolve_numeric_cell(content)
            else:
                txt = resolve_text_cell(attrs, content, shared_strings)
                if col == "A": a = txt
                elif col == "C": c = txt
                elif col == "D": d = txt
                elif col == "H": h = txt
        row_out.append(pos["row_nums"][i])
        a_out.append(a); c_out.append(c); d_out.append(d); h_out.append(h)
        s_out.append(s); t_out.append(t)
        if i % 100000 == 0:
            print(f"  {i:,} rows scanned ({time.time()-t0:.1f}s)")

    print(f"  extracted {n-1:,} existing rows in {time.time()-t0:.1f}s")
    return pd.DataFrame({
        "row": row_out, "A": a_out, "C": c_out, "D": d_out, "H": h_out,
        "S": s_out, "T": t_out,
    })


def build_project_lookup(sheet2_xml: bytes, shared_strings: list) -> dict:
    """Resolve the *real* Sheet2 tab's A:B columns (Project code -> Full
    name) into a plain dict, regardless of shared-string vs inline
    encoding."""
    mapping = {}
    for m in re.finditer(rb'<row r="\d+"[^>]*>(.*?)</row>', sheet2_xml, re.S):
        row_body = m.group(1)
        a_m = re.search(rb'<c r="A\d+"([^>]*)>(.*?)</c>', row_body, re.S)
        b_m = re.search(rb'<c r="B\d+"([^>]*)>(.*?)</c>', row_body, re.S)
        if not a_m or not b_m:
            continue
        a_txt = resolve_text_cell(a_m.group(1), a_m.group(2), shared_strings)
        b_txt = resolve_text_cell(b_m.group(1), b_m.group(2), shared_strings)
        if a_txt is not None and b_txt is not None:
            mapping[a_txt] = b_txt
    return mapping


def get_sheet_last_row(zin, sheet_path: str, fallback: int = 28) -> int:
    if not sheet_path:
        return fallback
    try:
        data = zin.read(sheet_path)
    except KeyError:
        return fallback
    m = re.search(rb'<dimension ref="[A-Z]+\d+:[A-Z]+(\d+)"/>', data)
    return int(m.group(1)) if m else fallback


# ──────────────────────────────────────────────────────────────────────────
# Full-sheet recompute of AA:AV (unchanged math from v2). Works over the
# COMBINED old+new rows so every row's group totals reflect the complete
# dataset.
# ──────────────────────────────────────────────────────────────────────────

def compute_all_derived(df: pd.DataFrame) -> pd.DataFrame:
    print("Recomputing AA:AV for every row (old + new) in Python "
          "(no Excel formulas, no #DIV/0! possible)...")
    t0 = time.time()
    df = df.copy()
    df["is1"] = (df["S"] == 1)
    df["is5"] = (df["S"] == 5)

    gH = df.groupby(["H", "A", "T"], dropna=False)
    df["AC"] = gH["H"].transform("size")
    df["AA"] = gH["is1"].transform("sum")
    df["AB"] = gH["is5"].transform("sum")

    gD = df.groupby(["D", "A", "T"], dropna=False)
    df["AF"] = gD["D"].transform("size")
    df["AD"] = gD["is1"].transform("sum")
    df["AE"] = gD["is5"].transform("sum")

    gC = df.groupby(["C", "A", "T"], dropna=False)
    df["AI"] = gC["C"].transform("size")
    df["AG"] = gC["is1"].transform("sum")
    df["AH"] = gC["is5"].transform("sum")

    assert df["AC"].min() >= 1 and df["AF"].min() >= 1 and df["AI"].min() >= 1, \
        "Internal error: a row's own group total came back as 0 -- should be impossible."

    def rating(a, b, c):
        raw = np.round((1 - ((a * 5 + b) / c)) * 10, 2)
        return np.where(raw < 0, np.nan, raw)

    df["AJ"] = rating(df["AA"], df["AB"], df["AC"])
    df["AK"] = rating(df["AD"], df["AE"], df["AF"])
    df["AL"] = rating(df["AG"], df["AH"], df["AI"])

    df["AM"] = np.round((df["AA"] + df["AB"]) / df["AC"], 4)
    df["AN"] = np.round((df["AD"] + df["AE"]) / df["AF"], 4)
    df["AO"] = np.round((df["AG"] + df["AH"]) / df["AI"], 4)

    df["AP"] = df.groupby("H")["AM"].transform("mean")
    df["AQ"] = df.groupby("D")["AN"].transform("mean")
    df["AR"] = df.groupby("C")["AO"].transform("mean")
    df["AS"] = df.groupby(["H", "A"])["AM"].transform("mean")
    df["AT"] = df.groupby(["D", "A"])["AN"].transform("mean")
    df["AU"] = df.groupby(["C", "A"])["AO"].transform("mean")

    n_div0 = int(df["AM"].isna().sum() + df["AN"].isna().sum() + df["AO"].isna().sum())
    print(f"  done in {time.time()-t0:.1f}s -- div-by-zero cases: {n_div0} (should be 0)")
    return df


def fnum(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    fv = float(v)
    if fv == int(fv):
        return str(int(fv))
    return repr(fv)


COLS_INT = ["AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI"]
COLS_FLOAT_BLANKABLE = ["AJ", "AK", "AL"]
COLS_FLOAT = ["AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU"]


def build_aa_av_xml(row_num: int, r: pd.Series, proj_map: dict,
                     vendor_map: dict, invit_map: dict, regional_head_map: dict,
                     unmapped_vendor: set, unmapped_invit: set, unmapped_regional: set,
                     ss: "SharedStringInterner") -> str:
    parts = []
    for col in COLS_INT:
        parts.append(f'<c r="{col}{row_num}"><v>{int(r[col])}</v></c>')
    for col in COLS_FLOAT_BLANKABLE:
        v = r[col]
        if pd.isna(v):
            parts.append(f'<c r="{col}{row_num}" t="str"><v></v></c>')
        else:
            parts.append(f'<c r="{col}{row_num}"><v>{fnum(v)}</v></c>')
    for col in COLS_FLOAT:
        parts.append(f'<c r="{col}{row_num}"><v>{fnum(r[col])}</v></c>')

    full_name = proj_map.get(r["A"], "#N/A")
    idx = ss.intern(full_name)
    parts.append(f'<c r="AV{row_num}" t="s"><v>{idx}</v></c>')

    # AW: Vendor/Inhouse, from the universal Vendor/Inhouse table (Asset
    # Type + Parameter), case-insensitive
    d_key = _norm(r["D"]) if pd.notna(r["D"]) else ""
    h_key = _norm(r["H"]) if pd.notna(r["H"]) else ""
    vendor_val = vendor_map.get((d_key, h_key))
    if vendor_val is None:
        unmapped_vendor.add((r["D"], r["H"]))
        vendor_val = "#N/A"
    idx = ss.intern(vendor_val)
    parts.append(f'<c r="AW{row_num}" t="s"><v>{idx}</v></c>')

    # AX: Invit/Non-Invit, from PROJECT_INVIT_MAP (Project Name)
    invit_val = invit_map.get(r["A"])
    if invit_val is None:
        unmapped_invit.add(r["A"])
        invit_val = "#N/A"
    idx = ss.intern(invit_val)
    parts.append(f'<c r="AX{row_num}" t="s"><v>{idx}</v></c>')

    # AY: Regional Head, from PROJECT_REGIONAL_HEAD_MAP (Project Name)
    regional_val = regional_head_map.get(r["A"])
    if regional_val is None:
        unmapped_regional.add(r["A"])
        regional_val = "#N/A"
    idx = ss.intern(regional_val)
    parts.append(f'<c r="AY{row_num}" t="s"><v>{idx}</v></c>')

    parts.append("</row>")
    return "".join(parts)


# ──────────────────────────────────────────────────────────────────────────
# Workbook-level patches
# ──────────────────────────────────────────────────────────────────────────

def patch_workbook_xml(xml: bytes, new_last_row: int, sheet5_last_row: int = None) -> bytes:
    # Turn OFF fullCalcOnLoad -- there's nothing left to recalculate in
    # Sheet1, and we don't want Excel forcing a recalc of the small handful
    # of formulas on other tabs (Sheet3/5/6/7/8) to somehow cascade either.
    if re.search(rb'<calcPr calcId="\d+"[^/]*/>', xml):
        xml = re.sub(
            rb'<calcPr calcId="(\d+)"[^/]*/>',
            lambda m: f'<calcPr calcId="{m.group(1).decode()}" calcMode="auto" fullCalcOnLoad="0"/>'.encode(),
            xml, count=1,
        )
    else:
        xml = re.sub(rb'</workbook>', b'<calcPr fullCalcOnLoad="0"/></workbook>', xml, count=1)
    xml = re.sub(
        rb'<definedName name="_xlnm\._FilterDatabase" localSheetId="0" hidden="1">'
        rb'Sheet1!\$A\$1:\$A[A-Z]\$\d+</definedName>',
        lambda m: re.sub(rb'\$A[A-Z]\$\d+', ('$AY$' + str(new_last_row)).encode(), m.group(0)),
        xml,
    )
    if sheet5_last_row is not None:
        # Sheet5 got rebuilt as a clean 3-column Asset Type/Parameter/Vendor
        # table -- fix its stale hidden filter-database name (it used to
        # point at whatever garbage range was there before, e.g. $A$1:$T$14483).
        xml = re.sub(
            rb'(<definedName name="_xlnm\._FilterDatabase" localSheetId="1" hidden="1">'
            rb'Sheet5!)\$[A-Z]+\$\d+:\$[A-Z]+\$\d+(</definedName>)',
            lambda m: m.group(1) + f'$A$1:$C${sheet5_last_row}'.encode() + m.group(2),
            xml,
        )
    return xml


def patch_content_types(xml: bytes) -> bytes:
    return re.sub(rb'<Override PartName="/xl/calcChain\.xml".*?/>', b"", xml)


def patch_workbook_rels(xml: bytes) -> bytes:
    return re.sub(rb'<Relationship [^>]*Target="calcChain\.xml"\s*/>', b"", xml)


def ensure_header_has_aw_ax_ay(head_bytes: bytes) -> bytes:
    """If the master's header row doesn't have AW1/AX1/AY1 yet (i.e. it only
    goes up to AV, from before this script's AW/AX/AY columns existed), add
    them so the output always has proper headers -- this is what lets you
    point this script straight at an old A:AV-only (or A:AX-only) master."""
    has_aw1 = b'<c r="AW1"' in head_bytes
    has_ax1 = b'<c r="AX1"' in head_bytes
    has_ay1 = b'<c r="AY1"' in head_bytes

    # Reuse the style index of AV1 (the column immediately before AW) so
    # AW1/AX1/AY1 always match the blue header formatting of AA1:AV1 -- NOT
    # the first styled cell in the row. Some early columns (e.g. the
    # "Month" header T1) carry unrelated styles like DATE_STYLE, and
    # picking the first match found those instead of the real header style.
    style_match = re.search(rb'<c r="AV1"[^>]*\bs="(\d+)"', head_bytes)
    if style_match:
        style_num = style_match.group(1).decode()
    else:
        all_matches = re.findall(rb'<c r="[A-Z]{1,2}1"[^>]*\bs="(\d+)"', head_bytes)
        style_num = all_matches[-1].decode() if all_matches else None

    def restyle_existing(hb: bytes, ref: str) -> bytes:
        """If ref (e.g. AW1) already exists, force its s="..." to style_num
        -- covers masters that already have AW1/AX1/AY1 from a run made
        before this styling fix existed, which would otherwise keep the bad
        style forever since we'd never hit the 'add new cell' branch below."""
        if style_num is None:
            return hb
        pattern = re.compile(rb'(<c r="' + ref.encode() + rb'")([^>]*)(>)')
        def repl(m):
            attrs = re.sub(rb'\s*s="\d+"', b'', m.group(2))
            return m.group(1) + f' s="{style_num}"'.encode() + attrs + m.group(3)
        return pattern.sub(repl, hb, count=1)

    if has_aw1:
        head_bytes = restyle_existing(head_bytes, "AW1")
    if has_ax1:
        head_bytes = restyle_existing(head_bytes, "AX1")
    if has_ay1:
        head_bytes = restyle_existing(head_bytes, "AY1")
    if has_aw1 and has_ax1 and has_ay1:
        return head_bytes

    style_attr = f' s="{style_num}"' if style_num else ""
    add = ""
    if not has_aw1:
        add += f'<c r="AW1"{style_attr} t="inlineStr"><is><t>Vendor/Inhouse</t></is></c>'
    if not has_ax1:
        add += f'<c r="AX1"{style_attr} t="inlineStr"><is><t>Invit/Non-Invit</t></is></c>'
    if not has_ay1:
        add += f'<c r="AY1"{style_attr} t="inlineStr"><is><t>Regional Head</t></is></c>'
    # header row is row 1; insert right before its closing </row>
    m = re.search(rb'(<row r="1"[^>]*>.*?)(</row>)', head_bytes, re.S)
    if not m:
        return head_bytes  # header row not in this chunk; nothing we can do here
    return head_bytes[:m.end(1)] + add.encode() + head_bytes[m.end(1):]


# ──────────────────────────────────────────────────────────────────────────
# Rebuild sharedStrings.xml (original entries untouched + newly interned
# strings appended) and a clean canonical Sheet5.
# ──────────────────────────────────────────────────────────────────────────

def _si_xml(s: str) -> str:
    text = "" if s is None else str(s)
    esc = html.escape(text)
    if text != text.strip() or text == "":
        return f'<si><t xml:space="preserve">{esc}</t></si>'
    return f'<si><t>{esc}</t></si>'


def build_shared_strings_xml(zin, ss: "SharedStringInterner") -> bytes:
    new_strings = ss.strings[ss._orig_len:]
    try:
        original = zin.read("xl/sharedStrings.xml")
    except KeyError:
        original = None

    if original is None:
        count = ss.new_ref_count
        unique = len(ss.strings)
        body = "".join(_si_xml(s) for s in ss.strings)
        xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{count}" uniqueCount="{unique}">{body}</sst>'
        )
        return xml.encode("utf-8")

    m_count = re.search(rb'<sst[^>]*\bcount="(\d+)"', original)
    m_unique = re.search(rb'<sst[^>]*\buniqueCount="(\d+)"', original)
    old_count = int(m_count.group(1)) if m_count else 0
    old_unique = int(m_unique.group(1)) if m_unique else len(ss.strings) - len(new_strings)
    new_count = old_count + ss.new_ref_count
    new_unique = old_unique + len(new_strings)

    patched = original
    if m_count:
        patched = re.sub(rb'(\bcount=")\d+(")', f'\\g<1>{new_count}\\g<2>'.encode(),
                          patched, count=1)
    if m_unique:
        patched = re.sub(rb'(\buniqueCount=")\d+(")', f'\\g<1>{new_unique}\\g<2>'.encode(),
                          patched, count=1)

    add_bytes = "".join(_si_xml(s) for s in new_strings).encode("utf-8")
    if b"</sst>" in patched:
        patched = patched.replace(b"</sst>", add_bytes + b"</sst>", 1)
    else:
        patched = patched + add_bytes  # shouldn't happen, but don't silently drop strings
    return patched


def build_clean_sheet5_xml(vendor_rows: list) -> bytes:
    """A small, clean Asset Type / Parameter / Vendor-Inhouse worksheet --
    tiny (a few hundred rows), so inlineStr here is fine (no bloat concern).
    Rewriting this every run is what keeps the universal table valid and
    self-consistent for next month's run, even if this month's master had a
    corrupted or missing Sheet5."""
    n = len(vendor_rows)
    rows_xml = []
    header = (
        '<row r="1"><c r="A1" t="inlineStr"><is><t>Asset Type</t></is></c>'
        '<c r="B1" t="inlineStr"><is><t>Parameter</t></is></c>'
        '<c r="C1" t="inlineStr"><is><t>Vendor/Inhouse</t></is></c></row>'
    )
    rows_xml.append(header)
    for i, (asset, param, verdict) in enumerate(vendor_rows, start=2):
        a = html.escape(str(asset))
        p = html.escape(str(param))
        v = html.escape(str(verdict))
        rows_xml.append(
            f'<row r="{i}">'
            f'<c r="A{i}" t="inlineStr"><is><t xml:space="preserve">{a}</t></is></c>'
            f'<c r="B{i}" t="inlineStr"><is><t xml:space="preserve">{p}</t></is></c>'
            f'<c r="C{i}" t="inlineStr"><is><t xml:space="preserve">{v}</t></is></c>'
            f'</row>'
        )
    last_row = n + 1
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:C{last_row}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="14.4"/>'
        f'<sheetData>{"".join(rows_xml)}</sheetData>'
        '</worksheet>'
    )
    return xml.encode("utf-8")


# ──────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    master_path = sys.argv[1]
    month_folder = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) >= 4 else \
        os.path.splitext(master_path)[0] + "_with_new_month.xlsx"

    t0 = time.time()

    category_map, asset_map, parameter_map, mode_map, overall_default, vendor_map, vendor_rows = \
        load_master_lookups(master_path)

    print("Loading raw month files...")
    month_folder_resolved = resolve_month_input(month_folder)
    raw = load_all_raw_files(month_folder_resolved)
    new_rows_df = compute_new_rows(raw, category_map, asset_map,
                                    parameter_map, mode_map, overall_default)
    check_month_consistency(new_rows_df, month_folder)
    check_duplicate_project_files(new_rows_df, month_folder)

    with zipfile.ZipFile(master_path, "r") as zin:
        names = zin.namelist()
        sheet1_xml = zin.read("xl/worksheets/sheet1.xml")
        workbook_xml = zin.read("xl/workbook.xml")
        content_types_xml = zin.read("[Content_Types].xml")
        workbook_rels_xml = zin.read("xl/_rels/workbook.xml.rels")

        last_row = load_master_last_row(sheet1_xml)
        print(f"Master currently has data through row {last_row:,}")

        shared_strings = load_shared_strings(zin)
        ss = SharedStringInterner(shared_strings)

        sheet5_path = find_real_sheet_file(workbook_xml, workbook_rels_xml, "Sheet5")

        sheet2_path = find_real_sheet_file(workbook_xml, workbook_rels_xml, "Sheet2")
        s2last = get_sheet_last_row(zin, sheet2_path)
        print(f"Resolved 'Sheet2' tab -> {sheet2_path} ({s2last:,} rows)")
        sheet2_xml = zin.read(sheet2_path) if sheet2_path else b""
        proj_map = build_project_lookup(sheet2_xml, shared_strings)
        print(f"  loaded {len(proj_map):,} project code -> full name mappings")

        pos = scan_row_positions(sheet1_xml)
        old_df = extract_old_rows(sheet1_xml, pos, shared_strings)

        n_new = len(new_rows_df)
        if n_new:
            new_df = pd.DataFrame({
                "row": range(last_row + 1, last_row + 1 + n_new),
                "A": new_rows_df["Project Name"].values,
                "C": new_rows_df["Category"].values,
                "D": new_rows_df["Asset Type"].values,
                "H": new_rows_df["Parameter"].values,
                "S": pd.to_numeric(new_rows_df["HO Final Rating"], errors="coerce").values,
                "T": [
                    (pd.Timestamp(m) - EXCEL_EPOCH).days if pd.notna(m) else np.nan
                    for m in new_rows_df["Month"]
                ],
            })
            combined = pd.concat([old_df, new_df], ignore_index=True)
        else:
            combined = old_df

        # unmapped-project sanity check up front (would show as #N/A in AV)
        unmapped = set(combined["A"].dropna().unique()) - set(proj_map.keys())
        if unmapped:
            print(f"  *** WARNING: {len(unmapped)} project code(s) not found in the "
                  f"Sheet2 lookup table -- their 'Project Name full' column will show "
                  f"#N/A: {sorted(unmapped)[:10]}")

        computed = compute_all_derived(combined)
        computed = computed.set_index("row", drop=False)

        new_last_row = last_row + n_new
        print(f"Writing {n_new:,} new row(s); final sheet will have {new_last_row:,} data rows.")

        unmapped_vendor = set()
        unmapped_invit = set()
        unmapped_regional = set()

        # Patch <dimension> in the small "head" slice (everything before row 1
        # -- a few hundred bytes) instead of regex-ing across the entire
        # (potentially huge) sheet, which is what used to force reading a
        # ~1GB temp file back into memory in one shot.
        head_slice = sheet1_xml[0:pos["row_tag_start"][0]]
        head_slice = re.sub(
            rb'<dimension ref="[A-Z]+1:[A-Z]+\d+"/>',
            f'<dimension ref="A1:AY{new_last_row}"/>'.encode(), head_slice, count=1)

        print(f"Streaming final sheet1.xml directly into the output .xlsx...")
        t_write = time.time()
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            with zout.open("xl/worksheets/sheet1.xml", "w") as out:
                out.write(head_slice)
                n_old_rows = len(pos["row_nums"])
                for i in range(n_old_rows):
                    if i == 0:
                        header_chunk = sheet1_xml[pos["row_tag_start"][i]:pos["row_end"][i]]
                        header_chunk = ensure_header_has_aw_ax_ay(header_chunk)
                        out.write(header_chunk)
                        last_end = pos["row_end"][i]
                        continue
                    out.write(sheet1_xml[last_end:pos["row_tag_start"][i]])
                    rn = pos["row_nums"][i]
                    out.write(sheet1_xml[pos["row_tag_start"][i]:pos["aa_start"][i]])
                    out.write(build_aa_av_xml(rn, computed.loc[rn], proj_map,
                                               vendor_map, PROJECT_INVIT_MAP, PROJECT_REGIONAL_HEAD_MAP,
                                               unmapped_vendor, unmapped_invit, unmapped_regional, ss).encode())
                    last_end = pos["row_end"][i]
                # tail of the original file (closes </sheetData> etc.) -- but
                # we still need to inject the brand-new rows BEFORE that tail.
                tail = sheet1_xml[last_end:]
                if n_new:
                    new_rows_dict = new_rows_df.to_dict("records")
                    for k in range(n_new):
                        rn = last_row + 1 + k
                        az_xml = build_new_row_az_xml(rn, new_rows_dict[k], ss)
                        out.write(az_xml.encode())
                        out.write(build_aa_av_xml(rn, computed.loc[rn], proj_map,
                                                   vendor_map, PROJECT_INVIT_MAP, PROJECT_REGIONAL_HEAD_MAP,
                                                   unmapped_vendor, unmapped_invit, unmapped_regional, ss).encode())
                tail = re.sub(
                    rb'<autoFilter ref="[A-Z]+1:[A-Z]+\d+"',
                    f'<autoFilter ref="A1:AY{new_last_row}"'.encode(), tail, count=1)
                out.write(tail)
            print(f"  wrote sheet1 in {time.time()-t_write:.1f}s "
                  f"({ss.new_ref_count:,} text-cell references written, "
                  f"{len(ss.strings) - len(shared_strings):,} brand-new distinct string(s))")

            if unmapped_vendor:
                print(f"  *** WARNING: {len(unmapped_vendor)} Asset Type/Parameter combo(s) not "
                      f"found in the Vendor/Inhouse table -- their 'Vendor/Inhouse' (AW) will "
                      f"show #N/A: {sorted(unmapped_vendor)[:10]}")
            if unmapped_invit:
                print(f"  *** WARNING: {len(unmapped_invit)} project code(s) not found in "
                      f"PROJECT_INVIT_MAP -- their 'Invit/Non-Invit' (AX) will show #N/A. Add "
                      f"them to PROJECT_INVIT_MAP at the top of this script: "
                      f"{sorted(x for x in unmapped_invit if x)[:10]}")
            if unmapped_regional:
                print(f"  *** WARNING: {len(unmapped_regional)} project code(s) not found in "
                      f"PROJECT_REGIONAL_HEAD_MAP -- their 'Regional Head' (AY) will show #N/A. "
                      f"Add them to PROJECT_REGIONAL_HEAD_MAP at the top of this script: "
                      f"{sorted(x for x in unmapped_regional if x)[:10]}")

            workbook_xml = patch_workbook_xml(workbook_xml, new_last_row)
            content_types_xml = patch_content_types(content_types_xml)
            workbook_rels_xml = patch_workbook_rels(workbook_rels_xml)

            skip = {"xl/calcChain.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml"}
            if sheet5_path:
                skip.add(sheet5_path)

            for item in names:
                if item in skip:
                    continue
                elif item == "xl/workbook.xml":
                    zout.writestr(item, workbook_xml)
                elif item == "[Content_Types].xml":
                    zout.writestr(item, content_types_xml)
                elif item == "xl/_rels/workbook.xml.rels":
                    zout.writestr(item, workbook_rels_xml)
                else:
                    zout.writestr(item, zin.read(item))

            # Write back sharedStrings.xml with any newly-added strings appended.
            zout.writestr("xl/sharedStrings.xml", build_shared_strings_xml(zin, ss))

            # Always (re)write a clean, canonical Sheet5 -- Asset Type /
            # Parameter / Vendor-Inhouse -- so that next month's run finds a
            # genuinely valid lookup table here regardless of what was in
            # this month's master. This is what makes the universal table
            # self-sustaining month over month instead of needing to be
            # re-seeded from Consolidated_June26.xlsx every time.
            if sheet5_path:
                zout.writestr(sheet5_path, build_clean_sheet5_xml(vendor_rows))
            else:
                print("  NOTE: could not resolve a 'Sheet5' tab in the master to write "
                      "the canonical Vendor/Inhouse table into -- Sheet5 was left as-is "
                      "(if present at all).")

    print(f"\nDone in {time.time()-t0:.1f}s -> {output_path}")
    print("The file has no formulas left in Sheet1 (AA:AY are all static values) "
          "and fullCalcOnLoad is off, so it will open instantly with everything "
          "already correct -- including Vendor/Inhouse, Invit/Non-Invit, and Regional Head.")


if __name__ == "__main__":
    main()

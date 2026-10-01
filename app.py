"""
ConsolidationConsole - Streamlit front end for merge_engine.py
----------------------------------------------------------------
Three-step wizard: upload the consolidated master, upload this month's
raw files (individual files OR a zip), name the output, and run the
merge. Wraps merge_engine.main() unchanged -- all the actual xlsx
surgery still happens in that file.

Run with:
    streamlit run app.py
"""

import contextlib
import io
import os
import shutil
import sys
import tempfile
import traceback

import streamlit as st

import merge_engine


# ──────────────────────────────────────────────────────────────────────────
# Page setup + styling
# ──────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="ConsolidationConsole",
    page_icon="🧩",
    layout="wide",
)

CSS = """
<style>
#MainMenu, footer, header[data-testid="stHeader"] {visibility: hidden; height: 0;}
.stApp {background: #fafbfc;}
.block-container {padding-top: 2.2rem; max-width: 1180px;}

/* ---------- top bar ---------- */
.cc-topbar {
    display: flex; align-items: center; justify-content: space-between;
    padding-bottom: 16px; border-bottom: 1px solid #e7e9ee; margin-bottom: 30px;
}
.cc-brand {display: flex; align-items: center; gap: 12px;}
.cc-logo {
    width: 38px; height: 38px; border-radius: 10px;
    background: linear-gradient(135deg, #6f6bf0, #2f8fe6);
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 2px 6px rgba(80,90,220,0.28);
}
.cc-brand-name {font-weight: 700; font-size: 17px; color: #12141c; line-height: 1.15;}
.cc-brand-sub {font-size: 12px; color: #9195a3;}
.cc-pill {
    font-size: 12.5px; padding: 5px 13px; border-radius: 999px; font-weight: 600;
    border: 1px solid #e2e4ea; color: #4a4e5c; display: inline-flex; align-items: center; gap: 6px;
    background: white;
}
.cc-pill.online {color: #17845a; border-color: #bfe8d6; background: #f1fbf6;}
.cc-dot {width: 7px; height: 7px; border-radius: 50%; background: #22c07a; display: inline-block;}

/* ---------- hero ---------- */
.cc-tag-row {display: flex; align-items: center; gap: 14px; margin-bottom: 16px;}
.cc-tag {
    font-size: 11px; font-weight: 700; letter-spacing: .06em; color: #5b57e8;
    background: #f0efff; border: 1px solid #dcdaff; border-radius: 6px;
    padding: 4px 10px; white-space: nowrap;
}
.cc-tag-line {flex: 1; height: 1px; background: #e7e9ee;}
.cc-title {font-size: 36px; font-weight: 800; color: #12141c; margin: 0 0 10px 0; letter-spacing: -0.01em;}
.cc-title span {color: #3f6df0;}
.cc-desc {color: #6b7080; font-size: 15.5px; margin-bottom: 30px; max-width: 720px; line-height: 1.5;}

/* ---------- step nav ---------- */
.cc-steps {display: flex; align-items: center; gap: 12px; margin-bottom: 34px;}
.cc-step {display: flex; align-items: center; gap: 9px; font-size: 14px; font-weight: 600;}
.cc-step-badge {
    width: 23px; height: 23px; border-radius: 50%; border: 1.6px solid;
    display: flex; align-items: center; justify-content: center; font-size: 12px; flex-shrink: 0;
}
.cc-step.blue {color: #3f6df0;}
.cc-step.blue .cc-step-badge {border-color: #3f6df0; color: #3f6df0;}
.cc-step.green {color: #1f9d63;}
.cc-step.green .cc-step-badge {border-color: #1f9d63; color: #1f9d63;}
.cc-step.done .cc-step-badge {background: #1f9d63; border-color: #1f9d63; color: white;}
.cc-step-line {flex: 1; height: 1px; background: #e2e4ea; min-width: 24px; max-width: 70px;}

/* ---------- cards ---------- */
.cc-card {
    border: 1px solid #e7e9ee; border-radius: 14px; padding: 24px 22px 8px 22px;
    background: white; box-shadow: 0 1px 3px rgba(20,20,40,0.03);
}
.cc-card-head {display: flex; align-items: center; gap: 8px; margin-bottom: 14px;}
.cc-card-step {
    width: 25px; height: 25px; border-radius: 50%; border: 1.6px solid; color: white;
    display: inline-flex; align-items: center; justify-content: center; font-size: 12.5px; font-weight: 700;
}
.cc-card.blue .cc-card-step {background: #3f6df0; border-color: #3f6df0;}
.cc-card.green .cc-card-step {background: #1f9d63; border-color: #1f9d63;}
.cc-card-eyebrow {font-size: 11.5px; font-weight: 700; letter-spacing: .06em;}
.cc-card.blue .cc-card-eyebrow {color: #3f6df0;}
.cc-card.green .cc-card-eyebrow {color: #1f9d63;}
.cc-card-title {font-size: 20px; font-weight: 700; color: #12141c; margin: 0 0 4px 0;}
.cc-card-desc {font-size: 13.5px; color: #8a8fa3; margin-bottom: 16px; line-height: 1.4;}

/* ---------- checklist ---------- */
.cc-check {display: flex; align-items: center; gap: 9px; font-size: 13.5px; color: #a7abb8; margin-bottom: 9px;}
.cc-check.done {color: #22222c; font-weight: 500;}
.cc-check-icon {
    width: 16px; height: 16px; border-radius: 50%; border: 1.6px solid #d7d9e2; flex-shrink: 0;
    display: flex; align-items: center; justify-content: center; font-size: 10px; color: white;
}
.cc-check.done .cc-check-icon {background: #1f9d63; border-color: #1f9d63;}
.cc-check-file {margin-left: auto; color: #9195a3; font-size: 11.5px; font-family: monospace;}

/* ---------- footer ---------- */
.cc-footer {text-align: center; color: #a7aab7; font-size: 12.5px; margin-top: 46px; line-height: 2;}
.cc-footer b {color: #8a8fa3; letter-spacing: .04em;}

/* ---------- native widget restyling ----------
   Selectors below deliberately omit a tag name (section/div/etc.) because
   the underlying element differs across Streamlit versions -- matching on
   [data-testid=...] alone keeps this working regardless of version. Every
   color is forced with !important so the light theme holds even if the
   browser/OS is in dark mode or a hosting environment ignores config.toml. */
:root, html, body {color-scheme: light !important;}
[data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"] {
    background: #fafbfc !important;
}

[data-testid="stFileUploaderDropzone"] {
    border-radius: 12px !important; border: 1.6px dashed #d7d9e2 !important;
    background: #fbfbfd !important; min-height: 170px !important;
    display: flex !important; flex-direction: column !important;
    align-items: center !important; justify-content: center !important; gap: 10px !important;
    padding: 18px !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] {
    flex-direction: column !important; text-align: center !important; gap: 2px !important;
    color: #12141c !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] * {color: #6b7080 !important; font-size: 13px !important;}
[data-testid="stBaseButton-secondary"] {
    background: #ffffff !important; border: 1.4px solid #d7d9e2 !important; border-radius: 8px !important;
    color: #12141c !important; font-weight: 700 !important;
}
[data-testid="stFileUploaderFile"], [data-testid="stFileUploaderFileName"] {color: #12141c !important;}
/* Hide the native per-file chip list -- we render our own one-line summary
   ("N files loaded") instead of a row per uploaded file. Different Streamlit
   versions use different testids for this, so cover both. */
[data-testid="stFileChips"],
[data-testid="stFileUploaderFile"] {display: none !important;}

.cc-upload-status {
    display: flex; align-items: center; gap: 8px; font-size: 13.5px; color: #12141c;
    font-weight: 600; margin-top: 12px;
}
.cc-upload-status .cc-check-icon {
    width: 16px; height: 16px; border-radius: 50%; background: #1f9d63; border: 1.6px solid #1f9d63;
    color: white; display: flex; align-items: center; justify-content: center; font-size: 10px; flex-shrink: 0;
}

/* text input (output filename) -- override every layer (root element,
   the react-aria wrapper, and the input itself) since which one actually
   carries the visible background differs across Streamlit versions. */
[data-testid="stTextInput"] div {
    background: #ffffff !important;
}
[data-testid="stTextInputRootElement"] {
    background: #ffffff !important; border: 1.4px solid #d7d9e2 !important; border-radius: 9px !important;
}
[data-testid="stTextInput"] input {
    background: transparent !important; color: #12141c !important; font-family: monospace !important;
    -webkit-text-fill-color: #12141c !important;
}
[data-testid="stTextInput"] input::placeholder {color: #9195a3 !important; opacity: 1 !important;}

/* segmented control (Individual files / Zip archive) -- BaseWeb radiogroup,
   testid stButtonGroup; the checked pill carries aria-checked="true". */
[data-testid="stButtonGroup"] button {
    background: #ffffff !important; border: 1.4px solid #d7d9e2 !important; border-radius: 9px !important;
    color: #4a4e5c !important; font-weight: 600 !important;
}
[data-testid="stButtonGroup"] button[aria-checked="true"] {
    background: rgba(63,109,240,0.08) !important; border-color: #3f6df0 !important; color: #3f6df0 !important;
}

/* Run merge button */
div.stButton > button[kind="primary"] {
    background: linear-gradient(90deg, #8b8bf0, #6070e0);
    border: none; border-radius: 10px; font-weight: 700; padding: 0.7rem 0; font-size: 15px;
    box-shadow: 0 4px 10px rgba(90,100,220,0.28);
}
div.stButton > button[kind="primary"]:hover {filter: brightness(1.04);}
div.stButton > button[kind="primary"]:disabled {
    background: linear-gradient(90deg, #c7cbf5, #b9c1ee); box-shadow: none; color: #f4f5ff;
}

/* ---------- progress bar ---------- */
.cc-dino-wrap {position: relative; margin: 22px 0 24px 0; padding-top: 20px;}
.cc-dino-track {
    width: 100%; height: 10px; border-radius: 999px; background: #e7e9ee; overflow: visible;
    position: relative;
}
.cc-dino-fill {
    height: 100%; border-radius: 999px;
    background: linear-gradient(90deg, #8b8bf0, #6070e0);
    transition: width 0.5s ease;
}
.cc-dino-pct {
    position: absolute; right: 0; top: -20px; font-size: 12px; font-weight: 700; color: #6070e0;
}
.cc-dino-stage {font-size: 13px; color: #6b7080; margin-top: 8px;}

/* ---------- expander header (force visible in dark OS/browser themes,
   same reasoning as the other native-widget overrides above) ---------- */
[data-testid="stExpander"] summary {
    background: #ffffff !important; color: #12141c !important; border-radius: 10px !important;
}
[data-testid="stExpander"] summary * {color: #12141c !important;}
[data-testid="stExpander"] summary svg {fill: #12141c !important;}
[data-testid="stExpander"] {
    background: #ffffff !important; border: 1px solid #e7e9ee !important; border-radius: 10px !important;
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

ICON_ARROW = """
<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M4 12l16-7-6 7 6 7-16-7z" fill="white"/>
</svg>
"""


# ──────────────────────────────────────────────────────────────────────────
# Session state
# ──────────────────────────────────────────────────────────────────────────

for k, v in {
    "output_name": "",
    "result_bytes": None,
    "result_name": None,
    "run_log": "",
    "raw_mode": "Individual files",
}.items():
    st.session_state.setdefault(k, v)


# ──────────────────────────────────────────────────────────────────────────
# Top bar + hero
# ──────────────────────────────────────────────────────────────────────────

st.markdown(
    f"""
    <div class="cc-topbar">
        <div class="cc-brand">
            <div class="cc-logo">{ICON_ARROW}</div>
            <div>
                <div class="cc-brand-name">ConsolidationConsole</div>
                <div class="cc-brand-sub">Consolidation Console</div>
            </div>
        </div>
    </div>
    <div class="cc-tag-row">
        <div class="cc-tag">INTERNAL TOOL</div>
        <div class="cc-tag-line"></div>
    </div>
    <div class="cc-title">Consolidation <span>Console</span></div>
    <div class="cc-desc">
        Merge this month's raw ratings into the consolidated master. Three steps —
        upload, configure, download.
    </div>
    """,
    unsafe_allow_html=True,
)


# ──────────────────────────────────────────────────────────────────────────
# Step indicator
# ──────────────────────────────────────────────────────────────────────────

def step_html(label, num, color, done):
    cls = f"cc-step {color}" + (" done" if done else "")
    badge = "✓" if done else str(num)
    return f'<div class="{cls}"><div class="cc-step-badge">{badge}</div>{label}</div>'

# Filled in after the file uploaders (below) have run this rerun, so the
# badges reflect the files just picked instead of last rerun's state.
step_indicator = st.empty()


# ──────────────────────────────────────────────────────────────────────────
# Three-column wizard
# ──────────────────────────────────────────────────────────────────────────

col1, col2, col3 = st.columns(3, gap="medium")

# ---- Step 1: master file ----
with col1:
    st.markdown(
        """
        <div class="cc-card blue">
        <div class="cc-card-head"><span class="cc-card-step">1</span><span class="cc-card-eyebrow">STEP 1</span></div>
        <div class="cc-card-title">Consolidated master</div>
        <div class="cc-card-desc">The workbook as of last month — e.g. <span style="font-family:monospace;">Consolidated_Apr_thru_Jun.xlsx</span></div>
        """,
        unsafe_allow_html=True,
    )
    master_upload = st.file_uploader(
        "Master workbook", type=["xlsx"], key="master_upload",
        label_visibility="collapsed",
    )
    if master_upload is not None:
        st.markdown(
            f'<div class="cc-upload-status"><span class="cc-check-icon">✓</span>'
            f'{master_upload.name}</div>',
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

# ---- Step 2: raw files ----
with col2:
    st.markdown(
        """
        <div class="cc-card blue">
        <div class="cc-card-head"><span class="cc-card-step">2</span><span class="cc-card-eyebrow">STEP 2</span></div>
        <div class="cc-card-title">This month's raw files</div>
        <div class="cc-card-desc">Choose one of the two — not both.</div>
        """,
        unsafe_allow_html=True,
    )
    raw_mode = st.segmented_control(
        "Raw file mode", options=["Individual files", "Zip archive"],
        default=st.session_state["raw_mode"], label_visibility="collapsed",
        key="raw_mode_widget",
    ) or st.session_state["raw_mode"]
    st.session_state["raw_mode"] = raw_mode

    raw_individual_files = None
    raw_zip_file = None
    if raw_mode == "Individual files":
        raw_individual_files = st.file_uploader(
            "Raw files", type=["csv", "xlsx"], accept_multiple_files=True,
            key="raw_individual", label_visibility="collapsed",
        )
        raw_ready = bool(raw_individual_files)
        if raw_individual_files:
            n = len(raw_individual_files)
            st.markdown(
                f'<div class="cc-upload-status"><span class="cc-check-icon">✓</span>'
                f'All {n} file{"s" if n != 1 else ""} loaded</div>',
                unsafe_allow_html=True,
            )
    else:
        raw_zip_file = st.file_uploader(
            "Zip archive", type=["zip"], key="raw_zip",
            label_visibility="collapsed",
        )
        raw_ready = raw_zip_file is not None
        if raw_zip_file is not None:
            st.markdown(
                f'<div class="cc-upload-status"><span class="cc-check-icon">✓</span>'
                f'{raw_zip_file.name} loaded</div>',
                unsafe_allow_html=True,
            )
    st.session_state["_raw_ready"] = raw_ready
    st.markdown("</div>", unsafe_allow_html=True)

step_indicator.markdown(
    f"""
    <div class="cc-steps">
        {step_html("Master file", 1, "blue", master_upload is not None)}
        <div class="cc-step-line"></div>
        {step_html("Raw files", 2, "blue", raw_ready)}
        <div class="cc-step-line"></div>
        {step_html("Run merge", 3, "green", False)}
    </div>
    """,
    unsafe_allow_html=True,
)

# ---- Step 3: run + download ----
with col3:
    st.markdown(
        """
        <div class="cc-card green">
        <div class="cc-card-head"><span class="cc-card-step">3</span><span class="cc-card-eyebrow">STEP 3</span></div>
        <div class="cc-card-title">Run &amp; download</div>
        <div class="cc-card-desc">Name the output file, then kick off the merge.</div>
        """,
        unsafe_allow_html=True,
    )

    default_name = ""
    if master_upload is not None:
        base = os.path.splitext(master_upload.name)[0]
        default_name = f"{base}_with_new_month.xlsx"

    output_name = st.text_input(
        "Output filename",
        value=st.session_state["output_name"] or default_name,
        placeholder="Consolidated_Apr_thru_Jul.xlsx",
    )
    if output_name and not output_name.lower().endswith(".xlsx"):
        output_name += ".xlsx"
    st.session_state["output_name"] = output_name

    master_present = master_upload is not None
    checks = [
        ("Consolidated master loaded", master_present, None),
        ("Raw files queued", raw_ready, None),
        ("Output filename set", bool(output_name), output_name),
    ]
    check_rows = []
    for label, done, right in checks:
        cls = "done" if done else ""
        icon = "✓" if done else ""
        right_html = f'<span class="cc-check-file">{right}</span>' if (done and right) else ""
        check_rows.append(
            f'<div class="cc-check {cls}"><span class="cc-check-icon">{icon}</span>{label}{right_html}</div>'
        )
    st.markdown("".join(check_rows), unsafe_allow_html=True)
    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

    ready = master_present and raw_ready and bool(output_name)
    run_clicked = st.button(
        "▶  Run merge", type="primary", use_container_width=True, disabled=not ready,
    )
    st.markdown("<div style='height:6px;'></div></div>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────
# Run the merge
# ──────────────────────────────────────────────────────────────────────────

class _LiveLog(io.StringIO):
    """Like io.StringIO, but also pushes every write into a Streamlit
    placeholder immediately, so the person watching sees merge_engine's
    print() output appear line by line instead of only after main()
    returns (or dies). Without this, a genuinely-working multi-minute
    run on a large master looks identical to a frozen one."""
    def __init__(self, placeholder):
        super().__init__()
        self._placeholder = placeholder

    def write(self, s):
        n = super().write(s)
        tail = "\n".join(self.getvalue().splitlines()[-40:])
        self._placeholder.code(tail or "(waiting for merge_engine output...)", language=None)
        return n


# Ordered, in the same sequence merge_engine.main() actually prints them.
# Each entry's percent is a rough map of "how far through main() are we"
# based on which of these lines has appeared in the log so far -- there's
# no true progress fraction available from the engine itself, so this is
# a stage-based approximation, not an exact byte/row count.
_STAGE_MARKERS = [
    ("Loading Sheet3 code tables", 5, "Loading code tables..."),
    ("Learning PC/FC/CC dominant class", 10, "Learning classification rules..."),
    ("Loading Sheet5 Vendor/Inhouse table", 15, "Loading vendor/inhouse table..."),
    ("Loading raw month files", 20, "Loading this month's raw files..."),
    ("raw rows from", 30, "Raw files loaded..."),
    ("Computing derived columns (R:Z)", 40, "Computing derived columns..."),
    ("new rows ready", 48, "New rows ready..."),
    ("Master currently has data through row", 55, "Reading existing master data..."),
    ("Resolved 'Sheet2' tab", 60, "Resolving project lookup table..."),
    ("Extracting existing rows' Project", 68, "Extracting existing rows..."),
    ("existing rows in", 76, "Existing rows extracted..."),
    ("Recomputing AA:AV for every row", 84, "Recomputing derived columns for every row..."),
    ("div-by-zero cases", 90, "Derived columns done..."),
    ("final sheet will have", 93, "Writing new rows..."),
    ("Streaming final sheet1.xml directly", 96, "Streaming output file..."),
    ("wrote sheet1 in", 98, "Output sheet written..."),
    ("Done in", 100, "Done."),
]


def _render_dino_bar(placeholder, log_text: str):
    pct, stage_label = 2, "Starting..."
    for marker, marker_pct, label in _STAGE_MARKERS:
        if marker in log_text:
            pct, stage_label = marker_pct, label
    placeholder.markdown(
        f"""
        <div class="cc-dino-wrap">
            <span class="cc-dino-pct">{pct}%</span>
            <div class="cc-dino-track"><div class="cc-dino-fill" style="width:{pct}%;"></div></div>
            <div class="cc-dino-stage">{stage_label}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


class _LiveLog(io.StringIO):
    """Like io.StringIO, but also pushes every write into the dino
    progress bar and (if expanded) the raw log placeholder immediately,
    so the person watching sees real progress instead of a static
    spinner caption. Without this, a genuinely-working multi-minute run
    on a large master looks identical to a frozen one."""
    def __init__(self, dino_placeholder, log_placeholder):
        super().__init__()
        self._dino_placeholder = dino_placeholder
        self._log_placeholder = log_placeholder

    def write(self, s):
        n = super().write(s)
        text = self.getvalue()
        _render_dino_bar(self._dino_placeholder, text)
        tail = "\n".join(text.splitlines()[-40:])
        self._log_placeholder.code(tail or "(waiting for merge_engine output...)", language=None)
        return n


def _run_merge(master_path: str, month_input: str, output_path: str,
                dino_placeholder, log_placeholder) -> str:
    """Call merge_engine.main() unchanged, via argv, streaming its stdout
    live into the dino progress bar and the (collapsed-by-default) log.

    merge_engine.main() calls sys.exit(1) on certain validation failures
    (e.g. raw files spanning more than one month) after printing a
    human-readable reason to stdout. SystemExit is a BaseException, not
    an Exception, so it is NOT caught by a plain `except Exception:` --
    left unhandled, it silently aborts the Streamlit script run with no
    spinner resolution, no error box, and no log (this is why "Run merge"
    can appear to do nothing). We catch it here and re-raise as a normal
    RuntimeError carrying whatever the engine already printed, so the
    caller's except block can show it to the user.
    """
    old_argv = sys.argv
    sys.argv = ["merge_engine.py", master_path, month_input, output_path]
    buf = _LiveLog(dino_placeholder, log_placeholder)
    try:
        with contextlib.redirect_stdout(buf):
            merge_engine.main()
    except SystemExit as e:
        log = buf.getvalue().strip()
        raise RuntimeError(log or f"merge_engine exited with code {e.code} "
                                   f"and no further explanation.") from None
    finally:
        sys.argv = old_argv
    return buf.getvalue()


if run_clicked:
    work_dir = tempfile.mkdtemp(prefix="consolidationconsole_")
    log_placeholder = None  # created once we reach the logging UI below;
                             # stays None if we fail before that point.
    try:
        master_path = os.path.join(work_dir, "master.xlsx")
        with open(master_path, "wb") as f:
            f.write(master_upload.getbuffer())

        if raw_mode == "Individual files":
            month_dir = os.path.join(work_dir, "raw_month")
            os.makedirs(month_dir, exist_ok=True)
            for uf in raw_individual_files:
                with open(os.path.join(month_dir, uf.name), "wb") as f:
                    f.write(uf.getbuffer())
            month_input = month_dir
        else:
            zip_path = os.path.join(work_dir, "raw_month.zip")
            with open(zip_path, "wb") as f:
                f.write(raw_zip_file.getbuffer())
            month_input = zip_path

        output_path = os.path.join(work_dir, output_name)
        dino_placeholder = st.empty()
        _render_dino_bar(dino_placeholder, "")
        with st.expander("Show logs", expanded=False):
            log_placeholder = st.empty()
            log_placeholder.code("(waiting for merge_engine output...)", language=None)

        log_text = _run_merge(master_path, month_input, output_path,
                               dino_placeholder, log_placeholder)
        st.session_state["run_log"] = log_text

        with open(output_path, "rb") as f:
            st.session_state["result_bytes"] = f.read()
        st.session_state["result_name"] = output_name

        st.success(f"Merge complete → {output_name}")

    except RuntimeError as e:
        # Raised by _run_merge when merge_engine hit a validation check
        # (sys.exit) -- the message is the engine's own printed reason,
        # so show it as-is rather than a Python traceback.
        st.session_state["result_bytes"] = None
        st.error("The merge was stopped by a validation check. Details below.")
        st.session_state["run_log"] = (st.session_state.get("run_log") or "") + "\n" + str(e)
    except Exception:
        st.session_state["result_bytes"] = None
        st.error("The merge failed. Details below.")
        st.session_state["run_log"] = (st.session_state.get("run_log") or "") + "\n" + traceback.format_exc()
    except BaseException:
        # Last-resort net: catches anything that isn't a normal Exception
        # (e.g. a SystemExit from a code path we haven't wrapped yet) so
        # the run always ends with visible feedback instead of silently
        # dying mid-script.
        st.session_state["result_bytes"] = None
        st.error("The merge stopped unexpectedly. Details below.")
        st.session_state["run_log"] = (st.session_state.get("run_log") or "") + "\n" + traceback.format_exc()
    finally:
        # Paint the accumulated log (including any traceback just added
        # above) immediately, in this same run -- previously it was only
        # stored in session_state and didn't appear on screen until the
        # user triggered a second, unrelated rerun (e.g. clicking Run
        # merge again), which made failures look like they had "no
        # details" the first time.
        final_log = st.session_state.get("run_log")
        if final_log:
            if log_placeholder is not None:
                log_placeholder.code(final_log, language=None)
            else:
                # We failed before the "Show logs" expander was even
                # created (e.g. while saving an uploaded file) -- open
                # a fresh one now so the error still isn't silent.
                with st.expander("Show logs", expanded=True):
                    st.code(final_log, language=None)
        shutil.rmtree(work_dir, ignore_errors=True)


if st.session_state.get("result_bytes"):
    st.download_button(
        "⬇ Download merged workbook",
        data=st.session_state["result_bytes"],
        file_name=st.session_state["result_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if not run_clicked and st.session_state.get("run_log"):
    with st.expander("Show logs", expanded=False):
        st.code(st.session_state["run_log"] or "(no output)", language=None)


st.markdown(
    """
    <div class="cc-footer">
        <b>CONSOLIDATIONCONSOLE</b> · INTERNAL OPERATIONS TOOL
    </div>
    """,
    unsafe_allow_html=True,
)

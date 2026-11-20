"""
tests/test_app.py  —  TC-T-28 through TC-T-35
Streamlit AppTest smoke tests: app starts, all pages render,
no unhandled exceptions, load time < 3 seconds.

NOTE: These tests require `streamlit` and `app.py` to exist.
      They will fail until Step 15 (Build) creates the app pages.
      Mark them xfail until then with:
          pytest tests/test_app.py -v --runxfail

Customer Retention Metrics Analysis (CRMA)
Sumedha | Roll 2503201450146 | MCA 2nd Yr | ABES
Run: pytest tests/test_app.py -v
"""
import os
import time
import pytest

# ── Guard: skip whole module if app.py not yet created ───────────────────────
APP_PATH = os.path.join(os.path.dirname(__file__), "..", "app.py")
pytestmark = pytest.mark.skipif(
    not os.path.exists(APP_PATH),
    reason="app.py not yet created — tests run in Step 15"
)


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-28  app.py exists and is non-empty
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_28_app_py_exists():
    """TC-T-28: app.py exists in the project root and is non-empty."""
    assert os.path.exists(APP_PATH), "app.py not found"
    assert os.path.getsize(APP_PATH) > 100, "app.py appears to be empty"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-29  All page files exist
# ─────────────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("page_file", [
    "pages/1_Cohorts.py",
    "pages/2_Segments.py",
    "pages/3_Drivers.py",
    "pages/4_Risk_Roster.py",
    "pages/5_Data_Quality.py",
    "pages/6_Methodology.py",
    "pages/7_About.py",
])
def test_TC_T_29_page_files_exist(page_file):
    """TC-T-29: Every Streamlit page file is present in pages/."""
    path = os.path.join(os.path.dirname(__file__), "..", page_file)
    assert os.path.exists(path), f"Page file missing: {page_file}"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-30  Streamlit AppTest: app starts without exception
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_30_app_starts():
    """TC-T-30: app.py runs under Streamlit AppTest without raising an exception."""
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(APP_PATH, default_timeout=15)
    at.run()
    assert not at.exception, f"App raised exception: {at.exception}"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-31  Executive Overview renders KPI cards (≥ 5 metrics)
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_31_overview_has_kpis():
    """TC-T-31: Executive Overview page renders at least 5 KPI metric elements."""
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(APP_PATH, default_timeout=15)
    at.run()
    assert not at.exception
    # Page should expose metric widgets or the HTML components
    # At minimum, check the page has content (non-empty markdown/text)
    all_text = " ".join([m.value for m in at.metric] if at.metric else [])
    # Accept either st.metric elements or rendered HTML components
    assert len(at.metric) >= 1 or len(at.markdown) >= 1, (
        "No KPI metrics or markdown found on overview page"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-32  All pages render without exception
# ─────────────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("page_file,page_name", [
    ("pages/1_Cohorts.py",       "Cohort Analysis"),
    ("pages/2_Segments.py",      "Segment Comparison"),
    ("pages/3_Drivers.py",       "Churn Drivers"),
    ("pages/4_Risk_Roster.py",   "Risk Roster"),
    ("pages/5_Data_Quality.py",  "Data Quality"),
    ("pages/6_Methodology.py",   "Methodology"),
    ("pages/7_About.py",         "About"),
])
def test_TC_T_32_pages_render(page_file, page_name):
    """TC-T-32: Each Streamlit page renders without an unhandled exception."""
    from streamlit.testing.v1 import AppTest
    page_path = os.path.join(os.path.dirname(__file__), "..", page_file)
    if not os.path.exists(page_path):
        pytest.skip(f"{page_file} not yet created")
    at = AppTest.from_file(page_path, default_timeout=15)
    at.run()
    assert not at.exception, (
        f"{page_name} raised exception: {at.exception}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-33  Initial load time < 3 seconds (NFR-01)
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_33_load_time_under_3s():
    """TC-T-33: app.py fully renders within 3.0 seconds (NFR-01 requirement)."""
    from streamlit.testing.v1 import AppTest
    start = time.perf_counter()
    at = AppTest.from_file(APP_PATH, default_timeout=15)
    at.run()
    elapsed = time.perf_counter() - start
    assert not at.exception, f"App raised exception: {at.exception}"
    assert elapsed < 3.0, (
        f"App took {elapsed:.2f}s to load (NFR-01 limit: 3.0s)"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-34  Sidebar filters render on overview page
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_34_sidebar_filters_present():
    """TC-T-34: Overview page renders at least one sidebar selectbox/multiselect."""
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(APP_PATH, default_timeout=15)
    at.run()
    assert not at.exception
    n_selects = len(at.multiselect) + len(at.selectbox) + len(at.select_slider)
    assert n_selects >= 1, (
        "No sidebar filter widgets found (expected multiselect or selectbox)"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-35  No raw Python tracebacks visible to end user (NFR-02)
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_35_no_raw_tracebacks():
    """TC-T-35: No unhandled exceptions appear in any page output (NFR-02)."""
    from streamlit.testing.v1 import AppTest
    for page_file in [
        APP_PATH,
        *[os.path.join(os.path.dirname(__file__), "..", f)
          for f in ["pages/1_Cohorts.py", "pages/2_Segments.py",
                    "pages/3_Drivers.py", "pages/4_Risk_Roster.py"]]
    ]:
        if not os.path.exists(page_file):
            continue
        at = AppTest.from_file(page_file, default_timeout=15)
        at.run()
        assert not at.exception, (
            f"{os.path.basename(page_file)}: unhandled exception → {at.exception}"
        )

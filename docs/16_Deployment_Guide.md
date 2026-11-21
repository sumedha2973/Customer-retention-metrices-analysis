# Streamlit Community Cloud Deployment Guide
**Project:** Customer Retention Metrics Analysis (CRMA)  
**Author:** Sumedha | Roll: 2500320140149 | MCA 2nd Year, ABES Engineering College  
**Target Platform:** Streamlit Community Cloud (`share.streamlit.io`)  

---

## 1. Pre-Deployment Repository Audit Checklist

Every item below has been audited and prepared in the local Git repository:

- [x] **Root Entrypoint:** `app.py` is located in the repository root directory.
- [x] **Pinned Dependencies:** `requirements.txt` is in the root directory with strictly pinned versions (Python 3.12 compatible).
- [x] **Database File Committed:** Embedded SQLite file `data/app.db` (455 KB) is committed and tracked in Git.
- [x] **OS-Agnostic Relative Paths:** `src/data_loader.py` uses `Path(__file__).resolve().parent.parent / "data" / "app.db"` with POSIX path resolution, guaranteeing zero path failures on Linux cloud containers.
- [x] **Zero Secret Leaks:** No API keys, credentials, or private access tokens exist in the repository.
- [x] **Clean `.gitignore`:** Configured to exclude `.venv/`, `__pycache__/`, temporary Office lock files (`~$*`), and test caches.
- [x] **Git Branch:** Renamed and committed to `main` branch.

---

## 2. Step-by-Step Cloud Deployment Runbook

### Step 1: Push Local Repository to GitHub
1. Create a new GitHub repository under your personal GitHub account named:  
   `2500320140149-customer-retention-analysis`
2. Link your local repository and push:
   ```bash
   git remote add origin https://github.com/<your-github-username>/2500320140149-customer-retention-analysis.git
   git branch -M main
   git push -u origin main
   ```

### Step 2: Connect to Streamlit Community Cloud
1. Navigate to [share.streamlit.io](https://share.streamlit.io/) in your browser.
2. Sign in using your GitHub account that owns the repository.
3. Click the **"Create app"** button (top right).

### Step 3: Configure Deployment Parameters
Fill in the deployment modal:
* **Repository:** `<your-github-username>/2500320140149-customer-retention-analysis`
* **Branch:** `main`
* **Main file path:** `app.py`
* **App URL (Custom Subdomain):** `2500320140149-crma`  
  *(Yields public live URL: `https://2500320140149-crma.streamlit.app`)*

### Step 4: Deploy & Monitor Build Log
1. Click **"Deploy!"**.
2. Streamlit Cloud spins up an isolated Debian/Ubuntu container, reads `requirements.txt`, installs packages, and starts Streamlit.
3. Observe the live build log in the bottom-right terminal drawer (typical install time: 2–3 minutes).
4. Verify the application displays the Executive Overview page with HTTP 200.

---

## 3. Cloud Troubleshooting & Error Resolution Matrix

| Symptom / Log Error | Root Cause | Exact Resolution |
| :--- | :--- | :--- |
| `ModuleNotFoundError: No module named 'src'` | Python path not discovering root directory on Linux container. | Ensure imports in `app.py` and `pages/*.py` use `from src.data_loader import ...`. Streamlit Cloud adds repo root to `sys.path` by default. |
| `sqlite3.OperationalError: unable to open database file` | Hardcoded Windows backslashes (`\`) or missing DB. | Resolved via `Path(__file__).parent.parent / "data" / "app.db"` and `as_posix()` URI formatting in `src/data_loader.py`. |
| `MemoryLimitExceededError` (Container 1GB cap) | In-memory DataFrames too large or repeated full DB scans. | Resolved via `@st.cache_data(ttl=3600)` and categorical downcasting (`category` dtype) implemented in `src/data_loader.py`. |
| `Tailwind components not rendering / blank iframe` | CDN script blocked by restrictive browser Content Security Policy. | `src/ui.py` injects `<script src="https://cdn.tailwindcss.com"></script>` directly into `st.components.v1.html()` iframe, isolating stylesheet evaluation. |

---

## 4. Verification Check
Once deployed, verify:
1. Executive Overview loads 5 top Tailwind KPI cards and Plotly trends.
2. Navigation sidebar opens `1_Cohorts`, `2_Segments`, `3_Drivers`, `4_Risk_Roster`, `5_Data_Quality`, `6_Methodology`, and `7_About`.
3. Risk Roster "Download CSV" button successfully triggers a browser file download.

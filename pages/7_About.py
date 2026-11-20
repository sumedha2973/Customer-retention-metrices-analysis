"""
pages/7_About.py — Screen UI-07: Academic Attribution & Metadata
Customer Retention Metrics Analysis (CRMA)
"""
import streamlit as st

st.set_page_config(page_title="CRMA — About", page_icon="ℹ️", layout="wide")

st.title("ℹ️ About Customer Retention Metrics Analysis (CRMA)")
st.caption("Master of Computer Applications (MCA) — Academic Capstone Project")

c1, c2 = st.columns([1, 1])

with c1:
    st.markdown("### 🎓 Academic Context & Attribution")
    st.markdown("""
    * **Learner / Author:** Sumedha
    * **Roll Number:** 2503201450146
    * **Academic Program:** Master of Computer Applications (MCA), 2nd Year
    * **Specialization:** Data Analytics (DA)
    * **Institution:** ABES Engineering College, Ghaziabad
    * **Faculty Mentor:** Ms. Anjali Pandey
    * **Submission Phase:** End-to-End Professional 18-Step Build (Step 15: Build, Test, Fix)
    """)

    st.markdown("### 🏢 Business Scenario")
    st.markdown("""
    * **Fictitious Enterprise:** NoidaSoft Technologies Pvt. Ltd.
    * **Industry Domain:** B2B SaaS Enterprise Cloud Solutions
    * **Business Problem:** Identifying early warning indicators of account churn, diagnosing SLA support bottlenecks, and prioritizing proactive customer success outreach across a 24-month contract history.
    """)

with c2:
    st.markdown("### 🛠 Technical Architecture & Stack")
    st.markdown("""
    * **Core Language:** Python 3.12 / 3.13
    * **Web Framework:** Streamlit 1.32.2 (Multipage Architecture)
    * **Database Storage:** Embedded SQLite 3 (`data/app.db`) in read-only URI mode
    * **Data Wrangling:** pandas 2.2.1, NumPy 1.26.4
    * **Scientific Computing:** SciPy 1.12.0 (Hypothesis Testing & Correlation Analysis)
    * **Visualizations:** Plotly 5.20.0 (Interactive Heatmaps, Dual-Axis Time Series, Grouped Bars)
    * **Front-End Styling:** Tailwind CSS CDN loaded via Streamlit iframe component bridge
    * **Quality Assurance:** pytest 8.1.1, pytest-cov 5.0.0 (35 automated technical unit & smoke tests)
    * **Deployment Target:** Streamlit Community Cloud
    """)

st.markdown("---")
st.info("📌 **Open Source Repository:** All project documents (Steps 1–14), database schemas, synthetic datasets, and source code are maintained under strict version control.")

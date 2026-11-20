"""
src/ui.py — Tailwind-styled UI components for CRMA
Renders HTML via st.components.v1.html() with Tailwind CDN.

Customer Retention Metrics Analysis
Sumedha | Roll 2503201450146 | MCA 2nd Yr | ABES
"""
import streamlit.components.v1 as components


def render_header(title: str = "Customer Retention Metrics Analysis",
                  subtitle: str = "NoidaSoft Technologies Pvt. Ltd."):
    """Render the app header with Tailwind gradient styling."""
    html = f"""
    <script src="https://cdn.tailwindcss.com"></script>
    <div class="bg-gradient-to-r from-[#1B365D] to-[#2B4C7E]
                rounded-xl shadow-lg px-6 py-4 mb-4">
        <h1 class="text-2xl font-bold text-white tracking-tight">
            {title}
        </h1>
        <p class="text-sm text-blue-200 mt-1">{subtitle}</p>
    </div>
    """
    components.html(html, height=90)


def render_kpi_card(title: str, value: str, delta: str = "",
                    status: str = "neutral"):
    """Render a single KPI metric card with delta badge.

    Args:
        title:  Metric label (e.g., "Churn Rate")
        value:  Formatted value (e.g., "2.3%")
        delta:  Change text (e.g., "▼ 0.2 pp")
        status: "positive" (green), "negative" (red), or "neutral" (grey)
    """
    color_map = {
        "positive": ("text-green-600", "bg-green-50"),
        "negative": ("text-red-500",   "bg-red-50"),
        "neutral":  ("text-slate-500", "bg-slate-50"),
    }
    text_color, badge_bg = color_map.get(status, color_map["neutral"])

    delta_html = ""
    if delta:
        delta_html = f"""
        <span class="inline-flex items-center px-2 py-0.5
                      rounded-full text-xs font-medium
                      {badge_bg} {text_color}">
            {delta}
        </span>
        """

    html = f"""
    <script src="https://cdn.tailwindcss.com"></script>
    <div class="bg-white rounded-xl shadow-sm border border-slate-200
                p-4 hover:shadow-md transition-shadow duration-200">
        <p class="text-xs font-medium text-slate-500 uppercase
                  tracking-wider mb-1">{title}</p>
        <div class="flex items-end gap-2">
            <p class="text-3xl font-bold text-slate-800">{value}</p>
            {delta_html}
        </div>
    </div>
    """
    components.html(html, height=110)


def render_kpi_row(metrics: list[dict]):
    """Render a row of KPI cards in a responsive grid.

    Args:
        metrics: List of dicts with keys: title, value, delta, status
                 Example: [{"title": "Churn Rate", "value": "2.3%",
                            "delta": "▼ 0.2 pp", "status": "positive"}]
    """
    cards_html = ""
    for m in metrics:
        text_color, badge_bg = {
            "positive": ("text-green-600", "bg-green-50"),
            "negative": ("text-red-500",   "bg-red-50"),
            "neutral":  ("text-slate-500", "bg-slate-50"),
        }.get(m.get("status", "neutral"), ("text-slate-500", "bg-slate-50"))

        delta_html = ""
        if m.get("delta"):
            delta_html = f"""
            <span class="inline-flex items-center px-2 py-0.5
                          rounded-full text-xs font-medium
                          {badge_bg} {text_color}">
                {m['delta']}
            </span>
            """

        cards_html += f"""
        <div class="bg-white rounded-xl shadow-sm border border-slate-200
                    p-4 hover:shadow-md transition-shadow duration-200">
            <p class="text-xs font-medium text-slate-500 uppercase
                      tracking-wider mb-1">{m['title']}</p>
            <div class="flex items-end gap-2">
                <p class="text-3xl font-bold text-slate-800">{m['value']}</p>
                {delta_html}
            </div>
        </div>
        """

    n_cols = min(len(metrics), 4)
    html = f"""
    <script src="https://cdn.tailwindcss.com"></script>
    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-{n_cols}
                gap-4 mb-4">
        {cards_html}
    </div>
    """
    height = 130 if len(metrics) <= 4 else 250
    components.html(html, height=height)


def render_stat_card(test_name: str, stat_value: float, p_value: float,
                     interpretation: str, alpha: float = 0.05):
    """Render a statistical test result card.

    Args:
        test_name:      Name of the test (e.g., "Mann-Whitney U")
        stat_value:     The test statistic
        p_value:        The p-value
        interpretation: Plain-English interpretation
        alpha:          Significance level (default 0.05)
    """
    is_sig = p_value < alpha
    badge_color = "bg-green-100 text-green-800" if is_sig else "bg-amber-100 text-amber-800"
    badge_icon = "✓" if is_sig else "✗"
    badge_text = f"Significant (p = {p_value:.4f})" if is_sig else f"Not Significant (p = {p_value:.4f})"

    html = f"""
    <script src="https://cdn.tailwindcss.com"></script>
    <div class="bg-gradient-to-r from-blue-50 to-indigo-50
                rounded-lg p-4 border-l-4 border-indigo-500 mb-3">
        <p class="text-sm font-semibold text-slate-700 mb-2">{test_name}</p>
        <p class="text-xs font-mono text-slate-600 mb-2">
            Statistic = {stat_value:.4f} &nbsp;|&nbsp; p-value = {p_value:.4f} &nbsp;|&nbsp; α = {alpha}
        </p>
        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full
                      text-xs font-medium {badge_color} mb-2">
            {badge_icon} {badge_text}
        </span>
        <p class="text-sm text-slate-600 mt-1">{interpretation}</p>
    </div>
    """
    components.html(html, height=160)


def render_risk_badge(tier: str) -> str:
    """Return HTML for an inline risk tier badge.

    Args:
        tier: "High", "Medium", or "Low"

    Returns:
        HTML string for embedding in table or card
    """
    styles = {
        "High":   "bg-red-100 text-red-700",
        "Medium": "bg-amber-100 text-amber-700",
        "Low":    "bg-green-100 text-green-700",
    }
    css = styles.get(tier, styles["Low"])
    return (f'<span class="inline-flex px-2 py-0.5 rounded-full '
            f'text-xs font-semibold {css}">{tier}</span>')


def render_section_header(title: str, subtitle: str = ""):
    """Render a section header with accent underline."""
    sub_html = f'<p class="text-sm text-slate-500 mt-1">{subtitle}</p>' if subtitle else ""
    html = f"""
    <script src="https://cdn.tailwindcss.com"></script>
    <div class="mb-4">
        <h2 class="text-2xl font-bold text-slate-800
                    border-b-2 border-indigo-500 pb-2">{title}</h2>
        {sub_html}
    </div>
    """
    components.html(html, height=70 if subtitle else 55)

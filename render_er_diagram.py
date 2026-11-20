import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def draw_er_diagram():
    fig, ax = plt.subplots(figsize=(16, 11), dpi=300)
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 11)
    ax.axis('off')

    # Color Palette
    BG_COLOR = "#F8FAFC"
    DIM_HDR_COLOR = "#1B365D"    # Deep Navy for Dimensions
    FACT_HDR_COLOR = "#2B4C7E"   # Steel Navy for Facts
    BOX_BG = "#FFFFFF"
    BORDER_COLOR = "#94A3B8"
    TEXT_MAIN = "#0F172A"
    PK_COLOR = "#B45309"         # Amber/Brown for PK
    FK_COLOR = "#047857"         # Emerald Green for FK
    LINE_COLOR = "#334155"

    fig.patch.set_facecolor(BG_COLOR)

    # Title Banner
    ax.text(8.0, 10.5, "CUSTOMER RETENTION METRICS ANALYSIS — RELATIONAL ENTITY RELATIONSHIP DIAGRAM", 
            fontsize=13, weight='bold', color="#1B365D", ha='center', va='center', fontfamily='sans-serif')
    ax.text(8.0, 10.15, "3NF Star-Schema Architecture | SQLite3 (data/app.db) | Sumedha (ABES EC)", 
            fontsize=9.5, style='italic', color="#64748B", ha='center', va='center', fontfamily='sans-serif')

    def draw_entity(x, y, w, h, title, is_fact, attributes):
        hdr_fill = FACT_HDR_COLOR if is_fact else DIM_HDR_COLOR
        entity_type = "FACT TABLE" if is_fact else "MASTER DIMENSION"
        
        # Main box shadow & container
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.04,rounding_size=0.12",
                                      facecolor=BOX_BG, edgecolor=BORDER_COLOR, linewidth=1.2, zorder=2)
        ax.add_patch(rect)
        
        # Header banner
        hdr_h = 0.72
        hdr_rect = patches.FancyBboxPatch((x, y + h - hdr_h), w, hdr_h, 
                                          boxstyle="round,pad=0.04,rounding_size=0.12",
                                          facecolor=hdr_fill, edgecolor=hdr_fill, linewidth=1.0, zorder=3)
        ax.add_patch(hdr_rect)
        
        # Flat lower edge of header
        ax.add_patch(patches.Rectangle((x, y + h - hdr_h), w, 0.15, facecolor=hdr_fill, edgecolor=hdr_fill, zorder=3))

        # Header Text
        ax.text(x + w/2, y + h - 0.26, title, fontsize=9.5, weight='bold', color="#FFFFFF", ha='center', va='center', zorder=4)
        ax.text(x + w/2, y + h - 0.52, f"« {entity_type} »", fontsize=7.0, style='italic', color="#E2E8F0", ha='center', va='center', zorder=4)

        # Attribute lines
        line_y = y + h - hdr_h - 0.28
        for attr, key_type, dtype in attributes:
            # Key Tag
            if key_type == "PK":
                ax.text(x + 0.18, line_y, "PK", fontsize=7.5, weight='bold', color=PK_COLOR, va='center', zorder=4)
            elif key_type == "FK":
                ax.text(x + 0.18, line_y, "FK", fontsize=7.5, weight='bold', color=FK_COLOR, va='center', zorder=4)
            else:
                ax.text(x + 0.22, line_y, "•", fontsize=8.0, color="#94A3B8", va='center', zorder=4)
            
            # Attribute Name
            ax.text(x + 0.65, line_y, attr, fontsize=7.8, weight='bold' if key_type in ['PK', 'FK'] else 'normal', 
                    color=TEXT_MAIN, va='center', zorder=4)
            
            # Data Type
            ax.text(x + w - 0.18, line_y, dtype, fontsize=7.2, style='italic', color="#64748B", ha='right', va='center', zorder=4)
            
            line_y -= 0.29

    # Entity 1: dim_customers (Left-Center)
    cust_attrs = [
        ("customer_id", "PK", "TEXT"),
        ("company_name", "", "TEXT"),
        ("industry", "", "TEXT"),
        ("company_size", "", "TEXT"),
        ("region", "", "TEXT"),
        ("acquisition_channel", "", "TEXT"),
        ("signup_date", "", "DATE"),
        ("signup_cohort_month", "", "TEXT")
    ]
    draw_entity(0.8, 5.2, 3.8, 3.4, "dim_customers", False, cust_attrs)

    # Entity 2: dim_plans (Bottom-Left)
    plan_attrs = [
        ("plan_id", "PK", "TEXT"),
        ("plan_name", "", "TEXT (UQ)"),
        ("plan_tier_level", "", "INTEGER"),
        ("monthly_base_price", "", "REAL"),
        ("max_user_seats", "", "INTEGER"),
        ("sla_support_level", "", "TEXT")
    ]
    draw_entity(0.8, 0.8, 3.8, 2.9, "dim_plans", False, plan_attrs)

    # Entity 3: fact_subscriptions (Center)
    sub_attrs = [
        ("subscription_id", "PK", "TEXT"),
        ("customer_id", "FK", "TEXT"),
        ("plan_id", "FK", "TEXT"),
        ("start_date", "", "DATE"),
        ("end_date", "", "DATE"),
        ("billing_cadence", "", "TEXT"),
        ("monthly_fee_inr", "", "REAL"),
        ("subscription_status", "", "TEXT"),
        ("is_churned", "", "INTEGER"),
        ("churn_date", "", "DATE"),
        ("tenure_months", "", "INTEGER")
    ]
    draw_entity(6.1, 2.5, 4.0, 4.3, "fact_subscriptions", True, sub_attrs)

    # Entity 4: fact_support_tickets (Top-Right)
    ticket_attrs = [
        ("ticket_id", "PK", "TEXT"),
        ("customer_id", "FK", "TEXT"),
        ("created_at", "", "DATETIME"),
        ("resolved_at", "", "DATETIME"),
        ("resolution_hours", "", "REAL"),
        ("priority", "", "TEXT"),
        ("category", "", "TEXT"),
        ("ticket_csat", "", "INTEGER")
    ]
    draw_entity(11.4, 5.7, 3.8, 3.4, "fact_support_tickets", True, ticket_attrs)

    # Entity 5: fact_monthly_usage (Bottom-Right)
    usage_attrs = [
        ("usage_id", "PK", "TEXT"),
        ("customer_id", "FK", "TEXT"),
        ("activity_month", "", "TEXT"),
        ("login_days", "", "INTEGER"),
        ("core_feature_clicks", "", "INTEGER"),
        ("active_users_count", "", "INTEGER"),
        ("csat_score", "", "INTEGER"),
        ("nps_score", "", "INTEGER")
    ]
    draw_entity(11.4, 1.2, 3.8, 3.4, "fact_monthly_usage", True, usage_attrs)

    # Relationship Connectors with Cardinality & Verb Labels
    def draw_rel(p1, p2, verb, card_p1, card_p2, rad=0.0):
        con = patches.ConnectionPatch(
            xyA=p1, xyB=p2, coordsA="data", coordsB="data",
            arrowstyle="-|>", mutation_scale=12, linewidth=1.5,
            color=LINE_COLOR, connectionstyle=f"arc3,rad={rad}", zorder=1
        )
        ax.add_patch(con)
        
        # Midpoint for verb badge
        mid_x = (p1[0] + p2[0]) / 2 + (0.0 if rad == 0 else rad * 0.8)
        mid_y = (p1[1] + p2[1]) / 2 + (0.2 if rad == 0 else -rad * 0.4)
        
        ax.text(mid_x, mid_y, f" {verb} ", fontsize=7.2, weight='bold', color="#1E293B",
                bbox=dict(boxstyle="round,pad=0.2", fc="#FFFFFF", ec="#CBD5E1", lw=0.8),
                ha='center', va='center', zorder=5)

    # 1. dim_customers -> fact_subscriptions (REL-01)
    draw_rel((4.6, 6.5), (6.1, 5.5), "holds (1:N)", "1", "0..N", rad=0.08)

    # 2. dim_plans -> fact_subscriptions (REL-02)
    draw_rel((4.6, 2.4), (6.1, 3.2), "governs (1:N)", "1", "0..N", rad=-0.08)

    # 3. dim_customers -> fact_support_tickets (REL-03)
    draw_rel((4.6, 7.8), (11.4, 7.5), "logs (1:N)", "1", "0..N", rad=0.15)

    # 4. dim_customers -> fact_monthly_usage (REL-04)
    # Route via lower curve
    draw_rel((4.6, 5.4), (11.4, 2.9), "generates (1:N)", "1", "0..N", rad=-0.22)

    # Legend & Metadata Box
    leg_x, leg_y = 6.1, 0.8
    leg_box = patches.FancyBboxPatch((leg_x, leg_y), 4.0, 1.25, boxstyle="round,pad=0.03", 
                                     facecolor="#F1F5F9", edgecolor="#CBD5E1", linewidth=0.8)
    ax.add_patch(leg_box)
    ax.text(leg_x + 0.15, leg_y + 0.98, "DIAGRAM LEGEND & CONSTRAINTS:", fontsize=7.2, weight='bold', color="#334155")
    ax.text(leg_x + 0.15, leg_y + 0.68, "• PK: Primary Key (Unique, Non-Null)", fontsize=6.8, weight='bold', color=PK_COLOR)
    ax.text(leg_x + 2.1, leg_y + 0.68, "• FK: Foreign Key (Indexed)", fontsize=6.8, weight='bold', color=FK_COLOR)
    ax.text(leg_x + 0.15, leg_y + 0.38, "• Cardinality: 1:N (One-to-Many)", fontsize=6.8, color="#475569")
    ax.text(leg_x + 2.1, leg_y + 0.38, "• Referential Action: RESTRICT/CASCADE", fontsize=6.8, color="#475569")
    ax.text(leg_x + 0.15, leg_y + 0.12, "• SQLite Engine: PRAGMA foreign_keys = ON", fontsize=6.5, style='italic', color="#64748B")

    os.makedirs("docs", exist_ok=True)
    png_path = os.path.join("docs", "09_ER_Diagram.png")
    plt.tight_layout()
    plt.savefig(png_path, dpi=300, facecolor=BG_COLOR, bbox_inches='tight')
    plt.close()
    print(f"Successfully rendered {png_path}")

if __name__ == "__main__":
    draw_er_diagram()

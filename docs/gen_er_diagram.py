"""Script to generate docs/09_ER_Diagram.png using matplotlib with professional styling."""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, PathPatch
import matplotlib.path as mpath

PNG_PATH = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\09_ER_Diagram.png"

def draw_er_diagram():
    fig, ax = plt.subplots(figsize=(20, 14), dpi=300)
    ax.set_xlim(0, 20)
    ax.set_ylim(0, 14)
    ax.axis('off')
    
    # Background fill
    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    # Color Palette
    PRIMARY_NAVY = "#1E3A8A"    # Main titles, headers
    HEADER_BLUE  = "#1D4ED8"    # Entity Header
    HEADER_ACCENT= "#0284C7"    # Secondary headers
    ENTITY_BG    = "#FFFFFF"    # Entity Card BG
    BORDER_COLOR = "#94A3B8"    # Border lines
    PK_COLOR     = "#D97706"    # Gold / Amber for PK
    FK_COLOR     = "#2563EB"    # Blue for FK
    UQ_COLOR     = "#7C3AED"    # Purple for Unique
    TEXT_MAIN    = "#0F172A"    # Dark slate text
    TEXT_MUTED   = "#64748B"    # Muted grey text
    LINE_COLOR   = "#334155"    # Connector line color
    VIEW_BG      = "#EFF6FF"    # Analytical view card bg

    # Title Banner Block
    title_box = FancyBboxPatch((0.8, 12.4), 18.4, 1.2,
                               boxstyle="round,pad=0.2,rounding_size=0.15",
                               ec="#CBD5E1", fc="#FFFFFF", lw=1.5, zorder=2)
    ax.add_patch(title_box)
    
    ax.text(1.2, 13.15, "EMPLOYEE PERFORMANCE RATING ANALYSIS — ENTITY RELATIONSHIP DIAGRAM (3NF)",
            fontsize=15, fontweight='bold', color=PRIMARY_NAVY, va='center', ha='left', fontfamily='sans-serif')
    
    subtitle_text = "Standard Relational Schema Architecture (5 Normalized Entities + Joined Analytics View) | Crow's Foot Notation\n" \
                    "Student: Aayush Mathur  |  Enrolment No: 2500320140001  |  Section: MCA-B  |  ABES Engineering College"
    ax.text(1.2, 12.65, subtitle_text, fontsize=9.5, color="#475569", va='center', ha='left', fontfamily='sans-serif')

    # Helper function to draw an entity table card
    def draw_entity(x, y, w, h, table_name, subtitle, columns, header_color=HEADER_BLUE):
        # Drop shadow effect
        shadow = FancyBboxPatch((x + 0.08, y - 0.08), w, h,
                                boxstyle="round,pad=0.1,rounding_size=0.15",
                                ec="none", fc="#E2E8F0", zorder=1)
        ax.add_patch(shadow)

        # Main Card Body
        body = FancyBboxPatch((x, y), w, h,
                              boxstyle="round,pad=0.1,rounding_size=0.15",
                              ec=BORDER_COLOR, fc=ENTITY_BG, lw=1.5, zorder=2)
        ax.add_patch(body)

        # Header Box
        header_h = 0.8
        header = FancyBboxPatch((x, y + h - header_h), w, header_h,
                                boxstyle="round,pad=0.1,rounding_size=0.15",
                                ec=header_color, fc=header_color, lw=1.2, zorder=3)
        ax.add_patch(header)
        
        # Header separator straight rectangle to square off bottom of header
        header_base = patches.Rectangle((x, y + h - header_h), w, 0.2,
                                        fc=header_color, ec=header_color, zorder=3)
        ax.add_patch(header_base)

        # Header Text
        ax.text(x + w/2, y + h - 0.32, table_name.upper(),
                fontsize=11.5, fontweight='bold', color="#FFFFFF",
                va='center', ha='center', zorder=4, fontfamily='sans-serif')
        ax.text(x + w/2, y + h - 0.60, subtitle,
                fontsize=8, fontstyle='italic', color="#E0F2FE",
                va='center', ha='center', zorder=4, fontfamily='sans-serif')

        # Draw Columns
        start_y = y + h - header_h - 0.35
        row_height = (h - header_h - 0.5) / max(len(columns), 1)

        for idx, col in enumerate(columns):
            curr_y = start_y - (idx * row_height)
            
            # Subtle alternating row background
            if idx % 2 == 1:
                row_bg = patches.Rectangle((x + 0.05, curr_y - row_height*0.4), w - 0.1, row_height*0.85,
                                           fc="#F8FAFC", ec="none", zorder=2.5)
                ax.add_patch(row_bg)

            col_name = col[0]
            col_type = col[1]
            key_tag  = col[2] if len(col) > 2 else ""

            # Key tag badge
            if "PK" in key_tag:
                ax.text(x + 0.3, curr_y, "[PK]", fontsize=8, fontweight='bold',
                        color=PK_COLOR, va='center', ha='left', zorder=4, fontfamily='monospace')
            elif "FK" in key_tag:
                ax.text(x + 0.3, curr_y, "[FK]", fontsize=8, fontweight='bold',
                        color=FK_COLOR, va='center', ha='left', zorder=4, fontfamily='monospace')
            elif "UQ" in key_tag:
                ax.text(x + 0.3, curr_y, "[UQ]", fontsize=8, fontweight='bold',
                        color=UQ_COLOR, va='center', ha='left', zorder=4, fontfamily='monospace')
            else:
                ax.text(x + 0.35, curr_y, "•", fontsize=9, color="#94A3B8",
                        va='center', ha='left', zorder=4)

            # Column Name
            is_bold = "bold" if ("PK" in key_tag or "FK" in key_tag) else "normal"
            ax.text(x + 0.95, curr_y, col_name, fontsize=8.5, fontweight=is_bold,
                    color=TEXT_MAIN, va='center', ha='left', zorder=4, fontfamily='sans-serif')

            # Column Data Type
            ax.text(x + w - 0.25, curr_y, col_type, fontsize=7.5, fontstyle='italic',
                    color=TEXT_MUTED, va='center', ha='right', zorder=4, fontfamily='monospace')

    # ENTITY DEFINITIONS & COORDINATES
    # 1. departments (Top-Left)
    dept_cols = [
        ("dept_id", "INTEGER", "PK"),
        ("dept_name", "TEXT", "UQ"),
        ("cost_center", "TEXT", ""),
        ("department_head", "TEXT", ""),
        ("location", "TEXT", "")
    ]
    draw_entity(0.8, 7.8, 4.2, 3.8, "departments", "Organizational Units (1:M)", dept_cols, header_color="#0369A1")

    # 2. employees (Center Hub)
    emp_cols = [
        ("emp_id", "TEXT", "PK"),
        ("first_name", "TEXT", ""),
        ("last_name", "TEXT", ""),
        ("dept_id", "INTEGER", "FK"),
        ("manager_id", "TEXT", "FK"),
        ("job_role", "TEXT", ""),
        ("education_level", "TEXT", ""),
        ("gender", "TEXT", ""),
        ("age", "INTEGER", ""),
        ("monthly_income", "REAL", ""),
        ("hire_date", "TEXT", ""),
        ("marital_status", "TEXT", ""),
        ("tenure_years", "REAL", "")
    ]
    draw_entity(7.5, 4.4, 5.0, 7.2, "employees", "Core Workforce Master (Hub)", emp_cols, header_color=PRIMARY_NAVY)

    # 3. performance_reviews (Top-Right)
    review_cols = [
        ("review_id", "INTEGER", "PK"),
        ("emp_id", "TEXT", "FK"),
        ("review_year", "INTEGER", ""),
        ("review_cycle", "TEXT", ""),
        ("performance_rating", "REAL", ""),
        ("goals_completed_pct", "REAL", ""),
        ("projects_completed", "INTEGER", ""),
        ("manager_feedback_score", "REAL", ""),
        ("peer_review_score", "REAL", ""),
        ("years_since_last_prom", "REAL", ""),
        ("promotion_readiness", "REAL", "")
    ]
    draw_entity(14.8, 6.2, 4.4, 5.4, "performance_reviews", "Annual Appraisals (1:M)", review_cols, header_color="#1E40AF")

    # 4. training_records (Bottom-Right)
    training_cols = [
        ("training_id", "INTEGER", "PK"),
        ("emp_id", "TEXT", "FK"),
        ("training_course_name", "TEXT", ""),
        ("training_category", "TEXT", ""),
        ("training_hours", "REAL", ""),
        ("completion_date", "TEXT", ""),
        ("training_score", "REAL", ""),
        ("certification_earned", "INTEGER", "")
    ]
    draw_entity(14.8, 1.0, 4.4, 4.4, "training_records", "Skill Development (1:M)", training_cols, header_color="#047857")

    # 5. attrition_records (Bottom-Left)
    attrition_cols = [
        ("attrition_id", "INTEGER", "PK"),
        ("emp_id", "TEXT", "FK, UQ"),
        ("attrition_flag", "INTEGER", ""),
        ("attrition_date", "TEXT", ""),
        ("attrition_reason", "TEXT", ""),
        ("attrition_risk_score", "REAL", ""),
        ("work_life_balance", "INTEGER", ""),
        ("job_satisfaction", "INTEGER", "")
    ]
    draw_entity(0.8, 2.2, 4.4, 4.6, "attrition_records", "Turnover & Sentiment (1:1)", attrition_cols, header_color="#B91C1C")

    # Helper function to draw relationship connector lines with Crow's Foot style markers
    def draw_crow_foot_connector(p1, p2, label, card_p1="1", card_p2="M", style="orthogonal", waypoints=None):
        """Draws orthogonal connector line with crow's foot notation and business label."""
        if waypoints is None:
            points = [p1, p2]
        else:
            points = [p1] + waypoints + [p2]
            
        x_coords = [p[0] for p in points]
        y_coords = [p[1] for p in points]

        # Draw main relationship line
        ax.plot(x_coords, y_coords, color=LINE_COLOR, lw=2.2, zorder=5, solid_capstyle='round')

        # Add cardinality text near ends
        # Near p1
        dx1 = (points[1][0] - points[0][0])
        dy1 = (points[1][1] - points[0][1])
        d_len1 = (dx1**2 + dy1**2)**0.5
        norm_x1 = dx1 / max(d_len1, 0.001)
        norm_y1 = dy1 / max(d_len1, 0.001)
        card1_pos = (points[0][0] + norm_x1 * 0.45, points[0][1] + norm_y1 * 0.45 + 0.18)
        
        # Near p2
        dx2 = (points[-2][0] - points[-1][0])
        dy2 = (points[-2][1] - points[-1][1])
        d_len2 = (dx2**2 + dy2**2)**0.5
        norm_x2 = dx2 / max(d_len2, 0.001)
        norm_y2 = dy2 / max(d_len2, 0.001)
        card2_pos = (points[-1][0] + norm_x2 * 0.45, points[-1][1] + norm_y2 * 0.45 + 0.18)

        # Draw Crow's foot / one markers:
        # One marker: two parallel ticks perpendicular to line
        def draw_one_marker(pt, nx, ny):
            px, py = -ny, nx  # perpendicular vector
            for dist in [0.15, 0.28]:
                cx = pt[0] + nx * dist
                cy = pt[1] + ny * dist
                ax.plot([cx - px*0.14, cx + px*0.14], [cy - py*0.14, cy + py*0.14],
                        color="#047857", lw=2.5, zorder=6)

        # Many marker: Crow's foot (fork)
        def draw_many_marker(pt, nx, ny):
            px, py = -ny, nx
            # Base tick
            bx = pt[0] + nx * 0.32
            by = pt[1] + ny * 0.32
            ax.plot([bx - px*0.14, bx + px*0.14], [by - py*0.14, by + py*0.14],
                    color=PRIMARY_NAVY, lw=2, zorder=6)
            # Prong lines to tip
            tip_x = pt[0]
            tip_y = pt[1]
            ax.plot([bx - px*0.16, tip_x], [by - py*0.16, tip_y], color=PRIMARY_NAVY, lw=2.2, zorder=6)
            ax.plot([bx + px*0.16, tip_x], [by + px*0.16, tip_y], color=PRIMARY_NAVY, lw=2.2, zorder=6)
            ax.plot([bx, tip_x], [by, tip_y], color=PRIMARY_NAVY, lw=2.2, zorder=6)

        if card_p1 == "1":
            draw_one_marker(points[0], norm_x1, norm_y1)
        elif card_p1 == "M":
            draw_many_marker(points[0], norm_x1, norm_y1)

        if card_p2 == "1":
            draw_one_marker(points[-1], norm_x2, norm_y2)
        elif card_p2 == "M":
            draw_many_marker(points[-1], norm_x2, norm_y2)

        # Draw relationship label badge in middle of longest segment
        mid_idx = len(points) // 2
        lbl_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
        lbl_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
        
        lbl_box = FancyBboxPatch((lbl_x - 1.15, lbl_y - 0.22), 2.3, 0.44,
                                 boxstyle="round,pad=0.08,rounding_size=0.1",
                                 ec="#CBD5E1", fc="#FFFFFF", lw=1, zorder=7)
        ax.add_patch(lbl_box)
        ax.text(lbl_x, lbl_y, label, fontsize=8, fontweight='bold', color=PRIMARY_NAVY,
                va='center', ha='center', zorder=8, fontfamily='sans-serif')

    # CONNECTORS:
    # 1. departments (1) -> employees (M) [dept_id]
    # departments right edge: (5.0, 9.8) -> employees left edge: (7.5, 9.8)
    draw_crow_foot_connector((5.0, 9.8), (7.5, 9.8), "1:M  employs", card_p1="1", card_p2="M")

    # 2. employees (1) -> performance_reviews (M) [emp_id]
    # employees right edge: (12.5, 9.5) -> reviews left edge: (14.8, 9.5)
    draw_crow_foot_connector((12.5, 9.5), (14.8, 9.5), "1:M  evaluates", card_p1="1", card_p2="M")

    # 3. employees (1) -> training_records (M) [emp_id]
    # employees right edge: (12.5, 5.5) -> waypoints -> training_records left edge: (14.8, 3.5)
    draw_crow_foot_connector((12.5, 5.5), (14.8, 3.5), "1:M  completes", card_p1="1", card_p2="M",
                             waypoints=[(13.65, 5.5), (13.65, 3.5)])

    # 4. employees (1) -> attrition_records (1) [emp_id]
    # employees left edge: (7.5, 5.8) -> waypoints -> attrition right edge: (5.2, 4.6)
    draw_crow_foot_connector((7.5, 5.8), (5.2, 4.6), "1:1  records risk", card_p1="1", card_p2="1",
                             waypoints=[(6.35, 5.8), (6.35, 4.6)])

    # 5. employees (1) -> employees (M) [manager_id self-reference]
    # Curve on top of employees box from emp_id (PK) back into manager_id (FK)
    loop_pts = [(9.0, 11.6), (9.0, 12.2), (11.0, 12.2), (11.0, 11.6)]
    ax.plot([p[0] for p in loop_pts], [p[1] for p in loop_pts], color="#7C3AED", lw=2.2, zorder=5)
    
    # Self-ref label
    self_box = FancyBboxPatch((9.2, 12.0), 1.6, 0.4,
                              boxstyle="round,pad=0.06,rounding_size=0.1",
                              ec="#DDD6FE", fc="#F5F3FF", lw=1, zorder=7)
    ax.add_patch(self_box)
    ax.text(10.0, 12.2, "1:M  manages", fontsize=7.5, fontweight='bold', color="#6D28D9",
            va='center', ha='center', zorder=8)
    
    # Loop arrows / ticks
    # Tick at 9.0, 11.6 (Manager side - 1)
    ax.plot([8.85, 9.15], [11.75, 11.75], color="#047857", lw=2.5, zorder=6)
    ax.plot([8.85, 9.15], [11.85, 11.85], color="#047857", lw=2.5, zorder=6)
    # Crow foot at 11.0, 11.6 (Subordinate side - M)
    ax.plot([10.85, 11.15], [11.85, 11.85], color="#6D28D9", lw=2, zorder=6)
    ax.plot([10.85, 11.0], [11.85, 11.6], color="#6D28D9", lw=2, zorder=6)
    ax.plot([11.15, 11.0], [11.85, 11.6], color="#6D28D9", lw=2, zorder=6)

    # 6. Bottom Master Analytical View & Architecture Box (v_employee_analytics)
    view_box = FancyBboxPatch((6.0, 0.8), 8.0, 2.8,
                              boxstyle="round,pad=0.12,rounding_size=0.15",
                              ec="#93C5FD", fc=VIEW_BG, lw=1.5, ls="--", zorder=2)
    ax.add_patch(view_box)

    ax.text(10.0, 3.25, "v_employee_analytics  [MASTER ANALYTICAL VIEW]",
            fontsize=10.5, fontweight='bold', color=PRIMARY_NAVY, va='center', ha='center', zorder=4)
    
    view_desc = (
        "• Unified Denormalized Analytics Layer: Joins all 5 physical tables on PK-FK pathways\n"
        "• Comprehensive 38-Column Schema: Demographics, Longitudinal Ratings, Training, Retention\n"
        "• Zero-Redundancy Physical Model: Backed by strict Third Normal Form (3NF) relational tables\n"
        "• Direct Consumer: Power BI Dashboard & Streamlit App (Sub-second cached analytical queries)"
    )
    ax.text(6.3, 2.0, view_desc, fontsize=8.2, color="#1E293B", va='center', ha='left', zorder=4, linespacing=1.6)

    # Legend Block (Bottom Left / Bottom Right)
    legend_box = FancyBboxPatch((0.8, 0.8), 4.4, 1.1,
                                boxstyle="round,pad=0.08,rounding_size=0.1",
                                ec="#CBD5E1", fc="#FFFFFF", lw=1.2, zorder=2)
    ax.add_patch(legend_box)
    ax.text(1.0, 1.62, "SCHEMA NOTATION LEGEND", fontsize=8, fontweight='bold', color=PRIMARY_NAVY, va='center')
    ax.text(1.0, 1.32, "[PK] Primary Key (Unique Identifier)", fontsize=7.5, color=PK_COLOR, fontweight='bold', va='center')
    ax.text(1.0, 1.05, "[FK] Foreign Key (Referential Integrity)", fontsize=7.5, color=FK_COLOR, fontweight='bold', va='center')
    ax.text(3.1, 1.32, "[UQ] Unique Constraint", fontsize=7.5, color=UQ_COLOR, fontweight='bold', va='center')
    ax.text(3.1, 1.05, "|| : 1 (One)   <| : M (Many)", fontsize=7.5, color="#334155", fontweight='bold', va='center')

    plt.tight_layout()
    os.makedirs(os.path.dirname(PNG_PATH), exist_ok=True)
    plt.savefig(PNG_PATH, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none', bbox_inches='tight')
    plt.close()
    print(f"Successfully generated high-resolution ER Diagram image at: {PNG_PATH}")

if __name__ == "__main__":
    draw_er_diagram()

import re
from io import BytesIO
import pandas as pd
import plotly.express as px
import streamlit as st
from agent import ComplianceDecisionAgent
import tools

st.set_page_config(
    page_title="Nutrition Nest Screening Intelligence",
    layout="wide",
    page_icon="🛡️",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
        border-left: 5px solid #0d6efd;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    .metric-val { font-size: 26px; font-weight: 700; color: #212529; }
    .metric-lbl { font-size: 13px; color: #6c757d; text-transform: uppercase; }
</style>
""",
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------
# SIDEBAR CONFIGURATION
# ------------------------------------------------------------------
st.sidebar.image("https://img.icons8.com/fluency/96/shield.png", width=64)
st.sidebar.title("Procurement AI")
st.sidebar.caption("Sheet 21 Compliance & Decision System")

uploaded_file = st.sidebar.file_uploader(
    "Upload Categories Workbook (.xlsx)", type=["xlsx"]
)
default_file = "NutritionNest_Premium_Categories-copy.xlsx"
active_source = uploaded_file or default_file

st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ Screening Engine Settings")
min_rating_threshold = st.sidebar.slider(
    "Amazon Rating Gate (G-09)", 3.5, 4.8, 4.2, 0.1
)
only_populated = st.sidebar.checkbox(
    "Show Only Active Populated Categories (7 Sheets)",
    value=True,
    help="Hides the 13 empty template sheets from charts and filters.",
)
prefer_existing = st.sidebar.checkbox(
    "Preserve Pre-Filled Manual Notes (Where Present)",
    value=False,
    help="If checked, keeps existing notes in Functional Beverages/Coffee verbatim instead of re-evaluating.",
)
use_ollama_llm = st.sidebar.checkbox(
    "Route Chat Through Local Ollama (llama3.1)",
    value=False,
    help="Uses local Ollama llama3.1 to format agent responses if running.",
)


@st.cache_data(show_spinner=False)
def run_screening_pipeline(source, rating_val, keep_existing, file_mtime=None):
  return tools.tool_screen_and_writeback(
      source, min_rating=rating_val, prefer_existing_notes=keep_existing
  )


try:
  # Check modification time if reading from local disk
  mtime = (
      os.path.getmtime(active_source)
      if isinstance(active_source, str) and os.path.exists(active_source)
      else None
  )

  with st.spinner(
      "Running deep Sheet 21 rule engine & writing statuses to Excel..."
  ):
    master_df, category_dfs, sku_counts, updated_excel_bytes = (
        run_screening_pipeline(
            active_source,
            min_rating_threshold,
            prefer_existing,
            file_mtime=mtime,
        )
    )
except Exception as e:
  st.error(f"Error loading workbook: {e}")
  st.stop()

if only_populated:
    active_categories = [cat for cat, cnt in sku_counts.items() if cnt > 0]
    display_df = master_df[master_df["Category"].isin(active_categories)].copy()
else:
    active_categories = list(sku_counts.keys())
    display_df = master_df.copy()

agent = ComplianceDecisionAgent(display_df, use_ollama=use_ollama_llm)

# ------------------------------------------------------------------
# EXECUTIVE KPI BANNER
# ------------------------------------------------------------------
st.title("🛡️ Nutrition Nest — Shortlisting & Compliance Command Center")
st.caption(
    f"Full Sheet 21 Rule Engine (Sections 1–10) | Auditing **{len(display_df):,} SKUs** across **{len(active_categories)} Active Categories**"
)

total_count = len(display_df)
passed_count = int((display_df["Screening Status"] == "Passed").sum())
hold_count = int((display_df["Screening Status"] == "Hold (Proof Required)").sum())
rejected_count = int((display_df["Screening Status"] == "Rejected").sum())
pending_count = int((display_df["Screening Status"] == "Pending Data").sum())

m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.markdown(
        f"<div class='metric-card'><div class='metric-lbl'>Total SKUs</div><div class='metric-val'>{total_count:,}</div></div>",
        unsafe_allow_html=True,
    )
with m2:
    st.markdown(
        f"<div class='metric-card' style='border-left-color: #28a745;'><div class='metric-lbl'>✅ Passed (No Flags)</div><div class='metric-val' style='color:#28a745;'>{passed_count:,} <span style='font-size:14px;'>({(passed_count/total_count*100):.1f}%)</span></div></div>",
        unsafe_allow_html=True,
    )
with m3:
    st.markdown(
        f"<div class='metric-card' style='border-left-color: #ffc107;'><div class='metric-lbl'>⏳ Hold (COA / Proof)</div><div class='metric-val' style='color:#d39e00;'>{hold_count:,} <span style='font-size:14px;'>({(hold_count/total_count*100):.1f}%)</span></div></div>",
        unsafe_allow_html=True,
    )
with m4:
    st.markdown(
        f"<div class='metric-card' style='border-left-color: #dc3545;'><div class='metric-lbl'>🚫 Hard Rejected</div><div class='metric-val' style='color:#dc3545;'>{rejected_count:,} <span style='font-size:14px;'>({(rejected_count/total_count*100):.1f}%)</span></div></div>",
        unsafe_allow_html=True,
    )
with m5:
    st.markdown(
        f"<div class='metric-card' style='border-left-color: #6c757d;'><div class='metric-lbl'>⚠️ Pending Data (G-01)</div><div class='metric-val' style='color:#6c757d;'>{pending_count:,} <span style='font-size:14px;'>({(pending_count/total_count*100):.1f}%)</span></div></div>",
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)
st.download_button(
    label="📥 Download Audited Excel Workbook (With Section 10 Color-Coded Statuses Written to All Sheets)",
    data=updated_excel_bytes.getvalue(),
    file_name="NutritionNest_Screened_Audited.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
st.markdown("---")

# ------------------------------------------------------------------
# 8 INTERACTIVE TABS
# ------------------------------------------------------------------
(
    tab_viz,
    tab_agent,
    tab_triage,
    tab_outreach,
    tab_toxins,
    tab_commercial,
    tab_compare,
    tab_sheets,
) = st.tabs([
    "📊 Executive Analytics",
    "🤖 Analytical Decision Agent",
    "📋 Interactive Catalog Triage",
    "📑 Vendor & COA Outreach",
    "🧪 Chemical & Toxin Audit",
    "💰 Commercial & Sizing",
    "🥊 SKU Comparison",
    "📂 Sheet-by-Sheet Audit",
])

STATUS_COLOR_MAP = {
    "Passed": "#28a745",
    "Hold (Proof Required)": "#ffc107",
    "Rejected": "#dc3545",
    "Pending Data": "#6c757d",
}

TIER_COLOR_MAP = {
    "Red": "#ff7c80",
    "Purple": "#cdb4f6",
    "Orange": "#ffc000",
    "Blue": "#9dc3e6",
    "Pink": "#f4b6c2",
    "Grey": "#bfbfbf",
    "Yellow": "#ffff99",
    "Amber": "#ffe699",
    "None": "#c6efce",
}

# ------------------------------------------------------------------
# TAB 1: EXECUTIVE ANALYTICS
# ------------------------------------------------------------------
with tab_viz:
    c_v1, c_v2 = st.columns([3, 2])
    with c_v1:
        st.subheader("Category Compliance Breakdown")
        cat_stats = (
            display_df.groupby(["Category", "Screening Status"])
            .size()
            .reset_index(name="SKU Count")
        )
        fig_bar = px.bar(
            cat_stats,
            x="Category",
            y="SKU Count",
            color="Screening Status",
            color_discrete_map=STATUS_COLOR_MAP,
            barmode="stack",
            text_auto=True,
        )
        fig_bar.update_layout(xaxis_tickangle=-30, height=380)
        st.plotly_chart(fig_bar, use_container_width=True)

    with c_v2:
        st.subheader("Section 10 Priority Color Tiers (L-01 to L-09)")
        fig_donut = px.pie(
            display_df,
            names="Priority Color",
            color="Priority Color",
            color_discrete_map=TIER_COLOR_MAP,
            hole=0.5,
        )
        fig_donut.update_layout(height=380)
        st.plotly_chart(fig_donut, use_container_width=True)

    st.markdown("---")
    st.subheader("Hierarchical Drill-Down: Category ➔ Status ➔ Primary Rule Trigger")
    rej_copy = display_df.copy()

    def extract_first_rule(summary_str):
        if pd.isna(summary_str):
            return "None"
        s = str(summary_str)
        body = s.split(":", 1)[-1] if ":" in s else s
        return body.split(";")[0].strip()[:55]

    rej_copy["Primary Rule Trigger"] = rej_copy["Screening Summary"].apply(extract_first_rule)
    fig_sun = px.sunburst(
        rej_copy,
        path=["Category", "Screening Status", "Primary Rule Trigger"],
        color="Screening Status",
        color_discrete_map=STATUS_COLOR_MAP,
        height=520,
    )
    st.plotly_chart(fig_sun, use_container_width=True)

# ------------------------------------------------------------------
# TAB 2: ANALYTICAL DECISION AGENT
# ------------------------------------------------------------------
with tab_agent:
    st.subheader("💬 AI Procurement & Compliance Decision Agent")
    st.caption("Ask about specific SKUs (e.g., Mini Gnocchi, Stuffed Gnocchi), Section 10 color tiers, pass rates, or COA outreach.")

    q_cols = st.columns(4)
    preset_q = None
    if q_cols[0].button("🔍 Check Stuffed Gnocchi"):
        preset_q = "Why was Stuffed Gnocchi – Mushroom & Truffle rejected?"
    if q_cols[1].button("🔍 Check Mini Gnocchi"):
        preset_q = "Check status of Mini Gnocchi"
    if q_cols[2].button("🎨 Section 10 Color Tiers"):
        preset_q = "Show Section 10 color breakdown"
    if q_cols[3].button("🚫 Top Chemical Violations"):
        preset_q = "What are the top banned chemical violations?"

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = [
            {
                "role": "assistant",
                "content": (
                    "Welcome! I am running the full Sheet 21 (Sections 1–10) compliance engine. "
                    "Ask me to audit any SKU, compare category pass rates, inspect NOVA-4 or W-06 triggers, or draft a vendor COA email."
                ),
            }
        ]

    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_q = st.chat_input("Ask about any product, rule ID, or category...") or preset_q
    if user_q:
        st.session_state.chat_history.append({"role": "user", "content": user_q})
        with st.chat_message("user"):
            st.markdown(user_q)

        reply = agent.run(user_q)
        st.session_state.chat_history.append({"role": "assistant", "content": reply})
        with st.chat_message("assistant"):
            st.markdown(reply)

# ------------------------------------------------------------------
# TAB 3: INTERACTIVE CATALOG TRIAGE
# ------------------------------------------------------------------
with tab_triage:
    st.subheader("Curated Catalog Triage (With Section 10 Priority Colors)")
    f_c1, f_c2, f_c3, f_c4 = st.columns([1, 1, 1, 1.5])
    with f_c1:
        f_status = st.multiselect(
            "Filter Status:",
            options=display_df["Screening Status"].unique(),
            default=list(display_df["Screening Status"].unique()),
        )
    with f_c2:
        f_color = st.multiselect(
            "Filter Section 10 Color:",
            options=display_df["Priority Color"].unique(),
            default=list(display_df["Priority Color"].unique()),
        )
    with f_c3:
        f_cat = st.multiselect(
            "Filter Category:", options=active_categories, default=active_categories
        )
    with f_c4:
        search_kw = st.text_input("Search Product, Ingredient, or Rule ID (e.g. F-14, W-06):", "")

    filtered = display_df[
        display_df["Screening Status"].isin(f_status)
        & display_df["Priority Color"].isin(f_color)
        & display_df["Category"].isin(f_cat)
    ]
    if search_kw:
        filtered = filtered[
            filtered["Product name"].astype(str).str.contains(search_kw, case=False, na=False)
            | filtered.get("Manufacturer name", pd.Series(dtype=str)).astype(str).str.contains(search_kw, case=False, na=False)
            | filtered["Screening Summary"].astype(str).str.contains(search_kw, case=False, na=False)
            | filtered.get("Ingredients", pd.Series(dtype=str)).astype(str).str.contains(search_kw, case=False, na=False)
        ]

    display_cols = [
        c
        for c in [
            "Category",
            "Product name",
            "Manufacturer name",
            "Distributed in USA",
            "Product rating (Amazon)",
            "Screening Status",
            "Priority Color",
            "Screening Summary",
        ]
        if c in filtered.columns
    ]

    def style_priority_color(val):
        styles = {
            "Red": "background-color: #ff7c80; color: #721c24; font-weight: bold;",
            "Purple": "background-color: #cdb4f6; color: #3b1e6d; font-weight: bold;",
            "Orange": "background-color: #ffc000; color: #7f6000; font-weight: bold;",
            "Blue": "background-color: #9dc3e6; color: #1f3864; font-weight: bold;",
            "Pink": "background-color: #f4b6c2; color: #721c24; font-weight: bold;",
            "Grey": "background-color: #bfbfbf; color: #333333; font-weight: bold;",
            "Yellow": "background-color: #ffff99; color: #7f6000; font-weight: bold;",
            "Amber": "background-color: #ffe699; color: #856404; font-weight: bold;",
            "None": "background-color: #c6efce; color: #006100; font-weight: bold;",
        }
        return styles.get(val, "")

    subset_cols = ["Priority Color"] if "Priority Color" in display_cols else []
    styler = filtered[display_cols].style
    styled_df = (
        styler.map(style_priority_color, subset=subset_cols)
        if hasattr(styler, "map")
        else styler.applymap(style_priority_color, subset=subset_cols)
    )
    st.dataframe(styled_df, use_container_width=True, height=460)

# ------------------------------------------------------------------
# TAB 4: VENDOR & COA OUTREACH
# ------------------------------------------------------------------
with tab_outreach:
    st.subheader("📑 Supplier COA & Documentation Outreach Tracker")
    hold_items = display_df[
        display_df["Screening Status"].astype(str).str.contains("Hold", case=False, na=False)
    ].copy()

    if hold_items.empty:
        st.warning("No 'Hold (Proof Required)' items found. Press 'C' on your keyboard to clear Streamlit cache and reload.")
    else:
        o_c1, o_c2 = st.columns([1, 1])
        with o_c1:
            o_cat = st.selectbox(
                "Filter Hold Items by Category:",
                ["All Categories"] + list(hold_items["Category"].unique()),
                key="outreach_cat_sel"
            )
        with o_c2:
            mfg_opts = ["All Manufacturers"] + sorted(
                [str(x) for x in hold_items.get("Manufacturer name", pd.Series(dtype=str)).dropna().unique()]
            )
            o_mfg = st.selectbox("Filter by Manufacturer:", mfg_opts, key="outreach_mfg_sel")

        if o_cat != "All Categories":
            hold_items = hold_items[hold_items["Category"] == o_cat]

        if o_mfg != "All Manufacturers":
            hold_items = hold_items[hold_items["Manufacturer name"].astype(str) == o_mfg]

        out_cols = [
            c for c in [
                "Category", "Product name", "Manufacturer name",
                "Contact Email", "Contact number", "Screening Summary"
            ] if c in hold_items.columns
        ]
        st.dataframe(hold_items[out_cols], use_container_width=True, height=280)

        if not hold_items.empty:
            st.markdown("---")
            st.subheader("✉️ Generate Supplier Outreach Email")
            sel_idx = st.selectbox(
                "Select Product:",
                range(len(hold_items)),
                format_func=lambda i: f"{hold_items.iloc[i]['Product name']} ({hold_items.iloc[i].get('Manufacturer name', 'N/A')})",
                key="outreach_prod_sel"
            )
            to_addr, email_body = tools.tool_generate_coa_email(hold_items.iloc[sel_idx])
            st.text_area(f"Email to {to_addr}:", value=email_body, height=220, key="outreach_area")

# ------------------------------------------------------------------
# TAB 5: CHEMICAL & TOXIN AUDIT
# ------------------------------------------------------------------
with tab_toxins:
    st.subheader("🧪 Chemical Toxins & Hard-Gate Violation Analytics")
    rej_df = display_df[display_df["Screening Status"] == "Rejected"].copy()

    if rej_df.empty:
        st.info("No rejected products found in the current view.")
    else:
        # Robustly parse both '; ' and ' | ' delimiters
        all_rules = []
        for summary in rej_df["Screening Summary"].dropna():
            body = str(summary).split(":", 1)[-1]
            for item in re.split(r";|\|", body):
                item_clean = item.strip()
                if item_clean and not item_clean.upper().startswith("HOLD"):
                    all_rules.append(item_clean)

        tox_data = pd.Series(all_rules).value_counts().head(12).reset_index()
        tox_data.columns = ["Screening Rule / Toxin", "Violations Count"]

        t_c1, t_c2 = st.columns([1, 1])
        with t_c1:
            st.write("### Top Flagged Hard-Gate Rules")
            if not tox_data.empty:
                fig_tox = px.bar(
                    tox_data,
                    x="Violations Count",
                    y="Screening Rule / Toxin",
                    orientation="h",
                    color="Violations Count",
                    color_continuous_scale="Reds",
                )
                fig_tox.update_layout(yaxis={"categoryorder": "total ascending"}, height=380)
                st.plotly_chart(fig_tox, use_container_width=True)
        with t_c2:
            st.write("### Inspect Offending Products by Rule")
            if not tox_data.empty:
                chosen_tox = st.selectbox(
                    "Select Screening Rule:",
                    tox_data["Screening Rule / Toxin"].tolist(),
                    key="tox_rule_sel"
                )
                matched = display_df[
                    display_df["Screening Summary"]
                    .astype(str)
                    .str.contains(re.escape(chosen_tox), case=False, na=False)
                ]
                st.write(f"Showing **{len(matched)}** matching SKUs:")
                t_cols = [
                    c for c in ["Category", "Product name", "Manufacturer name", "Ingredients", "Screening Summary"]
                    if c in matched.columns
                ]
                st.dataframe(matched[t_cols], height=320, use_container_width=True)

# ------------------------------------------------------------------
# TAB 6: COMMERCIAL & SIZING
# ------------------------------------------------------------------
with tab_commercial:
    st.subheader("💰 Commercial Order Sizing (Approved Clean Products)")
    comm_metrics = tools.tool_commercial_metrics(display_df)
    cm1, cm2, cm3 = st.columns(3)
    cm1.metric("Approved Clean SKUs", f"{comm_metrics['Total']:,}")
    cm2.metric(
        "Median Price/Unit",
        f"${comm_metrics['Median_Price']:.2f}" if comm_metrics["Median_Price"] else "N/A",
    )
    cm3.metric("Median Supplier MOQ", f"{comm_metrics['Median_MOQ']} units")

    clean_df = display_df[display_df["Screening Status"] == "Passed"].copy()
    st.markdown("---")
    if clean_df.empty:
        st.info("No 'Passed' SKUs available under current filter settings.")
    else:
        comm_cols = [
            c for c in [
                "Category", "Product name", "Manufacturer name",
                "Price (MRP)-Manuf", "Price/unit", "MOQ",
                "Preferred Size", "Manufacturer link"
            ] if c in clean_df.columns
        ]
        st.dataframe(clean_df[comm_cols], use_container_width=True, height=380)

# ------------------------------------------------------------------
# TAB 7: SKU COMPARISON
# ------------------------------------------------------------------
with tab_compare:
    st.subheader("🥊 Head-to-Head Product Comparison")
    all_names = display_df["Product name"].dropna().unique().tolist()
    default_compare = [
        p for p in ["Mini Gnocchi", "Stuffed Gnocchi – Mushroom & Truffle"] if p in all_names
    ]
    if not default_compare:
        default_compare = all_names[:2] if len(all_names) >= 2 else all_names

    selected_prods = st.multiselect(
        "Select Products to Compare:",
        options=all_names,
        default=default_compare,
        key="compare_multiselect"
    )

    if selected_prods:
        comp_sub = display_df[display_df["Product name"].isin(selected_prods)].drop_duplicates(
            subset=["Product name"]
        )
        cols_ui = st.columns(len(comp_sub))
        for i, (_, row) in enumerate(comp_sub.iterrows()):
            with cols_ui[i]:
                st.markdown(f"### **{row['Product name']}**")
                st.write(f"**Category:** {row['Category']}")
                st.write(f"**Brand:** {row.get('Manufacturer name', 'N/A')}")
                v_col = (
                    "green"
                    if row["Screening Status"] == "Passed"
                    else "orange"
                    if "Hold" in str(row["Screening Status"])
                    else "red"
                )
                st.markdown(
                    f"**Verdict:** :{v_col}[**{row['Screening Status']}**] (`{row.get('Priority Color', 'N/A')}`)"
                )
                st.info(f"**Audit Details:** {row['Screening Summary']}")
                st.write(
                    f"**Price / MOQ:** {row.get('Price/unit', row.get('Price (MRP)-Manuf', 'N/A'))} | MOQ: {row.get('MOQ', 'N/A')}"
                )
                with st.expander("Full Ingredients Label", expanded=True):
                    st.write(str(row.get("Ingredients", "N/A")))

# ------------------------------------------------------------------
# TAB 8: SHEET-BY-SHEET AUDIT
# ------------------------------------------------------------------
with tab_sheets:
    sheet_opts = active_categories if only_populated else list(sku_counts.keys())
    sel_sheet = st.selectbox(
        "Select Sheet to Audit:",
        options=sheet_opts,
        format_func=lambda s: f"{s} ({sku_counts.get(s, 0)} SKUs)"
        if sku_counts.get(s, 0) > 0
        else f"{s} (Empty Template)",
        key="sheet_audit_sel"
    )
    s_df = category_dfs.get(sel_sheet, pd.DataFrame())
    if s_df.empty:
        st.info(f"ℹ️ Category **'{sel_sheet}'** is currently an unpopulated template.")
    else:
        s_cols = [
            c for c in [
                "_excel_row", "Product name", "Manufacturer name",
                "Ingredients", "Screening Status", "Priority Color", "Screening Summary"
            ] if c in s_df.columns
        ]
        st.dataframe(s_df[s_cols], use_container_width=True, height=420)
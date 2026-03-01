"""
LA 2028 Olympic Games Strategic Playbook — Interactive Dashboard
SportsFanatics Consulting Agency | Milestone 10

Run with:
    streamlit run dashboard/app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="LA 2028 Olympic Playbook",
    page_icon="🏅",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Paths ──────────────────────────────────────────────────────────────────
BASE      = Path(__file__).parent.parent
PROCESSED = BASE / "data" / "processed"
TABLES    = BASE / "outputs" / "tables"

# ── Colors ─────────────────────────────────────────────────────────────────
C_BLUE   = "#0085C7"
C_YELLOW = "#F4C300"
C_GREEN  = "#009F6B"
C_RED    = "#DF0024"
C_BLACK  = "#000000"

CONT_COLORS = {
    "Europe":   C_BLUE,
    "Americas": C_RED,
    "Asia":     "#E8A000",
    "Africa":   C_GREEN,
    "Oceania":  C_BLACK,
    "Other":    "#AAAAAA",
}

MEDAL_COLORS = {"Gold": C_YELLOW, "Silver": "#C0C0C0", "Bronze": "#CD7F32"}

TEMPLATE = "plotly_white"

# ── Data loading (cached) ──────────────────────────────────────────────────
@st.cache_data
def load_data():
    noc_lookup   = pd.read_csv(PROCESSED / "noc_regions_clean.csv")
    lookup       = noc_lookup[["NOC", "region", "continent"]].drop_duplicates("NOC")
    forecast     = pd.read_csv(PROCESSED / "forecast_la2028_enriched.csv")
    alltime      = pd.read_csv(PROCESSED / "noc_alltime_medal_table_enriched.csv")
    modern_tbl   = pd.read_csv(PROCESSED / "noc_modern_medal_table_enriched.csv")
    cont_summary = pd.read_csv(PROCESSED / "continental_medal_summary.csv")
    cont_fc      = pd.read_csv(PROCESSED / "continental_forecast_2028.csv")
    venues       = pd.read_csv(TABLES / "city_la2028_venues.csv")
    home_adv     = pd.read_csv(TABLES / "athlete_homeAdvantage.csv")
    sport_growth = pd.read_csv(TABLES / "athlete_sportGrowth_1984_2016.csv")
    noc_summary  = pd.read_csv(TABLES / "athlete_NOC_summary.csv")
    gdp_medals   = pd.read_csv(TABLES / "noc_gdp_medals_2016.csv")
    # Pre-computed from athlete_events.csv (raw file excluded from repo)
    gender_df       = pd.read_csv(PROCESSED / "gender_by_year.csv")
    medals_noc_year = pd.read_csv(PROCESSED / "medals_by_noc_year.csv")

    return {
        "lookup":           lookup,
        "forecast":         forecast,
        "alltime":          alltime,
        "modern_tbl":       modern_tbl,
        "cont_summary":     cont_summary,
        "cont_fc":          cont_fc,
        "venues":           venues,
        "home_adv":         home_adv,
        "sport_growth":     sport_growth,
        "noc_summary":      noc_summary,
        "gdp_medals":       gdp_medals,
        "gender_df":        gender_df,
        "medals_noc_year":  medals_noc_year,
    }


data            = load_data()
forecast        = data["forecast"]
alltime         = data["alltime"]
lookup          = data["lookup"]
gender_df       = data["gender_df"]
medals_noc_year = data["medals_noc_year"]

# ── Sidebar ────────────────────────────────────────────────────────────────
st.sidebar.image(
    "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a7/Olympic_flag.svg/200px-Olympic_flag.svg.png",
    width=120,
)
st.sidebar.title("LA 2028 Olympic\nStrategic Playbook")
st.sidebar.caption("SportsFanatics Consulting Agency")
st.sidebar.divider()

page = st.sidebar.radio(
    "Navigate",
    ["🏠 Overview", "🏃 Athlete Edge", "🏙️ City Playbook", "🌍 NOC Intelligence", "🔮 Medal Forecast"],
)

st.sidebar.divider()
st.sidebar.caption("Data: 120 Years of Olympic History  \n1896–2016 | 271,116 records")

# ═══════════════════════════════════════════════════════════════════════════
# PAGE 1 — OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════
if page == "🏠 Overview":
    st.title("🏅 LA 2028 Olympic Games — Strategic Playbook")
    st.markdown(
        "**A data-driven intelligence report for athletes, city stakeholders, and National Olympic Committees.**  \n"
        "Prepared by SportsFanatics Consulting Agency | March 2026"
    )
    st.divider()

    # KPI row
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Athlete Records", "271,116", "1896–2016")
    col2.metric("Unique Athletes", "135,571", "230 NOCs")
    col3.metric("USA 2028 Forecast", "277 medals", "+13 vs Rio 2016")
    col4.metric("LA 2028 Budget", "~$6.9B", "Lowest since 1984")

    st.divider()

    # Three pillars
    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("🏃 Pillar 1 — Athlete Edge")
        st.markdown("""
- Women now **45%** of athletes (was 9% in 1952)
- Host advantage delivers **+22.8%** medal uplift
- 4 new sports: flag football, cricket, lacrosse, squash
- LA climate: perfect July/August conditions (27°C avg)
        """)
    with c2:
        st.subheader("🏙️ Pillar 2 — City Playbook")
        st.markdown("""
- LA 2028 cost: **~$6.9B** — most disciplined in decades
- **84%** of venues already exist (no new builds)
- LA 1984 returned a **$215M profit** — the template
- 15 venues across Greater Los Angeles
        """)
    with c3:
        st.subheader("🌍 Pillar 3 — NOC Intelligence")
        st.markdown("""
- Europe holds **59.2%** of all-time Olympic medals
- China has rapidly closed the gap with the USA
- Russia/Belarus absence → **~60–70 medals** redistributed
- Africa: 27 medal-winning NOCs with growing base
        """)

    st.divider()

    # Quick continental chart
    st.subheader("Medal Share by Continent — Modern Era (1948–2016)")
    cont = data["cont_summary"].copy()
    cont = cont[cont["continent"] != "Other"]
    fig = px.pie(
        cont, values="Total_Medals", names="continent",
        color="continent", color_discrete_map=CONT_COLORS,
        hole=0.38, template=TEMPLATE,
    )
    fig.update_traces(textinfo="label+percent", pull=[0.03] * len(cont))
    fig.update_layout(showlegend=False, margin=dict(t=30, b=10, l=10, r=10), height=380)
    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# PAGE 2 — ATHLETE EDGE
# ═══════════════════════════════════════════════════════════════════════════
elif page == "🏃 Athlete Edge":
    st.title("🏃 Pillar 1 — The Athlete Edge")
    st.markdown("Explore 120 years of athlete participation, gender trends, and home advantage data.")
    st.divider()

    tab1, tab2, tab3 = st.tabs(["Gender Participation", "Home Advantage", "Sport Growth"])

    # ── TAB 1: Gender ──────────────────────────────────────────────────────
    with tab1:
        st.subheader("Gender Participation — Summer Olympics (1948–2016)")

        gender_df_plot = gender_df.copy()
        gender_df_plot["Sex"] = gender_df_plot["Sex"].map({"M": "Male", "F": "Female"})

        fig = px.area(
            gender_df_plot, x="Year", y="Athletes", color="Sex",
            color_discrete_map={"Male": C_BLUE, "Female": C_RED},
            template=TEMPLATE,
            labels={"Athletes": "Unique Athletes", "Year": "Olympic Year"},
        )
        fig.update_layout(hovermode="x unified", height=420, legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)

        # Pct female over time
        pivot = gender_df_plot.pivot(index="Year", columns="Sex", values="Athletes").fillna(0)
        pivot["Total"] = pivot.sum(axis=1)
        if "Female" in pivot.columns:
            pivot["Pct_Female"] = (pivot["Female"] / pivot["Total"] * 100).round(1)
            fig2 = px.line(
                pivot.reset_index(), x="Year", y="Pct_Female",
                template=TEMPLATE,
                labels={"Pct_Female": "Women as % of Athletes", "Year": "Olympic Year"},
                markers=True, color_discrete_sequence=[C_RED],
            )
            fig2.add_hline(y=50, line_dash="dash", line_color=C_GREY if False else "#888",
                           annotation_text="50% parity target")
            fig2.update_layout(height=320)
            st.plotly_chart(fig2, use_container_width=True)

        col1, col2, col3 = st.columns(3)
        col1.metric("Women in 1952", "9%", "of athletes")
        col2.metric("Women in 2016", "45%", "+36 percentage points")
        col3.metric("2028 Target", "50%", "Full parity")

    # ── TAB 2: Home Advantage ─────────────────────────────────────────────
    with tab2:
        st.subheader("Host Nation Advantage — Does Playing at Home Really Help?")

        home_df = data["home_adv"].copy()
        home_df["Group"] = home_df["Is_Host"].map({1: "Host Nation", 0: "All Others"})

        avg_rates = home_df.groupby("Group")["Medal_Rate"].mean().reset_index()
        delta = (
            avg_rates[avg_rates["Group"] == "Host Nation"]["Medal_Rate"].values[0]
            - avg_rates[avg_rates["Group"] == "All Others"]["Medal_Rate"].values[0]
        )

        col1, col2 = st.columns([1, 2])
        with col1:
            st.metric("Average Host Uplift", f"+{delta:.1f}%", "vs non-host games")
            st.metric("USA LA 1984 Golds", "83", "All-time host record")
            st.metric("2028 USA Forecast", "277 medals", "Host advantage priced in")
            st.markdown(
                "**Host nations win more medals because:**\n"
                "- No long-distance travel or jet lag\n"
                "- Familiar venues and training conditions\n"
                "- Home crowd energy boosts performance\n"
                "- More government funding in host year"
            )
        with col2:
            fig = px.line(
                home_df, x="Year", y="Medal_Rate", color="Group",
                color_discrete_map={"Host Nation": C_RED, "All Others": C_BLUE},
                template=TEMPLATE,
                labels={"Medal_Rate": "Medal Rate (%)", "Year": "Olympic Year"},
                markers=True,
            )
            fig.update_layout(hovermode="x unified", height=400, legend_title_text="")
            st.plotly_chart(fig, use_container_width=True)

    # ── TAB 3: Sport Growth ───────────────────────────────────────────────
    with tab3:
        st.subheader("Sport Growth — Which Sports Grew Most from 1984 to 2016?")

        sg = data["sport_growth"].dropna().copy()
        sg = sg.sort_values("Growth_pct", ascending=False).head(15)

        fig = px.bar(
            sg, x="Growth_pct", y="Sport",
            orientation="h",
            color="Growth_pct",
            color_continuous_scale=[[0, C_BLUE], [1, C_RED]],
            template=TEMPLATE,
            labels={"Growth_pct": "Growth (%)", "Sport": ""},
            text="Growth_pct",
        )
        fig.update_traces(texttemplate="%{text:.0f}%", textposition="outside")
        fig.update_layout(
            yaxis={"categoryorder": "total ascending"},
            coloraxis_showscale=False,
            height=500,
        )
        st.plotly_chart(fig, use_container_width=True)

        st.info("**New sports at LA 2028:** Flag Football · Cricket · Lacrosse · Squash — all added specifically for the Los Angeles audience.")


# ═══════════════════════════════════════════════════════════════════════════
# PAGE 3 — CITY PLAYBOOK
# ═══════════════════════════════════════════════════════════════════════════
elif page == "🏙️ City Playbook":
    st.title("🏙️ Pillar 2 — The City Playbook")
    st.markdown("Economic benchmarking, venue readiness, and legacy planning for LA 2028.")
    st.divider()

    tab1, tab2 = st.tabs(["Cost Benchmarking", "LA 2028 Venues"])

    # ── TAB 1: Costs ──────────────────────────────────────────────────────
    with tab1:
        st.subheader("Host City Cost Comparison (USD Billion)")

        # Hardcoded benchmark (authoritative figures from report)
        cost_data = pd.DataFrame({
            "City":  ["LA 1984", "Barcelona 1992", "Atlanta 1996", "Sydney 2000",
                      "Athens 2004", "Beijing 2008", "London 2012",
                      "Rio 2016", "Tokyo 2020", "LA 2028 (Est.)"],
            "Cost_B": [0.5, 9.4, 1.7, 4.9, 9.0, 42.0, 14.8, 13.7, 15.4, 6.9],
            "Type":  ["LA", "Other", "Other", "Other",
                      "Other", "Other", "Other",
                      "Other", "Other", "LA"],
        })

        fig = px.bar(
            cost_data, x="City", y="Cost_B",
            color="Type",
            color_discrete_map={"LA": C_RED, "Other": C_BLUE},
            template=TEMPLATE,
            labels={"Cost_B": "Cost (USD Billion)", "City": ""},
            text="Cost_B",
        )
        fig.update_traces(texttemplate="$%{text}B", textposition="outside")
        fig.update_layout(
            showlegend=False,
            height=430,
            yaxis_title="Cost (USD Billion)",
        )
        fig.add_annotation(
            x="LA 2028 (Est.)", y=6.9 + 2,
            text="★ LA 2028", showarrow=False,
            font=dict(color=C_RED, size=12, family="Arial"),
        )
        st.plotly_chart(fig, use_container_width=True)

        col1, col2, col3 = st.columns(3)
        col1.metric("LA 2028 Budget", "~$6.9B", "Lowest since 1984")
        col2.metric("Existing Venues", "84%", "No new construction")
        col3.metric("LA 1984 Profit", "$215M", "The template to repeat")

        st.markdown("""
**Why is LA 2028 so affordable?**

Most Olympic host cities spend enormous sums building new stadiums and infrastructure.
LA already has world-class venues — SoFi Stadium, Crypto.com Arena, the Rose Bowl —
all operational and already hosting major events. The city essentially only needs to
add the Olympic overlay (signage, media facilities, security) rather than building from scratch.
        """)

    # ── TAB 2: Venues ─────────────────────────────────────────────────────
    with tab2:
        st.subheader("LA 2028 Venue Map")

        venues_df = data["venues"].copy()

        # Filter controls
        clusters = ["All"] + sorted(venues_df["Cluster"].unique().tolist())
        sel_cluster = st.selectbox("Filter by geographic cluster", clusters)
        if sel_cluster != "All":
            venues_df = venues_df[venues_df["Cluster"] == sel_cluster]

        fig = px.scatter_mapbox(
            venues_df,
            lat="Lat", lon="Lon",
            hover_name="Venue",
            hover_data={"Sport": True, "Capacity": True, "Status": True,
                        "Lat": False, "Lon": False},
            color="Cluster",
            size="Capacity",
            size_max=25,
            zoom=9,
            center={"lat": 34.05, "lon": -118.25},
            mapbox_style="carto-positron",
            template=TEMPLATE,
            height=520,
        )
        fig.update_layout(margin=dict(l=0, r=0, t=10, b=0), legend_title_text="Cluster")
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(
            venues_df[["Venue", "Cluster", "Sport", "Capacity", "Status"]].reset_index(drop=True),
            use_container_width=True,
            hide_index=True,
        )


# ═══════════════════════════════════════════════════════════════════════════
# PAGE 4 — NOC INTELLIGENCE
# ═══════════════════════════════════════════════════════════════════════════
elif page == "🌍 NOC Intelligence":
    st.title("🌍 Pillar 3 — NOC Intelligence Report")
    st.markdown("Explore the global medal landscape — rankings, geography, rivalry, and emerging nations.")
    st.divider()

    tab1, tab2, tab3 = st.tabs(["Medal Table", "Country Explorer", "GDP vs Medals"])

    # ── TAB 1: Medal Table ────────────────────────────────────────────────
    with tab1:
        st.subheader("All-Time Olympic Medal Table (Summer Games)")

        # Filters
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            cont_opts = ["All"] + sorted(alltime["continent"].dropna().unique().tolist())
            sel_cont = st.selectbox("Filter by continent", cont_opts)
        with col_f2:
            top_n = st.slider("Show top N countries", 5, 50, 20)

        disp = alltime.copy()
        disp["Country"] = disp["region"].fillna(disp["NOC"])
        if sel_cont != "All":
            disp = disp[disp["continent"] == sel_cont]
        disp = disp.sort_values("Total", ascending=False).head(top_n)

        fig = px.bar(
            disp, x="Total", y="Country",
            orientation="h",
            color="continent",
            color_discrete_map=CONT_COLORS,
            template=TEMPLATE,
            labels={"Total": "Total Medals", "Country": ""},
        )
        fig.update_layout(
            yaxis={"categoryorder": "total ascending"},
            height=max(350, top_n * 22),
            legend_title_text="Continent",
        )
        st.plotly_chart(fig, use_container_width=True)

        # Stacked by medal type
        st.subheader("Gold / Silver / Bronze Breakdown")
        melt = disp[["Country", "Gold", "Silver", "Bronze"]].melt(
            id_vars="Country", var_name="Medal", value_name="Count"
        )
        fig2 = px.bar(
            melt, x="Count", y="Country", color="Medal",
            color_discrete_map=MEDAL_COLORS,
            orientation="h",
            template=TEMPLATE,
            category_orders={"Medal": ["Gold", "Silver", "Bronze"]},
            labels={"Count": "Medals", "Country": ""},
        )
        fig2.update_layout(
            yaxis={"categoryorder": "total ascending"},
            height=max(350, top_n * 22),
            legend_title_text="",
            barmode="stack",
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── TAB 2: Country Explorer ───────────────────────────────────────────
    with tab2:
        st.subheader("Country Medal History Explorer")

        all_regions = sorted(lookup["region"].dropna().unique().tolist())
        sel_countries = st.multiselect(
            "Select countries to compare (up to 6)",
            all_regions,
            default=["USA", "UK", "China", "Germany"],
            max_selections=6,
        )

        if not sel_countries:
            st.warning("Please select at least one country.")
        else:
            sel_nocs = lookup[lookup["region"].isin(sel_countries)]["NOC"].tolist()

            medals_ts = medals_noc_year[medals_noc_year["NOC"].isin(sel_nocs)].copy()

            fig = px.line(
                medals_ts, x="Year", y="Medals", color="region",
                markers=True, template=TEMPLATE,
                labels={"region": "Country", "Medals": "Medals Won", "Year": "Olympic Year"},
            )
            fig.update_layout(hovermode="x unified", height=430, legend_title_text="Country")
            st.plotly_chart(fig, use_container_width=True)

            # Summary stats
            summary = (
                medals_ts
                .groupby("region")
                .agg(
                    Total_Medals=("Medals", "sum"),
                    Games_Appeared=("Year", "nunique"),
                )
                .reset_index()
                .rename(columns={"region": "Country"})
                .sort_values("Total_Medals", ascending=False)
            )
            st.dataframe(summary, use_container_width=True, hide_index=True)

    # ── TAB 3: GDP vs Medals ──────────────────────────────────────────────
    with tab3:
        st.subheader("GDP vs Medal Count — Does Money Buy Olympic Success?")

        gdp = data["gdp_medals"].copy()
        gdp = gdp.merge(lookup, on="NOC", how="left")
        gdp["Country"] = gdp["region"].fillna(gdp["NOC"])
        gdp["GDP_USD_B"] = pd.to_numeric(gdp["GDP_USD_B"], errors="coerce")
        gdp = gdp.dropna(subset=["GDP_USD_B", "Medals_2016"])

        fig = px.scatter(
            gdp, x="GDP_USD_B", y="Medals_2016",
            color="continent",
            size="Medals_2016",
            size_max=35,
            hover_name="Country",
            color_discrete_map=CONT_COLORS,
            template=TEMPLATE,
            log_x=True,
            labels={
                "GDP_USD_B": "GDP (USD Billion, log scale)",
                "Medals_2016": "Medals at Rio 2016",
                "continent": "Continent",
            },
            trendline="ols",
        )
        fig.update_layout(height=480, legend_title_text="Continent")
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("""
**Key insight:** There is a positive relationship between GDP and medal count — richer countries
generally win more medals. However, there are notable outliers:
- **Kenya, Ethiopia, Jamaica** win many medals despite modest GDPs (specialisation pays off)
- **Saudi Arabia, UAE** win very few medals despite very high GDPs (late investment in sport)

**Implication:** It's not just money — *where* you invest and *which sports* you target matters enormously.
        """)


# ═══════════════════════════════════════════════════════════════════════════
# PAGE 5 — MEDAL FORECAST
# ═══════════════════════════════════════════════════════════════════════════
elif page == "🔮 Medal Forecast":
    st.title("🔮 LA 2028 Medal Forecast")
    st.markdown(
        "Predictions from an ensemble machine learning model (Random Forest + Gradient Boosting).  \n"
        "Trained on Summer Olympics data 1984–2012. Validated with Leave-One-Games-Out cross-validation."
    )
    st.divider()

    tab1, tab2, tab3 = st.tabs(["Country Predictions", "Continental Outlook", "Model Performance"])

    # ── TAB 1: Country Predictions ────────────────────────────────────────
    with tab1:
        st.subheader("LA 2028 Predicted Medal Count — All NOCs")

        fc = forecast.copy()
        fc["Country"] = fc["region"].fillna(fc["NOC"])

        # Filters
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            cont_opts = ["All"] + sorted(fc["continent"].dropna().unique().tolist())
            sel_cont = st.selectbox("Filter by continent", cont_opts, key="fc_cont")
        with col_f2:
            top_n = st.slider("Show top N", 5, 50, 20, key="fc_n")
        with col_f3:
            show_host = st.checkbox("Highlight USA (host)", value=True)

        fc_filtered = fc.copy()
        if sel_cont != "All":
            fc_filtered = fc_filtered[fc_filtered["continent"] == sel_cont]
        fc_filtered = fc_filtered.sort_values("Pred_2028_Ensemble", ascending=False).head(top_n)

        bar_colors = [
            C_RED if (show_host and r == 1) else CONT_COLORS.get(c, "#888888")
            for r, c in zip(fc_filtered["Is_Host"], fc_filtered["continent"])
        ]

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=fc_filtered["Pred_2028_Ensemble"],
            y=fc_filtered["Country"],
            orientation="h",
            marker_color=bar_colors,
            error_x=dict(
                type="data",
                symmetric=False,
                array=(fc_filtered["Pred_2028_High"] - fc_filtered["Pred_2028_Ensemble"]).values,
                arrayminus=(fc_filtered["Pred_2028_Ensemble"] - fc_filtered["Pred_2028_Low"]).values,
                color="#555",
            ),
            text=fc_filtered["Pred_2028_Ensemble"].round(0).astype(int),
            textposition="outside",
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Predicted: %{x:.0f} medals<br>"
                "<extra></extra>"
            ),
        ))
        fig.update_layout(
            template=TEMPLATE,
            yaxis={"categoryorder": "total ascending"},
            xaxis_title="Predicted Medals",
            height=max(350, top_n * 22),
        )
        st.plotly_chart(fig, use_container_width=True)

        # Delta vs 2016
        st.subheader("Change vs Rio 2016")
        fc_delta = fc_filtered.copy()
        fc_delta["Delta"] = (fc_delta["Pred_2028_Ensemble"] - fc_delta["Medals_2016"]).round(1)
        fc_delta = fc_delta.sort_values("Delta", ascending=False)

        fig2 = px.bar(
            fc_delta, x="Delta", y="Country",
            orientation="h",
            color="Delta",
            color_continuous_scale=[[0, C_RED], [0.5, "#DDDDDD"], [1, C_GREEN]],
            template=TEMPLATE,
            labels={"Delta": "Change vs 2016", "Country": ""},
        )
        fig2.update_layout(
            yaxis={"categoryorder": "total ascending"},
            coloraxis_showscale=False,
            height=max(350, top_n * 22),
        )
        fig2.add_vline(x=0, line_color="#333", line_width=1)
        st.plotly_chart(fig2, use_container_width=True)

        # Data table
        with st.expander("View full data table"):
            display_fc = fc.sort_values("Rank_2028")[
                ["Rank_2028", "Country", "continent", "Medals_2016",
                 "Pred_2028_Ensemble", "Pred_2028_Low", "Pred_2028_High", "Delta_vs_2016"]
            ].rename(columns={
                "Rank_2028": "Rank",
                "Pred_2028_Ensemble": "Predicted",
                "Pred_2028_Low": "Low",
                "Pred_2028_High": "High",
                "Delta_vs_2016": "Delta",
                "continent": "Continent",
            })
            display_fc["Predicted"] = display_fc["Predicted"].round(0).astype(int)
            display_fc["Low"]       = display_fc["Low"].round(0).astype(int)
            display_fc["High"]      = display_fc["High"].round(0).astype(int)
            st.dataframe(display_fc, use_container_width=True, hide_index=True)

    # ── TAB 2: Continental Outlook ─────────────────────────────────────────
    with tab2:
        st.subheader("LA 2028 Predicted Medals by Continent")

        cf = data["cont_fc"].copy()
        cf = cf[cf["continent"] != "Other"].sort_values("Predicted_2028", ascending=False)

        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(
                cf, x="continent", y="Predicted_2028",
                color="continent",
                color_discrete_map=CONT_COLORS,
                template=TEMPLATE,
                text="Predicted_2028",
                labels={"Predicted_2028": "Predicted Medals", "continent": "Continent"},
                error_y=(cf["Pred_High"] - cf["Predicted_2028"]).values,
                error_y_minus=(cf["Predicted_2028"] - cf["Pred_Low"]).values,
            )
            fig.update_traces(texttemplate="%{text:.0f}", textposition="outside")
            fig.update_layout(
                showlegend=False,
                yaxis_title="Predicted Medals",
                height=400,
            )
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            fig2 = px.bar(
                cf, x="continent", y="Delta_vs_2016",
                color="Delta_vs_2016",
                color_continuous_scale=[[0, C_RED], [0.5, "#DDDDDD"], [1, C_GREEN]],
                template=TEMPLATE,
                text="Delta_vs_2016",
                labels={"Delta_vs_2016": "Change vs 2016", "continent": "Continent"},
            )
            fig2.update_traces(texttemplate="%{text:+.0f}", textposition="outside")
            fig2.update_layout(
                coloraxis_showscale=False,
                yaxis_title="Change vs Rio 2016",
                height=400,
            )
            fig2.add_hline(y=0, line_color="#333", line_width=1)
            st.plotly_chart(fig2, use_container_width=True)

        st.dataframe(
            cf[["continent", "Medals_2016", "Predicted_2028", "Pred_Low",
                "Pred_High", "Delta_vs_2016", "NOC_Count"]].rename(columns={
                "continent": "Continent",
                "Medals_2016": "Rio 2016",
                "Predicted_2028": "LA 2028 (Pred.)",
                "Pred_Low": "Low",
                "Pred_High": "High",
                "Delta_vs_2016": "Change",
                "NOC_Count": "# NOCs",
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.info(
            "**Americas** gets the biggest boost (+75 medals) driven by USA's home advantage.  \n"
            "**Europe** is predicted to slip slightly as Asian nations continue to grow."
        )

    # ── TAB 3: Model Performance ───────────────────────────────────────────
    with tab3:
        st.subheader("Model Validation — How Accurate Is the Forecast?")

        try:
            val = pd.read_csv(TABLES / "forecast_validation_results.csv")
            val = val.merge(lookup, on="NOC", how="left")
            val["Country"] = val["region"].fillna(val["NOC"])
            val["abs_error"] = val["error"].abs()

            col1, col2, col3 = st.columns(3)
            col1.metric("Mean Absolute Error", f"{val['abs_error'].mean():.2f}", "medals per NOC")
            col2.metric("Median Error", f"{val['error'].median():.2f}", "medals (bias check)")
            col3.metric("Countries Validated", f"{val['NOC'].nunique()}", "Leave-One-Games-Out")

            # Actual vs predicted scatter
            fig = px.scatter(
                val, x="medals", y="predicted",
                hover_name="Country",
                color="abs_error",
                color_continuous_scale=[[0, C_GREEN], [0.5, C_YELLOW], [1, C_RED]],
                template=TEMPLATE,
                labels={
                    "medals": "Actual Medals",
                    "predicted": "Predicted Medals",
                    "abs_error": "Absolute Error",
                },
                opacity=0.7,
            )
            max_val = max(val["medals"].max(), val["predicted"].max()) + 5
            fig.add_shape(
                type="line", x0=0, y0=0, x1=max_val, y1=max_val,
                line=dict(color="#333", dash="dash"),
            )
            fig.add_annotation(
                x=max_val * 0.7, y=max_val * 0.75,
                text="Perfect prediction line",
                showarrow=False, font=dict(size=10, color="#555"),
            )
            fig.update_layout(height=480)
            st.plotly_chart(fig, use_container_width=True)
            st.caption("Points close to the dashed line = accurate predictions. Colour = size of error.")

        except Exception as e:
            st.warning(f"Validation data not available: {e}")

        # Feature importance
        st.subheader("What Drives the Model? — Feature Importance")
        try:
            fi = pd.read_csv(TABLES / "forecast_featureImportance.csv")
            if "feature" in fi.columns and "importance" in fi.columns:
                fi = fi.sort_values("importance", ascending=True)
                fig = px.bar(
                    fi, x="importance", y="feature",
                    orientation="h",
                    color="importance",
                    color_continuous_scale=[[0, C_BLUE], [1, C_GREEN]],
                    template=TEMPLATE,
                    labels={"importance": "Importance Score", "feature": "Feature"},
                    text="importance",
                )
                fig.update_traces(texttemplate="%{text:.3f}", textposition="outside")
                fig.update_layout(
                    coloraxis_showscale=False,
                    yaxis={"categoryorder": "total ascending"},
                    height=400,
                )
                st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Feature importance data not available: {e}")

# ── Footer ─────────────────────────────────────────────────────────────────
st.sidebar.divider()
st.sidebar.markdown(
    "**LA 2028 Olympic Strategic Playbook**  \n"
    "SportsFanatics Consulting Agency  \n"
    "*March 2026 | Analyst: Shabeeb*"
)

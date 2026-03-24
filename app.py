import calendar
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Sutarčių ir sąskaitų registras", layout="wide")

LT_MONTHS = {
    "sausis": 1,
    "vasaris": 2,
    "kovas": 3,
    "balandis": 4,
    "gegužė": 5,
    "birželis": 6,
    "liepa": 7,
    "rugpjūtis": 8,
    "rugsėjis": 9,
    "spalis": 10,
    "lapkritis": 11,
    "gruodis": 12,
}

STATUS_VALUES = {"išrašyta", "neišrašyta"}
INDEFINITE_DATE = pd.Timestamp("2100-12-31")

REQUIRED_COLS = [
    "Klientas",
    "Sutarties Nr.",
    "Galioja nuo",
    "Galioja iki",
]


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_bool_taip(value) -> bool:
    text = normalize_text(value).lower()
    return text.startswith("taip")


@st.cache_data(show_spinner=False)
def read_excel(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    return {sheet: pd.read_excel(BytesIO(file_bytes), sheet_name=sheet) for sheet in xls.sheet_names}


def detect_month_columns(df: pd.DataFrame) -> list[str]:
    cols = []
    for col in df.columns:
        name = normalize_text(col).lower()
        if name in LT_MONTHS:
            cols.append(col)
    return cols


def assign_month_periods(month_cols: list[str], start_year: int) -> dict[str, pd.Period]:
    if not month_cols:
        return {}
    periods = {}
    prev_month_num = None
    year = start_year
    for col in month_cols:
        month_num = LT_MONTHS[normalize_text(col).lower()]
        if prev_month_num is not None and month_num < prev_month_num:
            year += 1
        periods[col] = pd.Period(f"{year}-{month_num:02d}", freq="M")
        prev_month_num = month_num
    return periods


def month_range_label(period: pd.Period) -> str:
    lt_name = list(LT_MONTHS.keys())[list(LT_MONTHS.values()).index(period.month)].capitalize()
    return f"{lt_name} {period.year}"


def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == object:
            out[col] = out[col].apply(normalize_text)

    for col in ["Galioja nuo", "Galioja iki", "Žiemos sezonas galioja iki", "Vasaros sezonas galioja iki"]:
        if col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce")

    if "Aktas išsiųstas" in out.columns:
        out["Aktas išsiųstas_bool"] = out["Aktas išsiųstas"].apply(to_bool_taip)
    else:
        out["Aktas išsiųstas_bool"] = False

    if "Ar turi būti aktas" in out.columns:
        out["Reikia_akto_bool"] = out["Ar turi būti aktas"].apply(to_bool_taip)
    else:
        out["Reikia_akto_bool"] = False

    if "Automatizuotas" in out.columns:
        out["Automatizuotas_bool"] = out["Automatizuotas"].apply(to_bool_taip)
    else:
        out["Automatizuotas_bool"] = False

    if "Galioja iki" in out.columns:
        out["Neterminuota"] = out["Galioja iki"].dt.normalize().eq(INDEFINITE_DATE)
    else:
        out["Neterminuota"] = False

    return out


def check_required_columns(df: pd.DataFrame) -> list[str]:
    return [col for col in REQUIRED_COLS if col not in df.columns]


def active_on_date(row: pd.Series, ref_date: pd.Timestamp) -> bool:
    start = row.get("Galioja nuo")
    end = row.get("Galioja iki")
    if pd.isna(start):
        return False
    if pd.isna(end):
        return start <= ref_date
    return start <= ref_date <= end


def overlaps_month(row: pd.Series, period: pd.Period) -> bool:
    start = row.get("Galioja nuo")
    end = row.get("Galioja iki")
    if pd.isna(start):
        return False
    month_start = pd.Timestamp(period.start_time.date())
    month_end = pd.Timestamp(period.end_time.date())
    if pd.isna(end):
        return start <= month_end
    return start <= month_end and end >= month_start


def summarize_month_status(df: pd.DataFrame, month_map: dict[str, pd.Period]) -> pd.DataFrame:
    rows = []
    for col, period in month_map.items():
        series = df[col].fillna("").astype(str).str.strip().str.lower()
        issued = int((series == "išrašyta").sum())
        not_issued = int((series == "neišrašyta").sum())
        total_marked = issued + not_issued
        rows.append(
            {
                "period": period,
                "Mėnuo": month_range_label(period),
                "Išrašyta": issued,
                "Neišrašyta": not_issued,
                "Pažymėta iš viso": total_marked,
                "Neišrašyta %": (not_issued / total_marked * 100) if total_marked else 0.0,
            }
        )
    return pd.DataFrame(rows).sort_values("period")


def workload_by_month(df: pd.DataFrame, periods: list[pd.Period]) -> pd.DataFrame:
    rows = []
    for period in periods:
        mask = df.apply(lambda r: overlaps_month(r, period), axis=1)
        active = df[mask]
        rows.append(
            {
                "period": period,
                "Mėnuo": month_range_label(period),
                "Aktyvios sutartys": int(len(active)),
                "Automatizuotos": int(active["Automatizuotas_bool"].sum()) if "Automatizuotas_bool" in active.columns else 0,
                "Rankinės / neautomatizuotos": int(len(active) - active["Automatizuotas_bool"].sum()) if "Automatizuotas_bool" in active.columns else int(len(active)),
                "Reikia akto": int(active["Reikia_akto_bool"].sum()) if "Reikia_akto_bool" in active.columns else 0,
            }
        )
    return pd.DataFrame(rows).sort_values("period")


def contract_flows(df: pd.DataFrame, periods: list[pd.Period]) -> pd.DataFrame:
    rows = []
    for period in periods:
        start_mask = df["Galioja nuo"].dt.to_period("M").eq(period) if "Galioja nuo" in df.columns else pd.Series(False, index=df.index)
        end_mask = (
            df["Galioja iki"].dt.to_period("M").eq(period)
            & ~df["Neterminuota"]
            if "Galioja iki" in df.columns else pd.Series(False, index=df.index)
        )
        rows.append(
            {
                "period": period,
                "Mėnuo": month_range_label(period),
                "Naujos sutartys": int(start_mask.sum()),
                "Pasibaigusios sutartys": int(end_mask.sum()),
                "Balansas": int(start_mask.sum() - end_mask.sum()),
            }
        )
    return pd.DataFrame(rows).sort_values("period")


def current_metrics(df: pd.DataFrame, ref_date: pd.Timestamp) -> dict:
    active_mask = df.apply(lambda r: active_on_date(r, ref_date), axis=1)
    active = df[active_mask]
    current_month = ref_date.to_period("M")
    end_this_month = (
        df["Galioja iki"].dt.to_period("M").eq(current_month) & ~df["Neterminuota"]
        if "Galioja iki" in df.columns else pd.Series(False, index=df.index)
    )
    return {
        "active_count": int(len(active)),
        "indefinite_count": int(df["Neterminuota"].sum()),
        "end_this_month": int(end_this_month.sum()),
        "auto_active": int(active["Automatizuotas_bool"].sum()) if "Automatizuotas_bool" in active.columns else 0,
        "manual_active": int(len(active) - active["Automatizuotas_bool"].sum()) if "Automatizuotas_bool" in active.columns else int(len(active)),
    }


def upcoming_season_alerts(df: pd.DataFrame, ref_date: pd.Timestamp, days_ahead: int) -> pd.DataFrame:
    frames = []
    for season_col, season_name in [
        ("Žiemos sezonas galioja iki", "Žiemos sezonas"),
        ("Vasaros sezonas galioja iki", "Vasaros sezonas"),
    ]:
        if season_col not in df.columns:
            continue
        temp = df[["Klientas", "Sutarties Nr.", season_col]].copy()
        temp = temp.rename(columns={season_col: "Data"})
        temp["Sezonas"] = season_name
        temp = temp.dropna(subset=["Data"])
        temp["Liko dienų"] = (temp["Data"].dt.normalize() - ref_date.normalize()).dt.days
        temp = temp[(temp["Liko dienų"] >= 0) & (temp["Liko dienų"] <= days_ahead)]
        frames.append(temp)
    if not frames:
        return pd.DataFrame(columns=["Klientas", "Sutarties Nr.", "Sezonas", "Data", "Liko dienų"])
    out = pd.concat(frames, ignore_index=True)
    return out.sort_values(["Data", "Klientas"])


def contracts_expiring_soon(df: pd.DataFrame, ref_date: pd.Timestamp, days_ahead: int) -> pd.DataFrame:
    if "Galioja iki" not in df.columns:
        return pd.DataFrame()
    temp = df[["Klientas", "Sutarties Nr.", "Galioja iki", "Automatizuotas"]].copy() if "Automatizuotas" in df.columns else df[["Klientas", "Sutarties Nr.", "Galioja iki"]].copy()
    temp = temp[~df["Neterminuota"]].copy()
    temp["Liko dienų"] = (temp["Galioja iki"].dt.normalize() - ref_date.normalize()).dt.days
    temp = temp[(temp["Liko dienų"] >= 0) & (temp["Liko dienų"] <= days_ahead)]
    return temp.sort_values(["Galioja iki", "Klientas"])


def unissued_clients(df: pd.DataFrame, month_col: str) -> pd.DataFrame:
    temp = df.copy()
    series = temp[month_col].fillna("").astype(str).str.strip().str.lower()
    temp = temp[series == "neišrašyta"].copy()
    wanted = [c for c in [
        "Klientas",
        "Sutarties Nr.",
        "Įmonė",
        "Pateikimo būdas",
        "Ar turi būti aktas",
        "Aktas išsiųstas",
        "Automatizuotas",
        "Galioja iki",
        "Pastabos",
    ] if c in temp.columns]
    if not wanted:
        return pd.DataFrame()
    return temp[wanted].sort_values(["Klientas", "Sutarties Nr."])


def build_download_excel(month_status_df, workload_df, flows_df, season_df, expiring_df, selected_unissued_df, output_name="analize.xlsx"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        month_status_df.drop(columns=["period"], errors="ignore").to_excel(writer, index=False, sheet_name="Neisrasytos")
        workload_df.drop(columns=["period"], errors="ignore").to_excel(writer, index=False, sheet_name="Apkrova")
        flows_df.drop(columns=["period"], errors="ignore").to_excel(writer, index=False, sheet_name="Sutarciu_srautas")
        season_df.to_excel(writer, index=False, sheet_name="Sezonai")
        expiring_df.to_excel(writer, index=False, sheet_name="Baigiasi_greitai")
        selected_unissued_df.to_excel(writer, index=False, sheet_name="Klientai_neisrasyta")
    output.seek(0)
    return output.getvalue()


st.title("📄 Sutarčių ir sąskaitų registras")
st.caption("Įkelk Excel failą ir matyk sezoniškumą, aktyvias sutartis, apkrovos pokytį, naujas / pasibaigusias sutartis ir neišrašytų sąskaitų vaizdą.")

with st.sidebar:
    st.header("Nustatymai")
    uploaded_file = st.file_uploader("Įkelk Excel failą", type=["xlsx"])
    reference_date = st.date_input("Ataskaitinė data", value=date.today())
    season_alert_days = st.slider("Perspėti apie sezonų pabaigą per kiek dienų", 7, 120, 45)
    contract_alert_days = st.slider("Perspėti apie sutarčių pabaigą per kiek dienų", 7, 180, 45)

if not uploaded_file:
    st.info("Pirmiausia įkelk savo Excel registrą.")
    st.stop()

file_bytes = uploaded_file.getvalue()
sheets = read_excel(file_bytes)

default_sheet = "Registras" if "Registras" in sheets else list(sheets.keys())[0]

with st.sidebar:
    selected_sheet = st.selectbox("Duomenų lapas", list(sheets.keys()), index=list(sheets.keys()).index(default_sheet))

raw_df = sheets[selected_sheet].copy()
df = clean_df(raw_df)

missing = check_required_columns(df)
if missing:
    st.error(f"Trūksta privalomų stulpelių: {', '.join(missing)}")
    st.stop()

month_cols = detect_month_columns(df)
if not month_cols:
    st.error("Neradau mėnesių stulpelių (pvz. Sausis, Vasaris, Kovas...).")
    st.stop()

first_month_name = normalize_text(month_cols[0]).lower()
default_start_year = 2024 if LT_MONTHS.get(first_month_name) in {9, 10, 11, 12} else date.today().year

with st.sidebar:
    start_year = st.number_input("Pirmo mėnesio stulpelio metai", min_value=2020, max_value=2100, value=default_start_year, step=1)

month_map = assign_month_periods(month_cols, int(start_year))
month_status_df = summarize_month_status(df, month_map)
periods = list(month_map.values())
workload_df = workload_by_month(df, periods)
flows_df = contract_flows(df, periods)

ref_ts = pd.Timestamp(reference_date)
metrics = current_metrics(df, ref_ts)
season_df = upcoming_season_alerts(df, ref_ts, season_alert_days)
expiring_df = contracts_expiring_soon(df, ref_ts, contract_alert_days)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Galiojančios sutartys šiandien", metrics["active_count"])
c2.metric("Neterminuotos sutartys", metrics["indefinite_count"])
c3.metric("Baigiasi šį mėnesį", metrics["end_this_month"])
c4.metric("Aktyvios automatizuotos", metrics["auto_active"])
c5.metric("Aktyvios rankinės", metrics["manual_active"])

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Neišrašytos sąskaitos",
    "Apkrova",
    "Sutarčių srautas",
    "Sezonai ir terminai",
    "Duomenys",
])

with tab1:
    st.subheader("Neišrašytų sąskaitų dinamika")
    fig_unissued = px.bar(
        month_status_df,
        x="Mėnuo",
        y=["Išrašyta", "Neišrašyta"],
        barmode="group",
    )
    st.plotly_chart(fig_unissued, use_container_width=True)

    fig_pct = px.line(month_status_df, x="Mėnuo", y="Neišrašyta %", markers=True)
    st.plotly_chart(fig_pct, use_container_width=True)

    selected_month_label = st.selectbox("Pasirink mėnesį klientų sąrašui", month_status_df["Mėnuo"].tolist(), index=max(len(month_status_df)-1, 0))
    selected_period = month_status_df.loc[month_status_df["Mėnuo"] == selected_month_label, "period"].iloc[0]
    selected_month_col = next(col for col, per in month_map.items() if per == selected_period)
    selected_unissued_df = unissued_clients(df, selected_month_col)

    left, right = st.columns([1, 2])
    with left:
        st.metric("Neišrašyta pasirinktą mėnesį", int((df[selected_month_col].fillna("").astype(str).str.strip().str.lower() == "neišrašyta").sum()))
    with right:
        st.write("**Klientai / sutartys su statusu „Neišrašyta“**")
        st.dataframe(selected_unissued_df, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Apkrovos pokytis kas mėnesį")
    fig_workload = px.line(workload_df, x="Mėnuo", y=["Aktyvios sutartys", "Rankinės / neautomatizuotos", "Reikia akto"], markers=True)
    st.plotly_chart(fig_workload, use_container_width=True)
    st.dataframe(workload_df.drop(columns=["period"]), use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Naujos ir pasibaigusios sutartys")
    fig_flow = px.bar(flows_df, x="Mėnuo", y=["Naujos sutartys", "Pasibaigusios sutartys"], barmode="group")
    st.plotly_chart(fig_flow, use_container_width=True)
    fig_balance = px.line(flows_df, x="Mėnuo", y="Balansas", markers=True)
    st.plotly_chart(fig_balance, use_container_width=True)
    st.dataframe(flows_df.drop(columns=["period"]), use_container_width=True, hide_index=True)

with tab4:
    st.subheader("Sezonų pabaiga ir sutarčių terminai")
    col_a, col_b = st.columns(2)
    with col_a:
        st.write(f"**Sezonai, kurie baigiasi per {season_alert_days} d.**")
        st.dataframe(season_df, use_container_width=True, hide_index=True)
    with col_b:
        st.write(f"**Sutartys, kurios baigiasi per {contract_alert_days} d.**")
        st.dataframe(expiring_df, use_container_width=True, hide_index=True)

with tab5:
    st.subheader("Žali duomenys")
    st.dataframe(df, use_container_width=True, hide_index=True)

report_bytes = build_download_excel(month_status_df, workload_df, flows_df, season_df, expiring_df, selected_unissued_df)

st.download_button(
    "⬇️ Atsisiųsti analizės Excel",
    data=report_bytes,
    file_name="sutarciu_saskaitu_analize.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

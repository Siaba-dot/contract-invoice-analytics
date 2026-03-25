import io
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Sutarčių ir sąskaitų registras", layout="wide")

LT_MONTHS = [
    "Sausis", "Vasaris", "Kovas", "Balandis", "Gegužė", "Birželis",
    "Liepa", "Rugpjūtis", "Rugsėjis", "Spalis", "Lapkritis", "Gruodis"
]
MONTH_TO_NUM = {m: i + 1 for i, m in enumerate(LT_MONTHS)}
INDEFINITE_DATE = pd.Timestamp("2100-12-31")


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_status(value):
    text = normalize_text(value).lower()
    replacements = {
        "į": "i", "š": "s", "ų": "u", "ū": "u", "ž": "z",
        "ė": "e", "ą": "a", "č": "c",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text.strip()


def status_bucket(value):
    norm = normalize_status(value)
    if norm == "israsyta":
        return "Išrašyta"
    if norm == "neisrasyta":
        return "Neišrašyta"
    return ""


def standardize_yes_no(value):
    text = normalize_text(value).lower()
    if text in {"taip", "taip.", "yes", "y"}:
        return "Taip"
    if text in {"ne", "ne.", "no", "n"}:
        return "Ne"
    return normalize_text(value)


def extract_base_month_name(column_name):
    raw = str(column_name).strip()
    if "." in raw:
        left, right = raw.rsplit(".", 1)
        if right.isdigit():
            raw = left.strip()
    return raw


def find_month_columns(df):
    month_columns = []
    for col in df.columns:
        base = extract_base_month_name(col)
        if base in LT_MONTHS:
            month_columns.append(col)
    return month_columns


def build_timeline(month_columns, start_year):
    timeline = []
    current_year = int(start_year)
    prev_month_num = None

    for col in month_columns:
        base = extract_base_month_name(col)
        month_num = MONTH_TO_NUM[base]

        if prev_month_num is not None and month_num < prev_month_num:
            current_year += 1

        timeline.append(
            {
                "column": col,
                "month_name": base,
                "month_num": month_num,
                "year": current_year,
                "label": f"{base} {current_year}",
                "period_start": pd.Timestamp(year=current_year, month=month_num, day=1),
            }
        )
        prev_month_num = month_num

    return timeline


@st.cache_data(show_spinner=False)
def read_excel(uploaded_file, sheet_name):
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def prepare_dataframe(df):
    out = df.copy()

    date_columns = [
        "Galioja nuo",
        "Galioja iki",
        "Žiemos sezonas galioja iki",
        "Vasaros sezonas galioja iki",
        "Data",
    ]
    for col in date_columns:
        if col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce")

    yes_no_columns = [
        "Ar turi būti aktas",
        "Aktas išsiųstas",
        "Automatizuotas",
        "MMA",
        "Parduodamos prekės",
        "Pasibaigusi sutartis (pildoma tik jeigu pasibaigė)",
    ]
    for col in yes_no_columns:
        if col in out.columns:
            out[col] = out[col].apply(standardize_yes_no)

    month_columns = find_month_columns(out)
    for col in month_columns:
        out[col] = out[col].apply(status_bucket)

    return out, month_columns


def client_column_name(df):
    for col in ["Klientas", "Pirkėjas", "Užsakovas", "Kliento pavadinimas"]:
        if col in df.columns:
            return col
    return None


def contract_column_name(df):
    for col in ["Sutarties Nr.", "Sutarties nr.", "Sutarties numeris", "Sutartis", "Sutarties Nr", "Sutarties nr"]:
        if col in df.columns:
            return col
    return None


def object_column_name(df):
    for col in ["Objektas", "Adresas", "Sutarties pavadinimas"]:
        if col in df.columns:
            return col
    return None


def _min_non_null(series):
    s = pd.to_datetime(series, errors="coerce").dropna()
    return s.min() if not s.empty else pd.NaT


def _max_non_null(series):
    s = pd.to_datetime(series, errors="coerce").dropna()
    return s.max() if not s.empty else pd.NaT


def _merge_status_series(series):
    vals = [x for x in series if x in ("Išrašyta", "Neišrašyta")]
    if not vals:
        return ""
    if "Neišrašyta" in vals:
        return "Neišrašyta"
    return "Išrašyta"


def aggregate_contracts(df, month_columns):
    contract_col = contract_column_name(df)
    if contract_col is None:
        return df.copy(), False, None

    work = df.copy()
    work[contract_col] = work[contract_col].astype(str).str.strip()
    work = work[work[contract_col] != ""].copy()

    agg_map = {}

    for col in work.columns:
        if col == contract_col:
            continue
        if col in month_columns:
            agg_map[col] = _merge_status_series
        elif col == "Galioja nuo":
            agg_map[col] = _min_non_null
        elif col in ["Galioja iki", "Žiemos sezonas galioja iki", "Vasaros sezonas galioja iki", "Data"]:
            agg_map[col] = _max_non_null
        elif col in [
            "Ar turi būti aktas", "Aktas išsiųstas", "Automatizuotas", "MMA", "Parduodamos prekės",
            "Pasibaigusi sutartis (pildoma tik jeigu pasibaigė)"
        ]:
            agg_map[col] = lambda s: "Taip" if "Taip" in set(s.dropna().astype(str)) else ("Ne" if "Ne" in set(s.dropna().astype(str)) else "")
        else:
            agg_map[col] = lambda s: next((x for x in s if normalize_text(x) != ""), "")

    out = work.groupby(contract_col, dropna=False, as_index=False).agg(agg_map)
    return out, True, contract_col


def is_contract_active_on(row, check_date):
    start = row.get("Galioja nuo", pd.NaT)
    end = row.get("Galioja iki", pd.NaT)

    if pd.notna(start) and check_date < start:
        return False
    if pd.notna(end) and end != INDEFINITE_DATE and check_date > end:
        return False
    return True


def is_contract_active_in_month(row, month_start, month_end):
    start = row.get("Galioja nuo", pd.NaT)
    end = row.get("Galioja iki", pd.NaT)

    if pd.notna(start) and start > month_end:
        return False
    if pd.notna(end) and end != INDEFINITE_DATE and end < month_start:
        return False
    return True



def filter_only_valid_contracts(df, report_date):
    if "Galioja iki" not in df.columns:
        return df

    return df[
        (df["Galioja iki"].isna()) |
        (df["Galioja iki"] == INDEFINITE_DATE) |
        (df["Galioja iki"] >= report_date)
    ].copy()


def current_active_contracts(df, report_date):
    mask = df.apply(lambda r: is_contract_active_on(r, report_date), axis=1)
    return df[mask].copy()


def contract_type_series(df):
    if "Galioja iki" not in df.columns:
        return pd.Series(["Nežinoma"] * len(df), index=df.index)
    return df["Galioja iki"].apply(
        lambda x: "Neterminuota" if pd.notna(x) and x == INDEFINITE_DATE else "Terminuota"
    )


def contracts_ending_this_month(df, report_date):
    if "Galioja iki" not in df.columns:
        return df.iloc[0:0].copy()

    month_start = pd.Timestamp(report_date.year, report_date.month, 1)
    month_end = month_start + pd.offsets.MonthEnd(0)

    mask = (
        df["Galioja iki"].notna()
        & (df["Galioja iki"] != INDEFINITE_DATE)
        & (df["Galioja iki"] >= month_start)
        & (df["Galioja iki"] <= month_end)
    )
    return df[mask].copy()


def summarize_workload_by_month(df, timeline):
    rows = []
    for item in timeline:
        month_end = item["period_start"] + pd.offsets.MonthEnd(0)
        active_count = int(df.apply(lambda r: is_contract_active_on(r, month_end), axis=1).sum())
        rows.append({"Mėnuo": item["label"], "Aktyvios sutartys": active_count})
    return pd.DataFrame(rows)


def summarize_new_and_ended(df, timeline):
    rows = []
    for item in timeline:
        month_start = item["period_start"]
        month_end = item["period_start"] + pd.offsets.MonthEnd(0)

        new_count = 0
        ended_count = 0

        if "Galioja nuo" in df.columns:
            new_count = int(((df["Galioja nuo"] >= month_start) & (df["Galioja nuo"] <= month_end)).sum())

        if "Galioja iki" in df.columns:
            ended_count = int((
                df["Galioja iki"].notna()
                & (df["Galioja iki"] != INDEFINITE_DATE)
                & (df["Galioja iki"] >= month_start)
                & (df["Galioja iki"] <= month_end)
            ).sum())

        rows.append(
            {
                "Mėnuo": item["label"],
                "Naujos sutartys": new_count,
                "Pasibaigusios sutartys": ended_count,
                "Balansas": new_count - ended_count,
            }
        )
    return pd.DataFrame(rows)


def summarize_unissued_active_only(df, timeline):
    rows = []
    for item in timeline:
        col = item["column"]
        month_start = item["period_start"]
        month_end = item["period_start"] + pd.offsets.MonthEnd(0)

        active_mask = df.apply(lambda r: is_contract_active_in_month(r, month_start, month_end), axis=1)
        month_df = df[active_mask].copy()

        issued = int((month_df[col] == "Išrašyta").sum())
        unissued = int((month_df[col] == "Neišrašyta").sum())
        total_filled = issued + unissued
        pct = round((unissued / total_filled) * 100, 2) if total_filled else 0.0

        rows.append(
            {
                "Mėnuo": item["label"],
                "Išrašyta": issued,
                "Neišrašyta": unissued,
                "% neišrašyta": pct,
                "Aktyvios sutartys mėnesyje": int(len(month_df)),
            }
        )
    return pd.DataFrame(rows)


def season_alerts(df, report_date):
    rows = []
    client_col = client_column_name(df)
    contract_col = contract_column_name(df)
    obj_col = object_column_name(df)

    for _, row in df.iterrows():
        client_val = row[client_col] if client_col else ""
        contract_val = row[contract_col] if contract_col else ""
        obj_val = row[obj_col] if obj_col else ""

        for season_col, season_name in [
            ("Žiemos sezonas galioja iki", "Žiemos sezonas"),
            ("Vasaros sezonas galioja iki", "Vasaros sezonas"),
        ]:
            if season_col not in df.columns:
                continue

            end_date = row.get(season_col, pd.NaT)
            if pd.isna(end_date):
                continue
            if end_date == INDEFINITE_DATE:
                continue
            if end_date < report_date:
                continue

            rows.append(
                {
                    "Klientas": client_val,
                    "Sutarties Nr.": contract_val,
                    "Objektas / adresas": obj_val,
                    "Sezonas": season_name,
                    "Galioja iki": end_date.date(),
                    "Liko dienų": int((end_date - report_date).days),
                }
            )

    if not rows:
        return pd.DataFrame(columns=["Klientas", "Sutarties Nr.", "Objektas / adresas", "Sezonas", "Galioja iki", "Liko dienų"])

    return pd.DataFrame(rows).sort_values(["Galioja iki", "Klientas", "Sutarties Nr."]).reset_index(drop=True)


def build_export_excel(frames_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, frame in frames_dict.items():
            frame.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output


st.title("📊 Sutarčių ir sąskaitų registras")

uploaded_file = st.file_uploader("Įkelk Excel failą", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Įkelk registrą ir programa parodys aktyvias sutartis, neišrašytas sąskaitas, sezonus ir apkrovą.")
    st.stop()

excel_file = pd.ExcelFile(uploaded_file)
sheet_names = excel_file.sheet_names

with st.sidebar:
    st.header("Nustatymai")
    selected_sheet = st.selectbox("Lapas", sheet_names)
    st.markdown("**Svarbu:** jei mėnesių stulpeliai prasideda ne nuo sausio, startinius metus įvesk taip, kaip prasideda pirmas mėnesio stulpelis.")
    start_year = st.number_input("Pirmo mėnesio stulpelio metai", min_value=2020, max_value=2100, value=2025, step=1)
    report_date = pd.Timestamp(st.date_input("Ataskaitos data", value=date.today()))
    use_unique_contracts = st.toggle("Skaičiuoti pagal unikalius sutarčių numerius", value=True)

df_raw = read_excel(uploaded_file, selected_sheet)
df, month_columns = prepare_dataframe(df_raw)

if not month_columns:
    st.error("Nepavyko rasti mėnesių stulpelių. Jie turi būti pavadinti Sausis, Vasaris, Kovas ir t. t.")
    st.stop()

if use_unique_contracts:
    analysis_df, aggregated, contract_id_col = aggregate_contracts(df, month_columns)
    analysis_df = filter_only_valid_contracts(analysis_df, report_date)
else:
    analysis_df, aggregated, contract_id_col = df.copy(), False, contract_column_name(df)
    analysis_df = filter_only_valid_contracts(analysis_df, report_date)
timeline = build_timeline(month_columns, start_year)

active_df = current_active_contracts(analysis_df, report_date)
ending_this_month_df = contracts_ending_this_month(analysis_df, report_date)
season_df = season_alerts(analysis_df, report_date)
workload_df = summarize_workload_by_month(analysis_df, timeline)
flow_df = summarize_new_and_ended(analysis_df, timeline)
unissued_df = summarize_unissued_active_only(analysis_df, timeline)

client_col = client_column_name(analysis_df)
obj_col = object_column_name(analysis_df)

contract_types = contract_type_series(active_df)
active_contracts_count = len(active_df)
indefinite_count = int((contract_types == "Neterminuota").sum())
terminated_this_month_count = len(ending_this_month_df)
current_unissued_count = int(unissued_df.iloc[-1]["Neišrašyta"]) if not unissued_df.empty else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Šiuo metu aktyvių sutarčių", active_contracts_count)
c2.metric("Neterminuotos", indefinite_count)
c3.metric("Baigiasi šį mėnesį", terminated_this_month_count)
c4.metric("Paskutinio mėnesio neišrašyta", current_unissued_count)

with st.expander("Techninis patikrinimas"):
    st.write("Rasti mėnesių stulpeliai iš Excel:", [str(c) for c in month_columns])
    st.write("Sugeneruota laiko ašis:", [x["label"] for x in timeline])
    st.write("Sutartys skaičiuojamos pagal unikalų numerį:", bool(use_unique_contracts))
    st.write("Sutarties numerio stulpelis:", contract_id_col if contract_id_col else "Nerastas")
    if aggregated:
        st.write("Eilučių prieš sujungimą:", len(df))
        st.write("Unikalių sutarčių po sujungimo:", len(analysis_df))

st.divider()

left, right = st.columns((1.2, 1))

with left:
    st.subheader("Aktyvių sutarčių apkrova pagal mėnesius")
    fig_workload = px.line(workload_df, x="Mėnuo", y="Aktyvios sutartys", markers=True)
    fig_workload.update_layout(xaxis_title="", yaxis_title="Sutarčių skaičius")
    st.plotly_chart(fig_workload, use_container_width=True)

with right:
    st.subheader("Naujos ir pasibaigusios sutartys")
    fig_flow = px.bar(flow_df, x="Mėnuo", y=["Naujos sutartys", "Pasibaigusios sutartys"], barmode="group")
    fig_flow.update_layout(xaxis_title="", yaxis_title="Sutarčių skaičius")
    st.plotly_chart(fig_flow, use_container_width=True)

st.divider()

left, right = st.columns((1.2, 1))

with left:
    st.subheader("Neišrašytos sąskaitos pagal mėnesius (tik galiojančioms sutartims)")
    fig_unissued = px.bar(unissued_df, x="Mėnuo", y="Neišrašyta")
    fig_unissued.update_layout(xaxis_title="", yaxis_title="Sutarčių su neišrašyta skaičius")
    st.plotly_chart(fig_unissued, use_container_width=True)

with right:
    st.subheader("% neišrašyta pagal mėnesius")
    fig_unissued_pct = px.line(unissued_df, x="Mėnuo", y="% neišrašyta", markers=True)
    fig_unissued_pct.update_layout(xaxis_title="", yaxis_title="%")
    st.plotly_chart(fig_unissued_pct, use_container_width=True)

st.divider()

selected_month_label = st.selectbox(
    "Pasirink mėnesį detaliai peržiūrai",
    options=[x["label"] for x in timeline],
    index=len(timeline) - 1
)
selected_item = next(x for x in timeline if x["label"] == selected_month_label)
selected_month_col = selected_item["column"]
selected_month_start = selected_item["period_start"]
selected_month_end = selected_item["period_start"] + pd.offsets.MonthEnd(0)

detail_df = analysis_df[
    analysis_df.apply(lambda r: is_contract_active_in_month(r, selected_month_start, selected_month_end), axis=1)
].copy()
detail_df = detail_df[detail_df[selected_month_col] == "Neišrašyta"].copy()

st.subheader(f"Neišrašytos sąskaitos / sutartys: {selected_month_label}")
if detail_df.empty:
    st.success("Šiam mėnesiui neišrašytų nerasta.")
else:
    show_cols = []
    for col in [client_col, contract_id_col, obj_col, "Galioja nuo", "Galioja iki", "Ar turi būti aktas", "Aktas išsiųstas", "Automatizuotas", selected_month_col]:
        if col and col in detail_df.columns and col not in show_cols:
            show_cols.append(col)
    st.dataframe(detail_df[show_cols], use_container_width=True)

    if client_col:
        top_clients = (
            detail_df.groupby(client_col)
            .size()
            .reset_index(name="Neišrašyta")
            .sort_values("Neišrašyta", ascending=False)
        )
        st.markdown("**Klientai, kuriems liko neišrašyta**")
        st.dataframe(top_clients, use_container_width=True)

st.divider()

left, right = st.columns((1.1, 1))

with left:
    st.subheader("Sezonų pabaigos (imama iš sezono stulpelių, 2100-12-31 ignoruojama)")
    if season_df.empty:
        st.info("Tinkamų sezonų pabaigų nuo ataskaitos datos nerasta.")
    else:
        st.dataframe(season_df, use_container_width=True)

with right:
    st.subheader("Sutartys, kurios baigiasi šį mėnesį")
    if ending_this_month_df.empty:
        st.info("Šį mėnesį nesibaigia jokia terminuota sutartis.")
    else:
        cols = []
        for col in [client_col, contract_id_col, obj_col, "Galioja nuo", "Galioja iki", "Automatizuotas"]:
            if col and col in ending_this_month_df.columns and col not in cols:
                cols.append(col)
        st.dataframe(ending_this_month_df[cols], use_container_width=True)

st.divider()

tab1, tab2, tab3, tab4 = st.tabs([
    "Aktyvios sutartys",
    "Neišrašyta pagal mėnesius",
    "Sutarčių judėjimas",
    "Visa lentelė",
])

with tab1:
    active_show_cols = []
    for col in [client_col, contract_id_col, obj_col, "Galioja nuo", "Galioja iki", "Automatizuotas", "Ar turi būti aktas", "Aktas išsiųstas"]:
        if col and col in active_df.columns and col not in active_show_cols:
            active_show_cols.append(col)
    st.dataframe(active_df[active_show_cols] if active_show_cols else active_df, use_container_width=True)

with tab2:
    st.dataframe(unissued_df, use_container_width=True)

with tab3:
    st.dataframe(flow_df, use_container_width=True)

with tab4:
    st.dataframe(analysis_df, use_container_width=True)

excel_bytes = build_export_excel({
    "Aktyvios sutartys": active_df,
    "Neisrasyta pagal menesius": unissued_df,
    "Sutarciu judejimas": flow_df,
    "Sezonai": season_df,
    "Baigiasi si menesi": ending_this_month_df,
})

st.download_button(
    label="⬇️ Atsisiųsti analizės Excel",
    data=excel_bytes,
    file_name="sutarciu_ir_saskaitu_analize_pilna.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

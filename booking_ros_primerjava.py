import io
import re
import csv
import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Booking.com vs ROS PMS", layout="wide")

# ---------------------------------------------------------------------------
# Pomožne funkcije za branje in čiščenje podatkov
# ---------------------------------------------------------------------------

def parse_ros_csv(file_bytes: bytes) -> pd.DataFrame:
    """Prebere ROS PMS CSV izvoz (';'-ločen, cp1250, lahko vsebuje
    večvrstična citirana polja in je na koncu lahko obrezan)."""
    text = None
    for enc in ["cp1250", "iso-8859-2", "cp1252", "utf-8"]:
        try:
            text = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = file_bytes.decode("cp1250", errors="replace")

    # poenoti konce vrstic, da csv modul pravilno obravnava citirana polja
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    reader = csv.reader(io.StringIO(text), delimiter=";", quotechar='"')
    rows = list(reader)
    if not rows:
        return pd.DataFrame()

    header = rows[0]
    ncols = len(header)
    data = []
    for r in rows[1:]:
        if not r or all((x is None or x == "") for x in r):
            continue
        if len(r) < ncols:
            r = r + [""] * (ncols - len(r))  # zadnja vrstica je lahko obrezana
        elif len(r) > ncols:
            r = r[:ncols]
        data.append(r)

    df = pd.DataFrame(data, columns=header)
    return df


def parse_sl_date(s):
    """Pretvori slovenski datumski zapis npr. '1. 07. 2026' v pd.Timestamp."""
    if s is None:
        return pd.NaT
    s = str(s).strip()
    if not s:
        return pd.NaT
    m = re.match(r"(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})", s)
    if not m:
        return pd.NaT
    d, mo, y = m.groups()
    try:
        return pd.Timestamp(year=int(y), month=int(mo), day=int(d))
    except ValueError:
        return pd.NaT


def parse_sl_number(s):
    """Pretvori slovenski format števila ('996,6') v float."""
    if s is None:
        return None
    s = str(s).strip().replace(" ", "")
    if s == "":
        return None
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def file_hash(uploaded_file) -> str:
    uploaded_file.seek(0)
    h = hashlib.md5(uploaded_file.read()).hexdigest()
    uploaded_file.seek(0)
    return h


# ---------------------------------------------------------------------------
# Obdelava podatkov
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_booking_df(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8")
    df.columns = [c.strip() for c in df.columns]
    df["Reservation number"] = df["Reservation number"].astype(str).str.strip()
    for col in ["Arrival", "Departure"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    for col in ["Original amount", "Final amount", "Commission amount", "Commission %"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df["Nights_BC"] = (df["Departure"] - df["Arrival"]).dt.days
    return df


@st.cache_data(show_spinner=False)
def load_ros_df(file_bytes: bytes) -> pd.DataFrame:
    raw = parse_ros_csv(file_bytes)
    if raw.empty:
        return raw

    raw["Zacetek_dt"] = raw["Začetek"].apply(parse_sl_date)
    raw["Konec_dt"] = raw["Konec"].apply(parse_sl_date)
    raw["Cena_num"] = raw["Cena"].apply(parse_sl_number)
    raw["Promet_num"] = raw["Promet"].apply(parse_sl_number)
    raw["Referenca"] = raw["Referenca"].astype(str).str.strip()

    group_cols = ["Rezervacija"]
    rows = []
    for rez, g in raw.groupby(group_cols, dropna=False):
        # Promet je včasih podvojen na več vrsticah iste rezervacije
        # (isti znesek, različne vrstice/dolocila) - v tem primeru ga ne
        # seštevamo, sicer pa seštejemo, ker gre za dejansko ločene postavke.
        unique_promet = pd.Series(g["Promet_num"].round(4)).dropna().unique()
        if len(unique_promet) == 0:
            total_promet = 0.0
        elif len(unique_promet) == 1:
            total_promet = float(unique_promet[0])
        else:
            total_promet = float(unique_promet.sum())

        cena_values = pd.Series(g["Cena_num"]).dropna().unique()
        cena_display = cena_values[0] if len(cena_values) == 1 else (
            ", ".join(f"{v:g}" for v in cena_values) if len(cena_values) else None
        )

        zacetek = g["Zacetek_dt"].min()
        konec = g["Konec_dt"].max()
        referenca = next((r for r in g["Referenca"] if r), "")
        naziv_skupine = next((v for v in g.get("Naziv skupine", []) if v), "")
        gost = next((v for v in g.get("IME_GOSTA", []) if v), "")
        prostor = next((v for v in g.get("Naziv tipa", []) if v), "")

        rows.append({
            "Rezervacija_ROS": rez,
            "Referenca": referenca,
            "Naziv skupine": naziv_skupine,
            "Gost_ROS": gost,
            "Tip prostora": prostor,
            "Zacetek_ROS": zacetek,
            "Konec_ROS": konec,
            "Nights_ROS": (konec - zacetek).days if pd.notna(zacetek) and pd.notna(konec) else None,
            "Cena_ROS": cena_display,
            "Promet_ROS": round(total_promet, 2),
            "St_vrstic": len(g),
        })

    return pd.DataFrame(rows)


def build_comparison(bc_df: pd.DataFrame, ros_df: pd.DataFrame, dodatki: dict, tolerance: float) -> pd.DataFrame:
    ros_valid = ros_df[ros_df["Referenca"] != ""].copy()

    merged = bc_df.merge(
        ros_valid,
        left_on="Reservation number",
        right_on="Referenca",
        how="left",
        suffixes=("", "_ros"),
    )

    merged["Dodatek_ROS"] = merged["Reservation number"].map(
        lambda r: dodatki.get(r, 0.0)
    ).fillna(0.0)

    merged["V_ROS_najdeno"] = merged["Rezervacija_ROS"].notna()

    merged["Skupaj_ROS"] = merged["Promet_ROS"].fillna(0.0) + merged["Dodatek_ROS"]

    merged["Razlika_znesek"] = merged["Skupaj_ROS"] - merged["Final amount"]
    merged["Znesek_OK"] = merged["V_ROS_najdeno"] & (merged["Razlika_znesek"].abs() <= tolerance)

    merged["Datumi_OK"] = (
        merged["V_ROS_najdeno"]
        & (merged["Arrival"] == merged["Zacetek_ROS"])
        & (merged["Departure"] == merged["Konec_ROS"])
    )

    merged["Provizija_pricakovana"] = (merged["Final amount"] * merged["Commission %"] / 100.0).round(2)
    merged["Provizija_razlika"] = merged["Provizija_pricakovana"] - merged["Commission amount"]
    merged["Provizija_OK"] = merged["Provizija_razlika"].abs() <= tolerance

    merged["Neto_znesek_Booking"] = merged["Final amount"] - merged["Commission amount"]

    def status(row):
        if not row["V_ROS_najdeno"]:
            return "❌ Ni v ROS"
        probs = []
        if not row["Datumi_OK"]:
            probs.append("datumi")
        if not row["Znesek_OK"]:
            probs.append("znesek")
        if not row["Provizija_OK"]:
            probs.append("provizija")
        if probs:
            return "⚠️ Ne štima: " + ", ".join(probs)
        return "✅ OK"

    merged["Status"] = merged.apply(status, axis=1)
    return merged


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.title("📊 Primerjava: Booking.com statement vs ROS PMS")
st.caption(
    "Naloži Booking.com 'reservation statements overview' (nosilna datoteka) in "
    "ROS PMS izvoz rezervacij. Aplikacija poveže rezervacije preko polja "
    "**Referenca** (ROS) = **Reservation number** (Booking.com), primerja datume, "
    "izračuna promet in preveri provizijo."
)

with st.sidebar:
    st.header("1. Nalaganje datotek")
    bc_file = st.file_uploader("Booking.com statement (reservation_statements_overview...)", type=["csv"])
    ros_file = st.file_uploader("ROS PMS izvoz rezervacij (SEZNAM_REZERVACIJ...)", type=["csv"])

    st.header("2. Nastavitve")
    tolerance = st.number_input(
        "Toleranca za ujemanje zneskov (EUR)",
        min_value=0.0, max_value=50.0, value=0.02, step=0.01,
        help="Znotraj te razlike se znesek/provizija šteje kot ujemajoč.",
    )
    show_only_mismatch = st.checkbox("Prikaži samo neujemajoče rezervacije", value=False)

if not bc_file or not ros_file:
    st.info("⬅️ V stranski vrstici naloži obe CSV datoteki, da se prikaže primerjava.")
    st.stop()

bc_bytes = bc_file.getvalue()
ros_bytes = ros_file.getvalue()

try:
    bc_df = load_booking_df(bc_bytes)
except Exception as e:
    st.error(f"Napaka pri branju Booking.com datoteke: {e}")
    st.stop()

try:
    ros_df = load_ros_df(ros_bytes)
except Exception as e:
    st.error(f"Napaka pri branju ROS PMS datoteke: {e}")
    st.stop()

if ros_df.empty:
    st.error("ROS PMS datoteka je prazna ali je ni bilo mogoče prebrati.")
    st.stop()

# ključ za session_state, da se dodatki ne pobrišejo ob vsakem ponovnem zagonu,
# a se resetirajo, če se naložita nove datoteke
data_key = f"dodatki_{file_hash(bc_file)}_{file_hash(ros_file)}"
if data_key not in st.session_state:
    st.session_state[data_key] = {}

dodatki = st.session_state[data_key]

comparison = build_comparison(bc_df, ros_df, dodatki, tolerance)

# ---------------------------------------------------------------------------
# Povzetek
# ---------------------------------------------------------------------------

total_bc = len(comparison)
found_in_ros = int(comparison["V_ROS_najdeno"].sum())
not_found = total_bc - found_in_ros
ok_count = int((comparison["Status"] == "✅ OK").sum())
mismatch_count = total_bc - ok_count - not_found

c1, c2, c3, c4 = st.columns(4)
c1.metric("Rezervacij v Booking.com", total_bc)
c2.metric("Najdenih v ROS", found_in_ros)
c3.metric("✅ Ujemajo se v celoti", ok_count)
c4.metric("⚠️ Ne štima / manjka", total_bc - ok_count)

st.divider()

# ---------------------------------------------------------------------------
# Urejevalna tabela - vnos dodatka na rezervacijo
# ---------------------------------------------------------------------------

st.subheader("💶 Dodatek na ROS znesek (poljubno polje)")
st.caption(
    "Če ima rezervacija v ROS dodatne postavke, ki ne spadajo v osnovni Promet "
    "(npr. turistična taksa, doplačila ipd.), jih vnesi tukaj po rezervaciji - "
    "znesek se bo prištel k Prometu iz ROS in ponovno preveril z Final amount."
)

editable_cols = ["Reservation number", "Guest name", "Promet_ROS"]
editable_df = comparison[editable_cols].copy()
editable_df["Dodatek_ROS"] = editable_df["Reservation number"].map(lambda r: dodatki.get(r, 0.0))
editable_df = editable_df.rename(columns={
    "Reservation number": "Booking rezervacija",
    "Guest name": "Gost",
    "Promet_ROS": "Promet (ROS, iz datoteke)",
})

edited = st.data_editor(
    editable_df,
    key=f"editor_{data_key}",
    hide_index=True,
    use_container_width=True,
    disabled=["Booking rezervacija", "Gost", "Promet (ROS, iz datoteke)"],
    column_config={
        "Dodatek_ROS": st.column_config.NumberColumn(
            "Dodatek ROS (EUR)", help="Poljuben dodaten znesek za to rezervacijo", step=0.01, format="%.2f"
        ),
    },
)

# shrani spremembe nazaj v session_state in ponovno izračunaj primerjavo
changed = False
for _, r in edited.iterrows():
    key = r["Booking rezervacija"]
    val = float(r["Dodatek_ROS"]) if pd.notna(r["Dodatek_ROS"]) else 0.0
    if dodatki.get(key, 0.0) != val:
        dodatki[key] = val
        changed = True

if changed:
    st.session_state[data_key] = dodatki
    comparison = build_comparison(bc_df, ros_df, dodatki, tolerance)
    ok_count = int((comparison["Status"] == "✅ OK").sum())

st.divider()

# ---------------------------------------------------------------------------
# Glavna primerjalna tabela
# ---------------------------------------------------------------------------

st.subheader("📋 Primerjalna tabela")

display_df = comparison.copy()
display_df["Arrival"] = display_df["Arrival"].dt.strftime("%d.%m.%Y")
display_df["Departure"] = display_df["Departure"].dt.strftime("%d.%m.%Y")
display_df["Zacetek_ROS"] = display_df["Zacetek_ROS"].dt.strftime("%d.%m.%Y")
display_df["Konec_ROS"] = display_df["Konec_ROS"].dt.strftime("%d.%m.%Y")

cols_map = {
    "Reservation number": "Rezervacija (Booking.com)",
    "Guest name": "Gost",
    "Status": "Status",
    "Arrival": "Prihod (BC)",
    "Departure": "Odhod (BC)",
    "Nights_BC": "Noči (BC)",
    "Zacetek_ROS": "Začetek (ROS)",
    "Konec_ROS": "Konec (ROS)",
    "Nights_ROS": "Noči (ROS)",
    "Datumi_OK": "Datumi OK?",
    "Cena_ROS": "Cena/dan (ROS)",
    "Promet_ROS": "Promet (ROS)",
    "Dodatek_ROS": "Dodatek (ROS)",
    "Skupaj_ROS": "Skupaj ROS",
    "Final amount": "Final amount (BC)",
    "Razlika_znesek": "Razlika (Skupaj ROS - Final)",
    "Znesek_OK": "Znesek OK?",
    "Commission %": "Provizija %",
    "Commission amount": "Provizija znesek (BC)",
    "Provizija_pricakovana": "Provizija pričakovana",
    "Provizija_razlika": "Razlika provizije",
    "Provizija_OK": "Provizija OK?",
    "Neto_znesek_Booking": "Neto znesek (Final - Provizija)",
}
display_df = display_df[list(cols_map.keys())].rename(columns=cols_map)

if show_only_mismatch:
    display_df = display_df[display_df["Status"] != "✅ OK"]

st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Promet (ROS)": st.column_config.NumberColumn(format="%.2f €"),
        "Dodatek (ROS)": st.column_config.NumberColumn(format="%.2f €"),
        "Skupaj ROS": st.column_config.NumberColumn(format="%.2f €"),
        "Final amount (BC)": st.column_config.NumberColumn(format="%.2f €"),
        "Razlika (Skupaj ROS - Final)": st.column_config.NumberColumn(format="%.2f €"),
        "Provizija znesek (BC)": st.column_config.NumberColumn(format="%.2f €"),
        "Provizija pričakovana": st.column_config.NumberColumn(format="%.2f €"),
        "Razlika provizije": st.column_config.NumberColumn(format="%.2f €"),
        "Neto znesek (Final - Provizija)": st.column_config.NumberColumn(format="%.2f €"),
    },
)

csv_export = display_df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
st.download_button(
    "⬇️ Prenesi primerjalno tabelo (CSV)",
    data=csv_export,
    file_name="primerjava_booking_ros.csv",
    mime="text/csv",
)

st.divider()

# ---------------------------------------------------------------------------
# Neujemajoče / manjkajoče rezervacije iz ROS strani (kanal BOOKING.COM,
# a brez ujemajočega Reservation number v Booking.com datoteki)
# ---------------------------------------------------------------------------

st.subheader("🔍 ROS rezervacije s kanalom BOOKING.COM brez ujemanja v Booking.com statementu")

bc_numbers = set(bc_df["Reservation number"])
ros_booking_channel = ros_df[
    ros_df["Naziv skupine"].astype(str).str.upper().str.contains("BOOKING", na=False)
    | ros_df["Referenca"].astype(str).str.match(r"^\d{9,10}$", na=False)
]
missing_in_bc = ros_booking_channel[~ros_booking_channel["Referenca"].isin(bc_numbers)]

if missing_in_bc.empty:
    st.success("Ni ROS rezervacij s kanalom Booking.com, ki bi manjkale v Booking.com statementu.")
else:
    show = missing_in_bc[[
        "Rezervacija_ROS", "Referenca", "Gost_ROS", "Zacetek_ROS", "Konec_ROS", "Promet_ROS"
    ]].rename(columns={
        "Rezervacija_ROS": "Rezervacija (ROS)",
        "Referenca": "Referenca (Booking št.)",
        "Gost_ROS": "Gost",
        "Zacetek_ROS": "Začetek",
        "Konec_ROS": "Konec",
        "Promet_ROS": "Promet (ROS)",
    })
    st.dataframe(show, use_container_width=True, hide_index=True)
    st.caption(
        "To so lahko rezervacije iz drugega meseca/obdobja statementa, storno rezervacije, "
        "ali napačno/manjkajoče vnesena referenca v ROS."
    )

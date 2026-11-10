# -*- coding: utf-8 -*-
"""
app10.py - Planification des actions radio selon le taux d'utilisation des PRB.

Lancement :  streamlit run app10.py

Fichier Excel attendu (2 feuilles) :
  Feuille 1 : Site | Longitude | Latitude | Technologie | Config
  Feuille 2 : soit  Site | PRB          (PRB par site)
              soit  Cellule | PRB       (PRB par cellule -> moyenne par site)
"""

import io
import re

import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# Constantes
# ----------------------------------------------------------------------------
R_TERRE_M = 6371000.0

RULE_L1800 = "Add L1800"
RULE_L2600 = "Add L2600"
RULE_5G = "Add 5G"
RULE_NEW_MIMO = "New Site / Mimo"
RULE_NONE = "Aucune action"
RULE_OPTIONS = [RULE_L1800, RULE_L2600, RULE_5G, RULE_NEW_MIMO, RULE_NONE]

# Cellule type : K_40_AGENCE-KARA_L18_1  ->  site = AGENCE-KARA
CELL_RE = re.compile(r"^[A-Za-z]+_\d+_(.+)_L\d+_\d+$")

# Sites/cellules "Penta" : ADAWLATO2-Penta-0 -> ADAWLATO2 ; BIA-PENTA-L1 -> BIA
PENTA_RE = re.compile(r"[-_ ]*penta.*$", re.IGNORECASE)

# Config attribuée aux sites présents en feuille 2 mais absents de la feuille 1
DEFAULT_MISSING_CONFIG = "Config2"

# Palette inspirée de l'identité Yas (jaune + bleu profond)
YAS_YELLOW = "#FFD400"
YAS_YELLOW_SOFT = "#FFF3B0"
YAS_BLUE = "#1A2A8C"
YAS_BLUE_DARK = "#0F1B5C"
YAS_BG = "#F6F7FB"
YAS_TEXT = "#14213D"

# Couleur associée à chaque action (cartes + tableau)
ACTION_COLORS = {
    "Add L1800": "#2F6FED",
    "Add L2600": "#00A6A6",
    "Add 5G": "#7B3FE4",
    "New Site": "#E5484D",
    "Mimo": "#F59E0B",
    "Aucune action": "#94A3B8",
    "PRB manquant": "#CBD5E1",
}


# ----------------------------------------------------------------------------
# Utilitaires de lecture / nettoyage
# ----------------------------------------------------------------------------
def norm_site(x) -> str:
    """Clé de comparaison pour les noms de sites."""
    return str(x).strip().upper()


def norm_config(x) -> str:
    """'Config Rural High' -> 'config_rural_high' ; 'Config1_5G' -> 'config1_5g'."""
    return re.sub(r"[\s\-]+", "_", str(x).strip().lower())


def default_rule_for(config_key: str) -> str:
    """Règle par défaut associée à une config."""
    if config_key == "config5":
        return RULE_L1800
    if config_key in ("config2", "config4") or "rural" in config_key:
        return RULE_L2600          # Config2 / Config4 / Rural High
    if config_key == "config1":
        return RULE_5G
    if config_key in ("config1_5g", "config2_5g"):
        return RULE_NEW_MIMO       # Config1_5G / Config2_5G
    return RULE_NONE


def clean_site_name(name: str) -> str:
    """Supprime le suffixe PENTA : 'ADAWLATO2-Penta-0' -> 'ADAWLATO2'."""
    name = str(name).strip()
    cleaned = PENTA_RE.sub("", name).strip()
    return cleaned if cleaned else name


def find_col(df: pd.DataFrame, candidates):
    """Retourne le nom réel de la colonne correspondant à l'un des candidats."""
    low = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand in low:
            return low[cand]
    return None


def site_from_cell(cell: str) -> str:
    """Extrait le nom du site à partir du nom de la cellule."""
    cell = str(cell).strip()
    m = CELL_RE.match(cell)
    if m:
        return m.group(1)
    parts = cell.split("_")
    if len(parts) >= 4:
        return "_".join(parts[2:-2])
    return cell


def load_sheets(uploaded):
    xls = pd.ExcelFile(uploaded)
    if len(xls.sheet_names) < 2:
        raise ValueError("Le fichier doit contenir au moins deux feuilles.")
    s1 = xls.parse(xls.sheet_names[0])
    s2 = xls.parse(xls.sheet_names[1])
    return s1, s2, xls.sheet_names[:2]


def prepare_sites(s1: pd.DataFrame) -> pd.DataFrame:
    col_site = find_col(s1, ["site", "sites", "nom", "nom site"])
    col_lon = find_col(s1, ["longitude", "lon", "long"])
    col_lat = find_col(s1, ["latitude", "lat"])
    col_cfg = find_col(s1, ["config", "configuration"])
    missing = [n for n, c in [("Site", col_site), ("Longitude", col_lon),
                              ("Latitude", col_lat), ("Config", col_cfg)] if c is None]
    if missing:
        raise ValueError("Colonnes introuvables dans la feuille 1 : " + ", ".join(missing))

    df = pd.DataFrame({
        "Site": s1[col_site].astype(str).str.strip(),
        "Longitude": pd.to_numeric(s1[col_lon], errors="coerce"),
        "Latitude": pd.to_numeric(s1[col_lat], errors="coerce"),
        "Config": s1[col_cfg].astype(str).str.strip(),
    })
    df["key"] = df["Site"].map(norm_site)
    df["config_key"] = df["Config"].map(norm_config)
    df = df.dropna(subset=["Longitude", "Latitude"])
    df = df.drop_duplicates(subset="key").reset_index(drop=True)
    return df


def detect_mode(s2: pd.DataFrame, site_keys: set, forced: str) -> str:
    """Retourne 'sites' ou 'cellules'."""
    if forced == "Sites":
        return "sites"
    if forced == "Cellules":
        return "cellules"
    first = s2.columns[0]
    if "cell" in str(first).lower():
        return "cellules"
    if "site" in str(first).lower():
        return "sites"
    vals = s2[first].dropna().astype(str).map(norm_site)
    ratio = vals.isin(site_keys).mean() if len(vals) else 0
    return "sites" if ratio >= 0.5 else "cellules"


def prepare_prb(s2: pd.DataFrame, mode: str):
    """Retourne le PRB par site : DataFrame[key, Site, PRB, nb_cellules]."""
    name_col = s2.columns[0]
    prb_col = find_col(s2, ["prb", "prb%", "prb (%)", "% prb"])
    if prb_col is None:
        nums = [c for c in s2.columns[1:] if pd.api.types.is_numeric_dtype(s2[c])]
        if not nums:
            raise ValueError("Aucune colonne PRB numérique trouvée dans la feuille 2.")
        prb_col = nums[0]

    d = pd.DataFrame({
        "nom": s2[name_col].astype(str).str.strip(),
        "PRB": pd.to_numeric(s2[prb_col], errors="coerce"),
    }).dropna(subset=["PRB"])

    # PRB exprimé en fraction (0-1) -> pourcentage
    if len(d) and d["PRB"].max() <= 1.0:
        d["PRB"] = d["PRB"] * 100.0

    if mode == "cellules":
        d["Site"] = d["nom"].map(site_from_cell)
    else:
        d["Site"] = d["nom"]
    # Les cellules / sites "PENTA" sont rattachés au site principal
    d["Site"] = d["Site"].map(clean_site_name)
    d["key"] = d["Site"].map(norm_site)

    # Moyenne des PRB des cellules d'un même site (ou des doublons de site)
    g = d.groupby("key", as_index=False).agg(
        Site=("Site", "first"), PRB=("PRB", "mean"), nb_cellules=("PRB", "size")
    )
    return g


def add_missing_sites(sites: pd.DataFrame, prb_site: pd.DataFrame, config_label: str):
    """
    Ajoute les sites présents en feuille 2 mais absents de la feuille 1.
    Ils sont considérés comme des sites de `config_label` (Config2) ; leurs
    coordonnées sont inconnues (vides).
    """
    missing = prb_site[~prb_site["key"].isin(set(sites["key"]))]
    if missing.empty:
        return sites, 0
    extra = pd.DataFrame({
        "Site": missing["Site"].values,
        "Longitude": np.nan,
        "Latitude": np.nan,
        "Config": config_label,
        "key": missing["key"].values,
        "config_key": norm_config(config_label),
    })
    return pd.concat([sites, extra], ignore_index=True), len(extra)


# ----------------------------------------------------------------------------
# Géographie
# ----------------------------------------------------------------------------
def haversine_m(lat1, lon1, lat2, lon2):
    """Distance (m) - accepte des tableaux numpy (diffusion)."""
    p1, p2 = np.radians(lat1), np.radians(lat2)
    dphi = p2 - p1
    dl = np.radians(lon2) - np.radians(lon1)
    a = np.sin(dphi / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dl / 2) ** 2
    return 2 * R_TERRE_M * np.arcsin(np.sqrt(a))


def destination(lat, lon, bearing_deg, dist_m):
    """Point atteint depuis (lat, lon) selon un azimut et une distance."""
    d = dist_m / R_TERRE_M
    br = np.radians(bearing_deg)
    la1, lo1 = np.radians(lat), np.radians(lon)
    la2 = np.arcsin(np.sin(la1) * np.cos(d) + np.cos(la1) * np.sin(d) * np.cos(br))
    lo2 = lo1 + np.arctan2(np.sin(br) * np.sin(d) * np.cos(la1),
                           np.cos(d) - np.sin(la1) * np.sin(la2))
    return np.degrees(la2), np.degrees(lo2)


def propose_new_site(lat0, lon0, all_lat, all_lon, target_m=350.0, tol_m=50.0):
    """
    Propose les coordonnées d'un nouveau site autour du site saturé (lat0, lon0).
    - On teste des points sur des anneaux autour du site (350 m, 400 m, ... 1050 m)
      et tous les 10 degrés d'azimut.
    - On retient, sur l'anneau le plus proche, le point dont la distance au site
      existant le plus proche est maximale, à condition qu'elle soit >= target - tol
      (donc idéalement ~350 m des sites environnants).
    - Si aucun point ne respecte la contrainte, on prend le meilleur compromis.
    """
    bearings = np.arange(0, 360, 10)
    radii = target_m + np.array([0, 50, 100, 150, 250, 350, 500, 700])
    min_ok = target_m - tol_m

    best = None  # (min_dist, lat, lon)
    for r in radii:
        cl, co = destination(lat0, lon0, bearings, r)
        dist = haversine_m(cl[:, None], co[:, None], all_lat[None, :], all_lon[None, :])
        mind = dist.min(axis=1)
        i = int(np.argmax(mind))
        if mind[i] >= min_ok:
            return float(cl[i]), float(co[i])
        if best is None or mind[i] > best[0]:
            best = (float(mind[i]), float(cl[i]), float(co[i]))
    return best[1], best[2]


# ----------------------------------------------------------------------------
# Algorithme
# ----------------------------------------------------------------------------
def run_algorithm(sites: pd.DataFrame, prb_site: pd.DataFrame, rules: dict, params: dict):
    df = sites.merge(prb_site[["key", "PRB"]], on="key", how="left")

    lat = df["Latitude"].to_numpy(float)
    lon = df["Longitude"].to_numpy(float)
    prb = df["PRB"].to_numpy(float)
    n = len(df)

    # Matrice des distances entre sites (m) pour la recherche des voisins
    dmat = haversine_m(lat[:, None], lon[:, None], lat[None, :], lon[None, :])
    np.fill_diagonal(dmat, np.inf)

    actions = [""] * n
    new_lat = [np.nan] * n
    new_lon = [np.nan] * n

    # Coordonnées de tous les sites (existants + nouveaux déjà proposés)
    placed_lat, placed_lon = [], []

    for i in range(n):
        p = prb[i]
        rule = rules.get(df.at[i, "config_key"], RULE_NONE)

        if np.isnan(p):
            actions[i] = "PRB manquant"
            continue

        if rule == RULE_L1800:
            actions[i] = RULE_L1800 if p > params["thr_l1800"] else "Aucune action"

        elif rule == RULE_L2600:
            actions[i] = RULE_L2600 if p > params["thr_l2600"] else "Aucune action"

        elif rule == RULE_5G:
            actions[i] = RULE_5G if p > params["thr_5g"] else "Aucune action"

        elif rule == RULE_NEW_MIMO:
            if p > params["thr_new"]:
                mask = (dmat[i] <= params["radius_m"]) & ~np.isnan(prb)
                nb = prb[mask]
                if len(nb) == 0:
                    nb_ok = False
                elif params["nb_mode"] == "Moyenne des voisins":
                    nb_ok = float(np.mean(nb)) > params["thr_nb"]
                elif params["nb_mode"] == "Au moins un voisin":
                    nb_ok = bool((nb > params["thr_nb"]).any())
                else:  # Tous les voisins
                    nb_ok = bool((nb > params["thr_nb"]).all())

                if nb_ok and not (np.isnan(lat[i]) or np.isnan(lon[i])):
                    actions[i] = "New Site"
                    all_lat = np.concatenate([lat, np.array(placed_lat)])
                    all_lon = np.concatenate([lon, np.array(placed_lon)])
                    ok = ~(np.isnan(all_lat) | np.isnan(all_lon))   # ignore les sites sans coordonnées
                    all_lat, all_lon = all_lat[ok], all_lon[ok]
                    nl, no = propose_new_site(lat[i], lon[i], all_lat, all_lon,
                                              params["target_m"], params["tol_m"])
                    new_lat[i], new_lon[i] = nl, no
                    placed_lat.append(nl)
                    placed_lon.append(no)
                else:
                    actions[i] = "Mimo"
            else:
                # Lecture littérale : "sinon Mimo" même si PRB <= seuil
                actions[i] = "Mimo" if params["mimo_literal"] else "Aucune action"
        else:
            actions[i] = "Aucune action"

    out = pd.DataFrame({
        "Sites": df["Site"],
        "Longitude": df["Longitude"],
        "Latitude": df["Latitude"],
        "Config": df["Config"],
        "PRB": df["PRB"].round(2),
        "Action": actions,
        "Longitude New Site": pd.Series(new_lon).round(6),
        "Latitude New Site": pd.Series(new_lat).round(6),
    })
    return out


# ----------------------------------------------------------------------------
# Export Excel
# ----------------------------------------------------------------------------
def to_excel_bytes(result: pd.DataFrame, prb_site_view: pd.DataFrame) -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="Resultats", index=False)
        prb_site_view.to_excel(writer, sheet_name="PRB_par_site", index=False)

        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for idx, col in enumerate(ws.columns, start=1):
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[get_column_letter(idx)].width = min(max(width + 3, 12), 45)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Habillage (CSS / composants visuels)
# ----------------------------------------------------------------------------
CSS = """
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', 'Inter', system-ui, sans-serif; }
.stApp { background: %%BG%%; }
.block-container { padding-top: 1.4rem; max-width: 1400px; }
header[data-testid="stHeader"] { background: transparent; }
footer { visibility: hidden; }

/* Texte principal */
[data-testid="stMain"] p, [data-testid="stMain"] li, [data-testid="stMain"] label,
section.main p, section.main li, section.main label { color: %%TEXT%%; }

/* ---------- Bandeau ---------- */
.yas-hero {
    background: linear-gradient(120deg, %%YELLOW%% 0%, #FFE766 100%);
    border-radius: 20px; padding: 1.5rem 2rem; margin-bottom: 1.2rem;
    display: flex; align-items: center; gap: 1.2rem;
    box-shadow: 0 6px 22px rgba(26,42,140,.14);
}
.yas-hero .logo {
    background: %%BLUE%%; color: %%YELLOW%% !important; font-weight: 900;
    font-size: 1.7rem; letter-spacing: .02em; border-radius: 16px;
    padding: .35rem 1.1rem; line-height: 1.2;
}
.yas-hero .title { color: %%BLUE%% !important; font-size: 1.85rem; font-weight: 800; line-height: 1.15; }
.yas-hero .sub { color: %%BLUE_DARK%% !important; opacity: .85; font-size: .98rem; margin-top: .25rem; }

/* ---------- Titres de section ---------- */
.sec { display: flex; align-items: center; gap: .65rem; margin: 1.5rem 0 .7rem; }
.sec .n {
    background: %%YELLOW%%; color: %%BLUE%% !important; font-weight: 800; border-radius: 50%;
    width: 2rem; height: 2rem; display: flex; align-items: center; justify-content: center;
    box-shadow: 0 2px 6px rgba(0,0,0,.12);
}
.sec .t { color: %%BLUE%% !important; font-size: 1.3rem; font-weight: 800; }

/* ---------- Cartes KPI ---------- */
.kpi {
    background: #fff; border-radius: 14px; padding: .9rem 1rem .8rem;
    box-shadow: 0 2px 12px rgba(26,42,140,.08); border-top: 5px solid var(--c);
}
.kpi .v { font-size: 1.95rem; font-weight: 800; color: %%BLUE%% !important; line-height: 1.1; }
.kpi .l { font-size: .74rem; font-weight: 700; color: #5B6490 !important;
          text-transform: uppercase; letter-spacing: .05em; }

/* ---------- Cartes d'accueil ---------- */
.step-card {
    background: #fff; border-radius: 16px; padding: 1.1rem 1.2rem; height: 100%;
    box-shadow: 0 2px 12px rgba(26,42,140,.08); border-left: 6px solid %%YELLOW%%;
}
.step-card .h { color: %%BLUE%% !important; font-weight: 800; margin-bottom: .3rem; }
.step-card .d { color: %%TEXT%% !important; font-size: .93rem; }

/* ---------- Barre latérale ---------- */
[data-testid="stSidebar"] { background: linear-gradient(180deg, %%BLUE%% 0%, %%BLUE_DARK%% 100%); }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: %%YELLOW%% !important; }
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
[data-testid="stSidebar"] [data-testid="stCheckbox"] p,
[data-testid="stSidebar"] [data-testid="stRadio"] label p,
[data-testid="stSidebar"] .stMarkdown p { color: #F4F6FF !important; }
/* Cases des paramètres : fond bleu translucide + texte blanc */
[data-testid="stSidebar"] [data-baseweb="input"],
[data-testid="stSidebar"] [data-baseweb="base-input"],
[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background: rgba(255,255,255,.14) !important;
    border: 1px solid rgba(255,255,255,.40) !important;
    border-radius: 10px !important;
}
[data-testid="stSidebar"] input {
    color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important;
    background: transparent !important; font-weight: 600;
}
[data-testid="stSidebar"] [data-baseweb="select"] *,
[data-testid="stSidebar"] [data-testid="stNumberInputStepUp"],
[data-testid="stSidebar"] [data-testid="stNumberInputStepDown"] { color: #FFFFFF !important; }
[data-testid="stSidebar"] [data-baseweb="select"] svg,
[data-testid="stSidebar"] [data-testid="stNumberInputStepUp"] svg,
[data-testid="stSidebar"] [data-testid="stNumberInputStepDown"] svg { fill: #FFFFFF !important; }
[data-testid="stSidebar"] [data-testid="stNumberInputStepUp"],
[data-testid="stSidebar"] [data-testid="stNumberInputStepDown"] { background: transparent !important; }

/* ---------- Onglets ---------- */
.stTabs [data-baseweb="tab-list"] { gap: 6px; }
.stTabs button[role="tab"] {
    border-radius: 12px 12px 0 0; font-weight: 700; padding: .5rem 1.1rem;
    background: #E9ECF8;
}
.stTabs button[role="tab"] p { color: %%BLUE%% !important; font-weight: 700; }
.stTabs button[role="tab"][aria-selected="true"] { background: %%YELLOW%%; }
.stTabs [data-baseweb="tab-highlight"] { background-color: %%BLUE%% !important; }

/* ---------- Boutons ---------- */
[data-testid="stDownloadButton"] button, button[kind="primary"],
[data-testid="stBaseButton-primary"] {
    background: %%YELLOW%% !important; color: %%BLUE%% !important; border: none !important;
    font-weight: 800 !important; border-radius: 12px !important; padding: .55rem 1.3rem !important;
    box-shadow: 0 3px 10px rgba(255,212,0,.45);
}
[data-testid="stDownloadButton"] button p, button[kind="primary"] p,
[data-testid="stBaseButton-primary"] p { color: %%BLUE%% !important; font-weight: 800; }
[data-testid="stDownloadButton"] button:hover, button[kind="primary"]:hover,
[data-testid="stBaseButton-primary"]:hover { background: #FFE24D !important; transform: translateY(-1px); }

/* ---------- Zone de dépôt ---------- */
[data-testid="stFileUploader"] section {
    border: 2px dashed %%YELLOW%%; background: #FFFBE0; border-radius: 16px;
}
[data-testid="stFileUploader"] section * { color: %%BLUE%% !important; }
[data-testid="stFileUploader"] section button { background: %%BLUE%% !important; }
[data-testid="stFileUploader"] section button * { color: %%YELLOW%% !important; }

/* ---------- Divers ---------- */
[data-testid="stExpander"] { border: 1px solid #E3E6F3; border-radius: 14px; background: #fff; }
[data-testid="stAlert"] { border-radius: 12px; }
[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; box-shadow: 0 2px 12px rgba(26,42,140,.08); }
.yas-footer { text-align: center; color: #7A83B0; font-size: .8rem; margin-top: 2rem; }
</style>
"""


def inject_css():
    css = (CSS.replace("%%YELLOW%%", YAS_YELLOW).replace("%%BLUE_DARK%%", YAS_BLUE_DARK)
              .replace("%%BLUE%%", YAS_BLUE).replace("%%BG%%", YAS_BG).replace("%%TEXT%%", YAS_TEXT))
    st.markdown(css, unsafe_allow_html=True)


def hero():
    st.markdown(
        """
        <div class="yas-hero">
            <div class="logo">yas</div>
            <div>
                <div class="title">Analyse PRB &amp; recommandations d'actions</div>
                <div class="sub">Optimisation de la capacité radio : Add L1800 · Add L2600 · Add 5G · Mimo · New Site</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section(num: str, title: str):
    st.markdown(f'<div class="sec"><div class="n">{num}</div><div class="t">{title}</div></div>',
                unsafe_allow_html=True)


def kpi_card(label: str, value: int, color: str) -> str:
    return (f'<div class="kpi" style="--c:{color}"><div class="v">{value:,}</div>'
            f'<div class="l">{label}</div></div>').replace(",", " ")


def style_results(df: pd.DataFrame):
    """Mise en forme du tableau de résultats (couleur de l'action, décimales)."""
    def color_action(v):
        c = ACTION_COLORS.get(v)
        if not c:
            return ""
        txt = YAS_BLUE_DARK if v in ("Mimo", "PRB manquant", "Aucune action") else "#FFFFFF"
        return f"background-color: {c}; color: {txt}; font-weight: 700;"

    # Affichage : cases vides plutôt que "None" pour les coordonnées absentes
    d = df.copy()
    for c in ["Longitude", "Latitude", "Longitude New Site", "Latitude New Site"]:
        d[c] = d[c].map(lambda v: "" if pd.isna(v) else f"{v:.6f}")

    sty = d.style
    apply_map = getattr(sty, "map", None) or sty.applymap
    sty = apply_map(color_action, subset=["Action"])
    return sty.format({"PRB": "{:.2f}"}, na_rep="")


def welcome():
    c1, c2, c3 = st.columns(3)
    cards = [
        ("1 · Préparez le fichier",
         "Feuille 1 : Site, Longitude, Latitude, Technologie, Config. "
         "Feuille 2 : PRB par site <b>ou</b> par cellule."),
        ("2 · Chargez-le",
         "Le contenu de la feuille 2 est détecté automatiquement. "
         "Les PRB des cellules d'un même site sont moyennés (Penta rattachés au site principal)."),
        ("3 · Téléchargez",
         "L'algorithme s'exécute immédiatement. Récupérez le tableau Excel avec les actions "
         "et les coordonnées des New Sites."),
    ]
    for col, (h, d) in zip((c1, c2, c3), cards):
        col.markdown(f'<div class="step-card"><div class="h">{h}</div><div class="d">{d}</div></div>',
                     unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# Interface Streamlit
# ----------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Yas - Analyse PRB", page_icon="📡", layout="wide")
    inject_css()
    hero()

    # ---- Paramètres (barre latérale)
    with st.sidebar:
        st.header("⚙️ Paramètres")
        st.subheader("Seuils PRB (%)")
        thr_l1800 = st.number_input("Config5 → Add L1800 si PRB >", 0.0, 100.0, 80.0, 1.0)
        thr_l2600 = st.number_input("Config2 / Rural High → Add L2600 si PRB >", 0.0, 100.0, 90.0, 1.0)
        thr_5g = st.number_input("Config1 → Add 5G si PRB >", 0.0, 100.0, 90.0, 1.0)
        thr_new = st.number_input("Config1_5G → seuil PRB du site >", 0.0, 100.0, 90.0, 1.0)
        thr_nb = st.number_input("Config1_5G → seuil PRB des voisins >", 0.0, 100.0, 80.0, 1.0)

        st.subheader("Sites voisins")
        radius_km = st.number_input("Rayon de voisinage (km)", 0.5, 20.0, 3.0, 0.5)
        nb_mode = st.selectbox(
            "Critère sur les voisins",
            ["Moyenne des voisins", "Au moins un voisin", "Tous les voisins"],
        )
        mimo_literal = st.checkbox(
            "Config1_5G : appliquer 'Mimo' même si PRB ≤ seuil (lecture littérale de 'sinon')",
            value=False,
        )

        st.subheader("Position du New Site")
        target_m = st.number_input("Distance cible aux sites voisins (m)", 100, 1500, 350, 10)
        tol_m = st.number_input("Tolérance ± (m)", 0, 500, 50, 10)

        st.subheader("Feuille 2")
        forced = st.radio("Type de contenu", ["Auto", "Sites", "Cellules"], horizontal=True)

    params = dict(
        thr_l1800=thr_l1800, thr_l2600=thr_l2600, thr_5g=thr_5g, thr_new=thr_new,
        thr_nb=thr_nb, radius_m=radius_km * 1000.0, nb_mode=nb_mode,
        mimo_literal=mimo_literal, target_m=float(target_m), tol_m=float(tol_m),
    )

    uploaded = st.file_uploader("📂 Chargez votre fichier Excel (.xlsx)", type=["xlsx", "xls"])
    if uploaded is None:
        welcome()
        return

    try:
        s1, s2, names = load_sheets(uploaded)
        sites = prepare_sites(s1)
        mode = detect_mode(s2, set(sites["key"]), forced)
        prb_site = prepare_prb(s2, mode)
    except Exception as e:  # noqa: BLE001
        st.error(f"Erreur de lecture du fichier : {e}")
        return

    # Sites de la feuille 2 absents de la feuille 1 -> considérés comme Config2
    sites_f1 = len(sites)
    sites, n_extra = add_missing_sites(sites, prb_site, DEFAULT_MISSING_CONFIG)

    st.success(
        f"Feuille 1 « {names[0]} » : {sites_f1} sites  |  "
        f"Feuille 2 « {names[1]} » : contenu détecté = **{mode}**"
    )

    # ---- Association config -> règle (modifiable)
    with st.sidebar:
        st.subheader("Règle par Config")
        rules = {}
        cfg_labels = sites.drop_duplicates("config_key").set_index("config_key")["Config"]
        for ck, label in cfg_labels.items():
            default = default_rule_for(ck)
            rules[ck] = st.selectbox(
                f"{label}", RULE_OPTIONS, index=RULE_OPTIONS.index(default), key=f"rule_{ck}"
            )

    # ---- Étape 1 : PRB par site
    section("1", "PRB par site" + (" (moyenne des cellules)" if mode == "cellules" else ""))
    if mode == "cellules":
        n_cells = len(s2.dropna(how="all"))
        st.write(f"{n_cells} cellules → {len(prb_site)} sites (moyenne des PRB des cellules d'un même site).")

    n_nodata = int((~sites["key"].isin(set(prb_site["key"]))).sum())
    if n_nodata:
        st.info(f"{n_nodata} sites de la feuille 1 n'ont pas de PRB (action = « PRB manquant »).")

    name_map = sites.set_index("key")["Site"]
    prb_view = prb_site.copy()
    prb_view["Site"] = prb_view["key"].map(name_map)
    prb_view = prb_view[["Site", "PRB", "nb_cellules"]].rename(columns={"PRB": "PRB moyen (%)"})
    prb_view["PRB moyen (%)"] = prb_view["PRB moyen (%)"].round(2)

    # ---- Étape 2 : algorithme
    section("2", "Résultats de l'algorithme")
    with st.spinner("Exécution de l'algorithme..."):
        result = run_algorithm(sites, prb_site, rules, params)

    counts = result["Action"].value_counts()
    kpis = [RULE_L1800, RULE_L2600, RULE_5G, "New Site", "Mimo", "Aucune action"]
    for col, act in zip(st.columns(len(kpis)), kpis):
        col.markdown(kpi_card(act, int(counts.get(act, 0)), ACTION_COLORS[act]), unsafe_allow_html=True)
    st.write("")

    only_actions = st.checkbox("N'afficher / exporter que les sites avec une action", value=True)
    shown = result.copy()
    if only_actions:
        shown = shown[~shown["Action"].isin(["Aucune action", "PRB manquant"])]
    shown = shown.sort_values(["Action", "PRB"], ascending=[True, False]).reset_index(drop=True)
    new_df = shown[shown["Action"] == "New Site"]

    tab_res, tab_prb, tab_map = st.tabs(
        ["📋 Résultats", "📈 PRB par site", f"🗺️ Carte des New Sites ({len(new_df)})"]
    )

    with tab_res:
        st.download_button(
            "⬇️ Télécharger le résultat (Excel)",
            data=to_excel_bytes(shown, prb_view),
            file_name="resultat_actions_PRB.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.dataframe(style_results(shown), hide_index=True)

    with tab_prb:
        if n_extra:
            extra_names = sites.iloc[sites_f1:]["Site"].tolist()
            with st.expander(
                f"ℹ️ {n_extra} sites de la feuille 2 absents de la feuille 1 : "
                f"traités comme {DEFAULT_MISSING_CONFIG} (sans coordonnées)"
            ):
                st.write(sorted(extra_names))
        st.dataframe(prb_view, hide_index=True)

    with tab_map:
        if new_df.empty:
            st.info("Aucun New Site proposé avec les paramètres actuels.")
        else:
            pts = new_df.rename(columns={"Latitude New Site": "latitude",
                                         "Longitude New Site": "longitude"})[["latitude", "longitude"]]
            try:
                st.map(pts, color=YAS_BLUE)
            except TypeError:  # anciennes versions de Streamlit
                st.map(pts)

    st.markdown('<div class="yas-footer">Yas · Analyse PRB &amp; planification radio</div>',
                unsafe_allow_html=True)


if __name__ == "__main__":
    main()
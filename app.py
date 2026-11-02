import hashlib
import html
import io
import json
import re
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path

import gspread
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials
from PIL import Image, ImageOps

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

# =====================================================
# 1. CONFIGURATION
# =====================================================
st.set_page_config(page_title="Agent Colis Pro", page_icon="📦", layout="wide")

SHEET_ID = "1m3FoOir2i3Je2YgMoyJzjuXSetbWYX74Qz8Ylw03Jeg"
ONGLET_DEFAUT = "Suivi_Colis"
MODELE = "gemini-3.8-flash"
MODELES_PAR_DEFAUT = "gemini-3.8-flash, gemini-flash-latest, gemini-flash-lite-latest"
THINKING_LEVEL = "low"
TIMEOUT_MS = 300000
MAX_WORKERS = 4
TAILLE_IMAGE = (800, 800)
SAUVEGARDE = Path(__file__).with_name("colis_sauvegarde.json")

CLES = ["nom_client", "numero_telephone", "ville", "pays", "infos_fournisseur",
        "poids", "quantite", "contenu", "prix_unitaire", "montant"]
COLONNES_ORDRE = ["ID", "Date", "nom_client", "numero_telephone", "ville", "pays",
                  "infos_fournisseur", "poids", "quantite", "contenu", "prix_unitaire",
                  "montant", "analyse_agent", "hash_photos"]
ENTETES = ["Name", "Country", "City", "Phone", "Weight", "Nber Of Box pieces", "Contain",
           "Unit Price", "Factory Name", "DATE RECU", "Montant", "Analyse IA"]
NB_COLS = len(ENTETES)

PROMPT = """Extrais les données du colis en JSON (un seul objet). Valeur "" si absente.
Clés : nom_client, numero_telephone, ville, pays, infos_fournisseur, poids, quantite, contenu, prix_unitaire, montant.
Règles : villes Cameroun (Douala/Yaoundé), RDC (Kinshasa). Préfixes tél (+243, +237, +221...). Poids souvent en kg."""

# Palette (valeurs 0-1) et dimensions partagées par Google Sheets et Excel
BLEU_FONCE = (0.10, 0.20, 0.50)
BLEU_CLAIR = (0.60, 0.80, 1.00)
JAUNE = (1.00, 1.00, 0.00)
ORANGE = (1.00, 0.60, 0.00)
ORANGE_CLAIR = (1.00, 0.88, 0.70)
GRIS = (0.75, 0.75, 0.75)
BLANC = (1.0, 1.0, 1.0)
NOIR = (0.0, 0.0, 0.0)
HAUTEURS = {"titre": 42, "entete": 34, "ville": 30, "colis": 30, "soustotal": 28,
            "vide": 10, "total_kg": 36, "total_colis": 30}
LARGEURS = [150, 100, 100, 130, 100, 130, 120, 90, 150, 110, 90, 230]
FORMAT_KG = '#,##0.0 "kg"'


JOURS_FR = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
MOIS_FR = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
           "septembre", "octobre", "novembre", "décembre"]


def _version_streamlit():
    try:
        return tuple(int(x) for x in st.__version__.split(".")[:2])
    except Exception:
        return (0, 0)


# use_container_width est supprimé des versions récentes de Streamlit
LARGEUR = {"width": "stretch"} if _version_streamlit() >= (1, 50) else {"use_container_width": True}


# =====================================================
# 2. PROTECTION PAR MOT DE PASSE (optionnelle)
# =====================================================
def verifier_mot_de_passe():
    """Active si APP_PASSWORD est défini dans secrets.toml."""
    try:
        attendu = st.secrets["APP_PASSWORD"]
    except Exception:
        return
    if st.session_state.get("auth_ok"):
        return
    st.title("Agent Logistique Intelligent")
    mdp = st.text_input("Mot de passe", type="password")
    if mdp:
        if mdp == attendu:
            st.session_state.auth_ok = True
            st.rerun()
        else:
            st.error("Mot de passe incorrect.")
    st.stop()


verifier_mot_de_passe()


# =====================================================
# 3. DONNÉES : SESSION, SAUVEGARDE, REGROUPEMENT
# =====================================================
def aujourdhui():
    return date.today().isoformat()


def jour_lisible(jour):
    """'2026-11-02' -> 'lundi 02/11/2026'"""
    try:
        d = datetime.strptime(jour, "%Y-%m-%d")
        return f"{JOURS_FR[d.weekday()]} {d.strftime('%d/%m/%Y')}"
    except Exception:
        return str(jour)


def charger_sauvegarde():
    """Retourne ({jour: [colis]}, compteur). Lit aussi l'ancien format (une seule liste)."""
    try:
        if SAUVEGARDE.exists():
            d = json.loads(SAUVEGARDE.read_text(encoding="utf-8"))
            compteur = int(d.get("compteur", 1))
            if isinstance(d.get("jours"), dict):
                return d["jours"], compteur
            jours = {}
            for r in d.get("colis", []):
                jour = str(r.get("Date", ""))[:10]
                if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", jour):
                    jour = aujourdhui()
                jours.setdefault(jour, []).append(r)
            return jours, compteur
    except Exception:
        pass
    return {}, 1


def synchroniser_jour():
    """Recopie la liste affichée dans le dossier du jour chargé."""
    jour = st.session_state.get("jour_charge")
    if jour is not None and "jours" in st.session_state:
        st.session_state.jours[jour] = st.session_state.historique_colis


def sauvegarder():
    synchroniser_jour()
    if not st.session_state.get("autosave", True):
        return
    try:
        d = {"compteur": st.session_state.compteur_id,
             "jours": {j: l for j, l in st.session_state.jours.items() if l}}
        SAUVEGARDE.write_text(json.dumps(d, ensure_ascii=False, default=str), encoding="utf-8")
    except Exception:
        pass


def normaliser(records):
    """Nettoie les lignes : texte partout, ID et date pour les lignes ajoutées à la main."""
    propres = []
    for r in records:
        rec = {}
        for col in COLONNES_ORDRE:
            v = r.get(col, "")
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                v = ""
            rec[col] = v
        if not any(str(rec[c]).strip() for c in CLES):
            continue  # ligne vide ajoutée par erreur
        try:
            rec["ID"] = int(rec["ID"])
        except (TypeError, ValueError):
            rec["ID"] = st.session_state.compteur_id
            st.session_state.compteur_id += 1
        for col in COLONNES_ORDRE:
            if col != "ID":
                rec[col] = str(rec[col]).strip()
        if not rec["Date"]:
            jour = st.session_state.get("jour_charge") or aujourdhui()
            rec["Date"] = f"{jour} {datetime.now().strftime('%H:%M')}"
        propres.append(rec)
    return propres


if "jours" not in st.session_state:
    jours_charges, compteur = charger_sauvegarde()
    st.session_state.compteur_id = max(compteur, 1)
    st.session_state.jour_charge = aujourdhui()
    st.session_state.jours = {j: normaliser(l) for j, l in jours_charges.items() if isinstance(l, list)}
    st.session_state.historique_colis = st.session_state.jours.get(aujourdhui(), [])
    st.session_state.jour_actif = aujourdhui()
    st.session_state.restaure = sum(len(l) for l in st.session_state.jours.values())
for _cle, _val in (("cle_upload", 0), ("cle_ajout", 0), ("autosave", True), ("fusion_auto", True),
                   ("onglet_auto", True), ("nom_onglet", ONGLET_DEFAUT),
                   ("modeles_txt", MODELES_PAR_DEFAUT)):
    if _cle not in st.session_state:
        st.session_state[_cle] = _val


def tous_les_hashes():
    """Photos déjà analysées, tous jours confondus (évite de payer deux fois la même photo)."""
    synchroniser_jour()
    return {h for l in st.session_state.jours.values() for h in hashes_connus(l)}


# ---------- Nombres ----------
def _nombre_ou_none(texte):
    m = re.search(r"\d+(?:[.,]\d+)?", str(texte if texte is not None else ""))
    return float(m.group().replace(",", ".")) if m else None


def _fmt_nombre(x):
    return f"{x:g}"


def nb_pieces(q):
    """Nombre de colis d'une ligne : la quantité si elle est lisible, sinon 1."""
    n = _nombre_ou_none(q)
    return n if n is not None and n > 0 else 1.0


# ---------- Regroupement : même client + même destination ----------
def _jetons(texte):
    sans_accents = "".join(c for c in unicodedata.normalize("NFKD", str(texte))
                           if not unicodedata.combining(c))
    return re.findall(r"[a-z0-9]+", sans_accents.casefold())


def cle_regroupement(rec):
    """(nom normalisé, ville normalisée). L'ordre des mots du nom n'a pas d'importance."""
    nom = sorted(_jetons(rec.get("nom_client", "")))
    if not nom:
        return None
    return (" ".join(nom), "".join(_jetons(rec.get("ville", ""))))


def _distincts(valeurs, sep=" / "):
    vus, res = set(), []
    for v in valeurs:
        for p in str(v).split(sep):
            p = p.strip()
            if p and p.casefold() not in vus:
                vus.add(p.casefold())
                res.append(p)
    return res


def _premier(groupe, cle):
    return next((r[cle] for r in groupe if r.get(cle)), "")


def fusionner_groupe(groupe):
    base = dict(groupe[0])
    for cle in ("nom_client", "ville", "pays"):
        base[cle] = _premier(groupe, cle)

    # téléphones : on garde chaque numéro différent (comparés sur les chiffres)
    tels, vus = [], set()
    for r in groupe:
        for p in str(r.get("numero_telephone", "")).split(" / "):
            chiffres = re.sub(r"\D", "", p)
            if chiffres and chiffres not in vus:
                vus.add(chiffres)
                tels.append(p.strip())
    base["numero_telephone"] = " / ".join(tels)

    base["quantite"] = _fmt_nombre(sum(nb_pieces(r["quantite"]) for r in groupe))

    poids = [_nombre_ou_none(r["poids"]) for r in groupe]
    poids = [p for p in poids if p is not None]
    base["poids"] = f"{_fmt_nombre(sum(poids))} kg" if poids else ""

    for cle in ("infos_fournisseur", "contenu", "prix_unitaire"):
        base[cle] = " / ".join(_distincts([r[cle] for r in groupe]))

    montants = [r["montant"] for r in groupe if r["montant"]]
    if montants and all(re.fullmatch(r"\d+(?:[.,]\d+)?", m.strip()) for m in montants):
        base["montant"] = _fmt_nombre(sum(float(m.replace(",", ".")) for m in montants))
    else:
        base["montant"] = " / ".join(_distincts(montants))

    dates = [r["Date"] for r in groupe if r["Date"]]
    base["Date"] = min(dates) if dates else base["Date"]
    base["hash_photos"] = ",".join(_distincts([r["hash_photos"] for r in groupe], sep=","))
    modele = re.sub(r" · regroupé$", "", base.get("analyse_agent", ""))
    base["analyse_agent"] = f"{modele} · regroupé".strip(" ·")
    return base


def regrouper(records, ids_nouveaux=None):
    """Fusionne les lignes du même client et de la même destination.
    Si ids_nouveaux est fourni, seuls les groupes contenant un nouveau colis sont fusionnés."""
    ordre, groupes = [], {}
    for r in records:
        k = cle_regroupement(r)
        if k is None:
            ordre.append(("seul", r))
        else:
            if k not in groupes:
                groupes[k] = []
                ordre.append(("groupe", k))
            groupes[k].append(r)
    sortie = []
    for typ, x in ordre:
        if typ == "seul":
            sortie.append(x)
            continue
        g = groupes[x]
        concerne = ids_nouveaux is None or any(r["ID"] in ids_nouveaux for r in g)
        if len(g) == 1 or not concerne:
            sortie.extend(g)
        else:
            sortie.append(fusionner_groupe(g))
    return sortie


def hashes_connus(records):
    return {h for r in records for h in str(r.get("hash_photos", "")).split(",") if h}


# =====================================================
# 4. MISE EN PAGE COMMUNE (Google Sheets + Excel)
# =====================================================
def extraire_nombre(texte):
    if texte is None or (not isinstance(texte, str) and pd.isna(texte)):
        return 0.0
    texte = str(texte).strip()
    if texte in ("", "Non spécifié"):
        return 0.0
    match = re.search(r"\d+(?:\.\d+)?", texte.replace(",", "."))
    return float(match.group()) if match else 0.0


def formater_date(d):
    try:
        return datetime.strptime(str(d), "%Y-%m-%d %H:%M").strftime("%d - %m - %Y")
    except Exception:
        return str(d)


def construire_layout(df, jour=None):
    """Liste de (type_de_ligne, valeurs) : sert à Google Sheets et à Excel."""
    df = df.fillna("").copy()
    for c in COLONNES_ORDRE:
        if c not in df.columns:
            df[c] = ""
    for c in ("pays", "ville"):
        df[c] = df[c].astype(str).str.strip()

    try:
        d = datetime.strptime(jour, "%Y-%m-%d") if jour else datetime.now()
    except Exception:
        d = datetime.now()
    titre = f"SHIPPING DU {d.day:02d} {MOIS_FR[d.month - 1].upper()} {d.year}"
    lignes = [("titre", [titre] + [""] * (NB_COLS - 1)), ("entete", list(ENTETES))]

    total_general = 0.0
    premier = True
    tri = df.sort_values(by=["pays", "ville"], kind="stable")
    for (_, ville), groupe in tri.groupby(["pays", "ville"], sort=False):
        if not premier:
            lignes.append(("vide", [""] * NB_COLS))
        premier = False
        nom_ville = ville.upper() or "VILLE NON RENSEIGNÉE"
        lignes.append(("ville", [nom_ville] + [""] * (NB_COLS - 1)))
        total = 0.0
        for _, row in groupe.iterrows():
            total += extraire_nombre(row["poids"])
            lignes.append(("colis", [
                row["nom_client"], row["pays"], row["ville"], row["numero_telephone"],
                row["poids"], row["quantite"], row["contenu"], row["prix_unitaire"],
                row["infos_fournisseur"], formater_date(row["Date"]), row["montant"],
                row["analyse_agent"],
            ]))
        total_general += total
        lignes.append(("soustotal", [f"SOUS-TOTAL {nom_ville}", "", "", "", round(total, 2)]
                       + [""] * (NB_COLS - 5)))

    lignes.append(("vide", [""] * NB_COLS))
    lignes.append(("total_kg", ["TOTAL KG", "", "", "", round(total_general, 2)] + [""] * (NB_COLS - 5)))
    total_pieces = sum(nb_pieces(q) for q in df["quantite"])
    lignes.append(("total_colis", ["TOTAL COLIS", "", "", "", round(total_pieces, 2)] + [""] * (NB_COLS - 5)))
    return lignes


# ---------- Google Sheets ----------
@st.cache_resource
def get_gspread_client():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        scope = ["https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        return gspread.authorize(creds)
    except Exception:
        return None


def _c(t):
    return {"red": t[0], "green": t[1], "blue": t[2]}


def requetes_format(lignes, sid, total_rows):
    """Toute la mise en forme dans une seule liste de requêtes batch."""
    def rng(r1, r2, c1, c2):  # lignes 1-indexées, colonnes 0-indexées (c2 exclu)
        return {"sheetId": sid, "startRowIndex": r1 - 1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    def fmt(r1, r2, c1, c2, f):
        return {"repeatCell": {"range": rng(r1, r2, c1, c2),
                               "cell": {"userEnteredFormat": f}, "fields": "userEnteredFormat"}}

    def merge(r1, r2, c1, c2):
        return {"mergeCells": {"range": rng(r1, r2, c1, c2), "mergeType": "MERGE_ALL"}}

    gris = {"style": "SOLID", "color": _c(GRIS)}
    noir = {"style": "SOLID", "color": _c(NOIR)}
    cadre = {"top": gris, "bottom": gris, "left": gris, "right": gris}
    kg = {"type": "NUMBER", "pattern": FORMAT_KG}

    reqs = [
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": total_rows},
            "properties": {"pixelSize": 21}, "fields": "pixelSize"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 2}},
            "fields": "gridProperties.frozenRowCount"}},
    ]

    for i, (t, _) in enumerate(lignes):
        n = i + 1
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": HAUTEURS[t]}, "fields": "pixelSize"}})

        if t == "titre":
            reqs.append(merge(n, n, 0, NB_COLS))
            reqs.append(fmt(n, n, 0, NB_COLS, {
                "textFormat": {"bold": True, "fontSize": 14, "foregroundColor": _c(BLANC)},
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                "backgroundColor": _c(BLEU_FONCE)}))
        elif t == "entete":
            reqs.append(fmt(n, n, 0, NB_COLS, {
                "textFormat": {"bold": True}, "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP",
                "backgroundColor": _c(BLEU_CLAIR), "borders": {"bottom": noir}}))
        elif t == "ville":
            reqs.append(merge(n, n, 0, NB_COLS))
            reqs.append(fmt(n, n, 0, NB_COLS, {
                "textFormat": {"bold": True, "fontSize": 12},
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                "backgroundColor": _c(JAUNE)}))
        elif t == "colis":
            reqs.append(fmt(n, n, 0, NB_COLS, {
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                "wrapStrategy": "WRAP", "borders": cadre}))
        elif t == "soustotal":
            reqs.append(merge(n, n, 0, 4))
            reqs.append(fmt(n, n, 0, 4, {
                "textFormat": {"bold": True}, "horizontalAlignment": "RIGHT",
                "verticalAlignment": "MIDDLE", "backgroundColor": _c(ORANGE_CLAIR)}))
            reqs.append(fmt(n, n, 4, 5, {
                "textFormat": {"bold": True, "fontSize": 11}, "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE", "backgroundColor": _c(ORANGE), "numberFormat": kg}))
        elif t in ("total_kg", "total_colis"):
            taille = 14 if t == "total_kg" else 12
            base = {"textFormat": {"bold": True, "fontSize": taille},
                    "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                    "backgroundColor": _c(ORANGE), "borders": {"top": noir, "bottom": noir}}
            reqs.append(merge(n, n, 0, 4))
            reqs.append(fmt(n, n, 0, 4, base))
            reqs.append(fmt(n, n, 4, 5, dict(base, numberFormat=
                            kg if t == "total_kg" else {"type": "NUMBER", "pattern": "0"})))

    for j, w in enumerate(LARGEURS):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": j, "endIndex": j + 1},
            "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    return reqs


def envoyer_et_formater_sheet(client, df, nom_onglet, jour=None):
    try:
        doc = client.open_by_key(SHEET_ID)
        try:
            ws = doc.worksheet(nom_onglet)
        except gspread.exceptions.WorksheetNotFound:
            ws = doc.add_worksheet(title=nom_onglet, rows=200, cols=NB_COLS)

        lignes = construire_layout(df, jour)
        total_rows = max(len(lignes) + 30, 100)

        ws.clear()
        doc.batch_update({"requests": [
            {"unmergeCells": {"range": {"sheetId": ws.id}}},
            {"repeatCell": {"range": {"sheetId": ws.id}, "cell": {"userEnteredFormat": {}},
                            "fields": "userEnteredFormat"}},
        ]})
        ws.resize(rows=total_rows, cols=NB_COLS)
        ws.update(values=[v for _, v in lignes], range_name="A1")
        doc.batch_update({"requests": requetes_format(lignes, ws.id, total_rows)})
        return True
    except Exception as e:
        return str(e)


# ---------- Excel ----------
def _hex(t):
    return "%02X%02X%02X" % tuple(int(round(x * 255)) for x in t)


def construire_excel(lignes):
    wb = Workbook()
    ws = wb.active
    ws.title = "Suivi_Colis"

    def fond(t):
        return PatternFill("solid", start_color=_hex(t), end_color=_hex(t))

    gris = Side(style="thin", color=_hex(GRIS))
    noir = Side(style="thin", color=_hex(NOIR))
    centre = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, (t, vals) in enumerate(lignes, start=1):
        for j, v in enumerate(vals, start=1):
            ws.cell(row=i, column=j, value=v)
        ws.row_dimensions[i].height = HAUTEURS[t] * 0.75  # pixels -> points
        cellules = [ws.cell(row=i, column=j) for j in range(1, NB_COLS + 1)]

        if t == "titre":
            ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=NB_COLS)
            for c in cellules:
                c.fill, c.alignment = fond(BLEU_FONCE), centre
                c.font = Font(bold=True, size=14, color=_hex(BLANC))
        elif t == "entete":
            for c in cellules:
                c.fill, c.alignment, c.font = fond(BLEU_CLAIR), centre, Font(bold=True)
                c.border = Border(bottom=noir)
        elif t == "ville":
            ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=NB_COLS)
            for c in cellules:
                c.fill, c.alignment, c.font = fond(JAUNE), centre, Font(bold=True, size=12)
        elif t == "colis":
            for c in cellules:
                c.alignment = centre
                c.border = Border(top=gris, bottom=gris, left=gris, right=gris)
        elif t == "soustotal":
            ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4)
            for c in cellules[:4]:
                c.fill, c.font = fond(ORANGE_CLAIR), Font(bold=True)
                c.alignment = Alignment(horizontal="right", vertical="center")
            e = cellules[4]
            e.fill, e.font, e.alignment = fond(ORANGE), Font(bold=True, size=11), centre
            e.number_format = FORMAT_KG
        elif t in ("total_kg", "total_colis"):
            ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4)
            for c in cellules[:5]:
                c.fill, c.alignment = fond(ORANGE), centre
                c.font = Font(bold=True, size=14 if t == "total_kg" else 12)
                c.border = Border(top=noir, bottom=noir)
            cellules[4].number_format = FORMAT_KG if t == "total_kg" else "0"

    for j, w in enumerate(LARGEURS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w / 7
    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =====================================================
# 5. GEMINI
# =====================================================
client_gspread = get_gspread_client()

client_ia = None
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    try:
        http_options = types.HttpOptions(
            timeout=TIMEOUT_MS,
            retry_options=types.HttpRetryOptions(attempts=1),
        )
    except (AttributeError, TypeError):
        http_options = types.HttpOptions(timeout=TIMEOUT_MS)
    client_ia = genai.Client(api_key=api_key, http_options=http_options)
except Exception:
    client_ia = None


def preparer_image(data):
    """Redimensionne et compresse en JPEG (léger = envoi rapide). `data` = octets de la photo."""
    image = Image.open(io.BytesIO(data))
    image = ImageOps.exif_transpose(image).convert("RGB")
    image.thumbnail(TAILLE_IMAGE, Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    image.save(buf, format="JPEG", quality=80)
    return buf.getvalue()


def _config(niveau):
    kwargs = dict(response_mime_type="application/json", temperature=0, max_output_tokens=800)
    if niveau:
        kwargs["thinking_config"] = types.ThinkingConfig(thinking_level=niveau)
    return types.GenerateContentConfig(**kwargs)


def _nettoyer_json(texte):
    texte = (texte or "").strip()
    texte = re.sub(r"^```(?:json)?\s*|\s*```$", "", texte)
    donnees = json.loads(texte)
    if isinstance(donnees, list):
        donnees = donnees[0] if donnees else {}
    return donnees


def _raison(msg):
    if "RESOURCE_EXHAUSTED" in msg or msg.startswith("429"):
        return "quota épuisé"
    if "NOT_FOUND" in msg or msg.startswith("404"):
        return "modèle indisponible"
    if "UNAVAILABLE" in msg or msg.startswith("503"):
        return "modèle surchargé"
    if "timed out" in msg.lower() or "timeout" in msg.lower():
        return "délai dépassé"
    return msg[:300]


def analyser_une_image(img_bytes, client, modeles_txt):
    if client is None:
        return {"erreur": "Clé Gemini manquante ou invalide dans secrets.toml", "nom_client": "Erreur"}

    part = types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg")
    modeles = [m.strip() for m in modeles_txt.split(",") if m.strip()]
    journal = []

    for modele in modeles:                       # secours : modèle suivant si quota/indisponible
        raison = ""
        for niveau in (THINKING_LEVEL, None):    # réglage refusé (400) : on retente sans
            try:
                response = client.models.generate_content(
                    model=modele, contents=[PROMPT, part], config=_config(niveau))
                donnees = _nettoyer_json(response.text)
                resultat = {k: str(donnees.get(k, "") or "").strip() for k in CLES}
                resultat["analyse_agent"] = f"Extraction ({modele})"
                return resultat
            except Exception as e:
                msg = str(e)
                raison = _raison(msg)
                if "INVALID_ARGUMENT" in msg or msg.startswith("400"):
                    continue
                break
        journal.append(f"{modele} : {raison}")

    return {"erreur": "\n".join(journal) or "Aucun modèle configuré", "nom_client": "Erreur"}


def detecter_modeles(client):
    """Liste les modèles 'flash' utilisables avec ta clé (generateContent)."""
    exclus = ("image", "tts", "live", "audio", "embedding", "robotics", "computer", "thinking-exp")
    noms = []
    for m in client.models.list():
        nom = (m.name or "").replace("models/", "")
        actions = getattr(m, "supported_actions", None) or []
        if "generateContent" in actions and "flash" in nom and not any(x in nom for x in exclus):
            noms.append(nom)
    noms = sorted(set(noms), reverse=True)
    if MODELE in noms:
        noms.remove(MODELE)
        noms.insert(0, MODELE)
    return noms




def traiter_photos(photos, client, modeles, defauts, historique, fusion_auto, progress=None,
                   hashes_extra=None):
    """Analyse des photos (octets) puis ajout/regroupement dans l'historique.
    Retourne (nouvel_historique, resume)."""
    connus = hashes_connus(historique) | set(hashes_extra or ())
    a_analyser, ignorees, vus = [], 0, set(connus)
    for data in photos:
        h = hashlib.md5(data).hexdigest()[:12]
        if h in vus:
            ignorees += 1          # photo déjà analysée : on n'utilise pas de quota
            continue
        vus.add(h)
        a_analyser.append((h, data))

    resume = {"total": len(a_analyser), "ok": 0, "erreurs": [], "regroupes": 0,
              "ignorees": ignorees, "ids_recents": []}
    if not a_analyser:
        return historique, resume

    images = [preparer_image(d) for _, d in a_analyser]
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(analyser_une_image, b, client, modeles) for b in images]
        for i, _ in enumerate(as_completed(futures)):
            if progress:
                progress((i + 1) / len(futures))
        resultats = [f.result() for f in futures]  # ordre conservé

    nouveaux, nouveaux_hashes = [], set()
    for res, (h, _) in zip(resultats, a_analyser):
        if "erreur" in res:
            resume["erreurs"].append(res["erreur"])
            continue
        for k, v in defauts.items():
            if v and not res.get(k):
                res[k] = v
        res["hash_photos"] = h
        nouveaux.append(res)
        nouveaux_hashes.add(h)
    resume["ok"] = len(nouveaux)

    ids_avant = {r["ID"] for r in historique}
    tous = normaliser(list(historique) + nouveaux)
    ids_nouveaux = {r["ID"] for r in tous if r["ID"] not in ids_avant}
    final = regrouper(tous, ids_nouveaux) if fusion_auto else tous
    resume["regroupes"] = len(tous) - len(final)
    resume["ids_recents"] = [r["ID"] for r in final
                             if nouveaux_hashes & set(r["hash_photos"].split(","))]
    return final, resume


# =====================================================
# 6. COMPOSANTS D'INTERFACE (pensés pour le téléphone)
# =====================================================
CSS = """
<style>
/* --- structure générale --- */
#MainMenu, footer {visibility: hidden;}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"] {display: none !important;}
.block-container, [data-testid="stMainBlockContainer"] {
  padding: 0.8rem 0.9rem 5rem 0.9rem !important; max-width: 920px;}
h3 {font-size: 1.15rem !important;}
h4 {font-size: 1.05rem !important; margin: 0 0 0.4rem 0 !important; padding: 0 !important;}

/* --- bandeau --- */
.entete {background: linear-gradient(135deg, #1a3a8f 0%, #2f6bff 100%); color: #fff;
  border-radius: 16px; padding: 1rem 1.1rem; margin-bottom: 0.9rem;
  box-shadow: 0 4px 14px rgba(26,58,143,0.30);}
.entete-titre {font-size: 1.5rem; font-weight: 800; line-height: 1.2; color: #fff;}
.entete-sous {opacity: 0.88; font-size: 0.9rem; margin-top: 0.15rem; color: #fff;}

/* --- onglets en pastilles (ancienne et nouvelle structure Streamlit) --- */
[data-testid="stTabs"] [role="tablist"], .stTabs [data-baseweb="tab-list"] {
  display: flex; gap: 0.25rem; padding: 0.25rem; border-radius: 14px;
  background: rgba(128,128,128,0.16); border-bottom: none !important; box-shadow: none !important;}
[data-testid="stTabs"] [role="tab"], .stTabs [data-baseweb="tab"] {
  flex: 1 1 0; justify-content: center; text-align: center; border-radius: 10px;
  padding: 0.55rem 0.2rem; margin: 0; height: auto; font-weight: 600;}
[data-testid="stTabs"] [role="tab"] p, .stTabs [data-baseweb="tab"] p {margin: 0; font-size: 1rem;}
[data-testid="stTabs"] [role="tab"][aria-selected="true"], .stTabs [aria-selected="true"] {
  background: #2f6bff !important; color: #fff !important;}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] p {color: #fff !important;}
.react-aria-SelectionIndicator, [data-baseweb="tab-highlight"], [data-baseweb="tab-border"] {
  display: none !important;}
[data-testid="stTabs"] > div:first-child {border-bottom: none !important;}

/* --- boutons et champs --- */
.stButton button, .stDownloadButton button, [data-testid="stFormSubmitButton"] button {
  min-height: 3.1rem; font-size: 1.05rem; font-weight: 600; border-radius: 12px;}
button[kind="primary"], [data-testid="stBaseButton-primary"],
[data-testid="stBaseButton-primaryFormSubmit"] {
  background: #2f6bff !important; border-color: #2f6bff !important; color: #fff !important;}
button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover {filter: brightness(1.12);}
button[kind="primary"] p, [data-testid="stBaseButton-primary"] p {color: #fff !important;}
.stTextInput input {font-size: 16px; min-height: 2.8rem;}

/* --- titres d'étapes --- */
.etape {display: flex; align-items: center; gap: 0.6rem; font-weight: 700; font-size: 1.1rem;
  margin: 1.1rem 0 0.5rem 0;}
.etape span {display: inline-flex; align-items: center; justify-content: center;
  min-width: 1.9rem; height: 1.9rem; padding: 0 0.35rem; border-radius: 999px;
  background: #2f6bff; color: #fff; font-size: 0.95rem; font-weight: 800;}

/* --- état vide --- */
.vide {text-align: center; padding: 1.6rem 1rem; margin: 0.8rem 0; border-radius: 14px;
  border: 2px dashed rgba(128,128,128,0.45);}
.vide .grand {font-size: 2.4rem; margin-bottom: 0.3rem;}
.vide b {display: block; font-size: 1.05rem;}
.vide small {display: block; opacity: 0.75; margin-top: 0.2rem; font-size: 0.9rem;}

/* --- tuiles de chiffres --- */
.barre {display: flex; gap: 0.6rem; margin: 0.2rem 0 0.8rem 0;}
.barre div {flex: 1; text-align: center; padding: 0.6rem 0.2rem; border-radius: 12px;
  background: rgba(128,128,128,0.14); border-top: 3px solid #2f6bff;
  font-size: 0.85rem; line-height: 1.2;}
.barre b {display: block; font-size: 1.4rem;}

/* --- destinations et cartes --- */
.groupe {display: flex; justify-content: space-between; align-items: center; gap: 0.5rem;
  background: linear-gradient(90deg, #f5c400, #ffdc4d); color: #222; font-weight: 800;
  border-radius: 10px; padding: 0.5rem 0.8rem; margin: 1rem 0 0.5rem 0; font-size: 1rem;}
.groupe-info {font-weight: 600; font-size: 0.8rem; text-align: right;}
.carte {display: flex; gap: 0.75rem; align-items: flex-start;
  border: 1px solid rgba(128,128,128,0.35); border-left: 5px solid #2e9e5b;
  border-radius: 12px; padding: 0.7rem 0.85rem; margin-bottom: 0.55rem;
  background: rgba(128,128,128,0.10);}
.carte.manque {border-left-color: #ff9800;}
.avatar {flex: 0 0 2.6rem; width: 2.6rem; height: 2.6rem; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  background: #2f6bff; color: #fff; font-weight: 800; font-size: 1.15rem;}
.carte.manque .avatar {background: #ff9800; color: #111;}
.carte-corps {flex: 1; min-width: 0;}
.carte-tete {display: flex; justify-content: space-between; align-items: flex-start; gap: 0.5rem;
  margin-bottom: 0.3rem;}
.carte-nom {font-size: 1.05rem; font-weight: 700; line-height: 1.2; word-break: break-word;}
.carte-id {font-size: 0.78rem; opacity: 0.6;}
.badge {background: #ff9800; color: #111; border-radius: 999px; padding: 0.05rem 0.6rem;
  font-size: 0.85rem; font-weight: 700; white-space: nowrap;}
.puces {display: flex; flex-wrap: wrap; gap: 0.35rem; margin-top: 0.4rem;}
.puce {background: rgba(128,128,128,0.20); border-radius: 999px; padding: 0.15rem 0.6rem;
  font-size: 0.82rem; word-break: break-word;}
.carte-ligne {font-size: 0.95rem; margin: 0.1rem 0; word-break: break-word;}
.carte-ligne.petit {font-size: 0.85rem; opacity: 0.85;}
.carte-alerte {font-size: 0.85rem; color: #ff9800; margin-top: 0.3rem; font-weight: 600;}
</style>
"""

CHAMPS_EDIT = [
    ("nom_client", "Nom du client"),
    ("numero_telephone", "Téléphone"),
    ("ville", "Ville de destination"),
    ("pays", "Pays"),
    ("poids", "Poids (kg)"),
    ("quantite", "Nombre de colis"),
    ("contenu", "Contenu"),
    ("infos_fournisseur", "Fournisseur / usine"),
    ("prix_unitaire", "Prix unitaire"),
    ("montant", "Montant"),
]
CHAMPS_DEFAUT = [(k, l) for k, l in CHAMPS_EDIT if k not in ("poids", "montant")]


def e(x):
    return html.escape(str(x))


def flash(type_, texte):
    st.session_state.flash = (type_, texte)


def afficher_flash():
    f = st.session_state.pop("flash", None)
    if f:
        getattr(st, f[0])(f[1])


def etape(numero, titre):
    st.markdown(f'<div class="etape"><span>{numero}</span>{e(titre)}</div>', unsafe_allow_html=True)


def vide_html(emoji, titre, sous_titre):
    st.markdown(f'<div class="vide"><div class="grand">{emoji}</div><b>{e(titre)}</b>'
                f'<small>{e(sous_titre)}</small></div>', unsafe_allow_html=True)


def chiffres_cles(recs):
    pieces = sum(nb_pieces(r["quantite"]) for r in recs)
    kg = sum(extraire_nombre(r["poids"]) for r in recs)
    villes = len({r["ville"].strip().casefold() for r in recs if r["ville"].strip()})
    return pieces, kg, villes


def barre_html(recs):
    pieces, kg, villes = chiffres_cles(recs)
    return ('<div class="barre">'
            f'<div><b>{_fmt_nombre(round(pieces, 2))}</b>colis</div>'
            f'<div><b>{_fmt_nombre(round(kg, 1))}</b>kg</div>'
            f'<div><b>{villes}</b>ville(s)</div></div>')


def carte_html(rec):
    manques = [lbl for k, lbl in (("numero_telephone", "téléphone"), ("ville", "ville"), ("pays", "pays"))
               if not rec.get(k)]
    nom = rec["nom_client"] or "Sans nom"
    initiale = e(next((c for c in nom if c.isalnum()), "?").upper())
    qte = rec.get("quantite", "")
    badge = f'<span class="badge">× {e(qte)}</span>' if qte else ""
    lieu = ", ".join(x for x in (rec["ville"], rec["pays"]) if x) or "—"
    puces = []
    if rec["poids"]:
        poids = rec["poids"] if re.search(r"[a-zA-Z]", rec["poids"]) else rec["poids"] + " kg"
        puces.append(f"⚖️ {e(poids)}")
    if rec["contenu"]:
        puces.append(f"📦 {e(rec['contenu'])}")
    if rec["infos_fournisseur"]:
        puces.append(f"🏭 {e(rec['infos_fournisseur'])}")
    h = f'<div class="carte{" manque" if manques else ""}"><div class="avatar">{initiale}</div><div class="carte-corps">'
    h += f'<div class="carte-tete"><div><div class="carte-nom">{e(nom)}</div>'
    h += f'<div class="carte-id">#{rec["ID"]}</div></div>{badge}</div>'
    h += f'<div class="carte-ligne">📍 {e(lieu)}</div>'
    h += f'<div class="carte-ligne">📞 {e(rec["numero_telephone"] or "—")}</div>'
    if puces:
        h += '<div class="puces">' + "".join(f'<span class="puce">{p}</span>' for p in puces) + "</div>"
    if manques:
        h += f'<div class="carte-alerte">À compléter : {e(", ".join(manques))}</div>'
    return h + "</div></div>"


def cartes_html(recs):
    return "".join(carte_html(r) for r in reversed(recs))  # les plus récents d'abord


def cartes_groupees_html(recs):
    """Cartes regroupées par destination, comme dans le Google Sheet."""
    groupes = {}
    for r in recs:
        groupes.setdefault((r["pays"].strip(), r["ville"].strip()), []).append(r)
    h = ""
    for (pays, ville), g in sorted(groupes.items(), key=lambda kv: (kv[0][0].casefold(), kv[0][1].casefold())):
        pieces = sum(nb_pieces(r["quantite"]) for r in g)
        kg = sum(extraire_nombre(r["poids"]) for r in g)
        infos = [x for x in (pays, f"{_fmt_nombre(round(pieces, 2))} colis",
                             f"{_fmt_nombre(round(kg, 1))} kg" if kg else "") if x]
        h += (f'<div class="groupe"><span>{e((ville or "Ville non renseignée").upper())}</span>'
              f'<span class="groupe-info">{e(" · ".join(infos))}</span></div>')
        h += "".join(carte_html(r) for r in reversed(g))
    return h


def filtrer(recs, recherche):
    q = (recherche or "").strip().casefold()
    if not q:
        return recs
    return [r for r in recs if q in " ".join(
        str(r[k]) for k in ("nom_client", "numero_telephone", "ville", "pays", "contenu",
                            "infos_fournisseur")).casefold()]


def libelle(rec):
    return f"#{rec['ID']} · {rec['nom_client'] or 'Sans nom'} · {rec['ville'] or '—'}"


def signature(rec):
    return hashlib.md5(json.dumps(rec, sort_keys=True, default=str).encode()).hexdigest()[:6]


def formulaire_colis(rec, prefixe):
    vals = {}
    for cle, label in CHAMPS_EDIT:
        vals[cle] = st.text_input(label, value=rec.get(cle, ""), key=f"{prefixe}_{cle}").strip()
    return vals


def section_edition(recs):
    etape("✏️", "Modifier un colis")
    par_id = {r["ID"]: r for r in recs}
    if st.session_state.get("colis_sel") not in par_id:
        st.session_state.colis_sel = list(par_id)[-1]
    choix = st.selectbox("Colis à modifier", list(par_id), key="colis_sel",
                         format_func=lambda i: libelle(par_id[i]))
    rec = par_id[choix]

    with st.form(f"form_edit_{choix}"):
        vals = formulaire_colis(rec, f"e{choix}_{signature(rec)}")
        c1, c2 = st.columns(2)
        enregistrer = c1.form_submit_button("💾 Enregistrer", type="primary", **LARGEUR)
        supprimer = c2.form_submit_button("🗑 Supprimer", **LARGEUR)

    if enregistrer:
        maj = [dict(r, **vals) if r["ID"] == choix else r for r in recs]
        st.session_state.historique_colis = normaliser(maj)
        sauvegarder()
        flash("success", f"Colis #{choix} enregistré.")
        st.rerun()
    if supprimer:
        st.session_state.suppr_id = choix
    if st.session_state.get("suppr_id") != choix:
        st.session_state.pop("suppr_id", None)
    else:
        st.warning(f"Supprimer le colis #{choix} ({rec['nom_client'] or 'Sans nom'}) ?")
        a, b = st.columns(2)
        if a.button("Oui, supprimer", type="primary", key="suppr_oui", **LARGEUR):
            st.session_state.historique_colis = [r for r in recs if r["ID"] != choix]
            st.session_state.pop("suppr_id", None)
            sauvegarder()
            flash("success", f"Colis #{choix} supprimé.")
            st.rerun()
        if b.button("Annuler", key="suppr_non", **LARGEUR):
            st.session_state.pop("suppr_id", None)
            st.rerun()


def section_ajout_manuel(recs):
    with st.expander("➕ Ajouter un colis à la main"):
        n = st.session_state.cle_ajout
        with st.form(f"form_ajout_{n}"):
            vals = formulaire_colis({}, f"ajout{n}")
            ajouter = st.form_submit_button("Ajouter ce colis", type="primary", **LARGEUR)
        if ajouter:
            if not vals["nom_client"]:
                st.error("Le nom du client est obligatoire.")
            else:
                nouveau = dict(vals, analyse_agent="Saisie manuelle")
                ids_avant = {r["ID"] for r in recs}
                tous = normaliser(list(recs) + [nouveau])
                ids_nouveaux = {r["ID"] for r in tous if r["ID"] not in ids_avant}
                final = regrouper(tous, ids_nouveaux) if st.session_state.fusion_auto else tous
                st.session_state.historique_colis = final
                st.session_state.cle_ajout += 1
                sauvegarder()
                flash("success", "Colis ajouté" + (" et regroupé avec un colis du même client."
                                                   if len(final) < len(tous) else "."))
                st.rerun()


def vue_tableau(recs):
    df = pd.DataFrame(recs)[COLONNES_ORDRE]
    df_affichage = df.rename(columns={
        "nom_client": "Name", "pays": "Country", "ville": "City", "numero_telephone": "Phone",
        "poids": "Weight", "quantite": "Nber Of Box pieces", "contenu": "Contain",
        "prix_unitaire": "Unit Price", "infos_fournisseur": "Factory Name",
        "montant": "Montant", "analyse_agent": "Analyse IA",
    })
    st.caption("Touchez une case pour corriger. Pour supprimer : cochez le début de la ligne puis "
               "icône corbeille. Pour ajouter : dernière ligne du tableau.")
    df_modifie = st.data_editor(
        df_affichage, hide_index=True, num_rows="dynamic", **LARGEUR,
        column_config={
            "hash_photos": None,
            "ID": st.column_config.NumberColumn("ID", disabled=True),
            "Date": st.column_config.TextColumn("Date", disabled=True),
            "Analyse IA": st.column_config.TextColumn("Analyse IA", disabled=True),
            **{c: st.column_config.TextColumn(c) for c in [
                "Name", "Country", "City", "Phone", "Weight", "Nber Of Box pieces",
                "Contain", "Unit Price", "Factory Name", "Montant"]},
        },
    )
    df_technique = df_modifie.rename(columns={
        "Name": "nom_client", "Country": "pays", "City": "ville", "Phone": "numero_telephone",
        "Weight": "poids", "Nber Of Box pieces": "quantite", "Contain": "contenu",
        "Unit Price": "prix_unitaire", "Factory Name": "infos_fournisseur",
        "Montant": "montant", "Analyse IA": "analyse_agent",
    })
    nouvel_etat = normaliser(df_technique.to_dict("records"))
    if nouvel_etat != st.session_state.historique_colis:
        st.session_state.historique_colis = nouvel_etat
        sauvegarder()


def bloc_stats(recs):
    df = pd.DataFrame(recs)
    df["kg"] = df["poids"].apply(extraire_nombre)
    df["pieces"] = df["quantite"].apply(nb_pieces)
    df["ville"] = df["ville"].replace("", "Non renseignée")
    df["pays"] = df["pays"].replace("", "Non renseigné")
    resume = (df.groupby(["pays", "ville"]).agg(Colis=("pieces", "sum"), Kg=("kg", "sum"))
              .reset_index().rename(columns={"pays": "Pays", "ville": "Ville"}))
    resume["Kg"] = resume["Kg"].round(1)
    st.write("**Détail par destination**")
    st.dataframe(resume, hide_index=True)
    st.write("**Colis par ville**")
    st.bar_chart(df.groupby("ville")["pieces"].sum())
    st.write("**Colis par pays**")
    st.bar_chart(df.groupby("pays")["pieces"].sum())


def nom_onglet_pour(jour):
    """Un onglet Google Sheets par jour (ou le nom choisi dans Réglages)."""
    if st.session_state.get("onglet_auto", True):
        try:
            return "Shipping " + datetime.strptime(jour, "%Y-%m-%d").strftime("%d-%m-%Y")
        except Exception:
            return ONGLET_DEFAUT
    return (st.session_state.get("nom_onglet", "") or "").strip() or ONGLET_DEFAUT


def etiquette_jour(jour):
    n = sum(nb_pieces(r["quantite"]) for r in st.session_state.jours.get(jour, []))
    prefixe = "Aujourd'hui · " if jour == aujourdhui() else ""
    return f"{prefixe}{jour_lisible(jour)} — " + (f"{_fmt_nombre(round(n, 2))} colis" if n else "vide")


def options_jours():
    synchroniser_jour()
    jours = {j for j, l in st.session_state.jours.items() if l}
    jours |= {aujourdhui(), st.session_state.jour_charge}
    return sorted(jours, reverse=True)


def bloc_historique():
    synchroniser_jour()
    lignes = []
    for j in sorted(st.session_state.jours, reverse=True):
        liste = st.session_state.jours[j]
        if liste:
            lignes.append({"Dossier": jour_lisible(j), "Clients": len(liste),
                           "Colis": round(sum(nb_pieces(r["quantite"]) for r in liste), 2),
                           "Kg": round(sum(extraire_nombre(r["poids"]) for r in liste), 1)})
    with st.expander("📚 Tous les dossiers (historique)"):
        if lignes:
            st.dataframe(pd.DataFrame(lignes), hide_index=True)
        else:
            st.info("Aucun dossier enregistré pour le moment.")
        st.caption("Pour ouvrir un autre jour, utilise le sélecteur « 📅 Dossier du jour » en haut de la page.")


def bloc_export(recs):
    """Boutons Excel / Google Sheets / CSV du dossier affiché (grisés s'il est vide)."""
    vide = not recs
    jour = st.session_state.jour_charge
    nom_onglet = nom_onglet_pour(jour)
    with st.container(border=True):
        st.markdown(f"#### 📤 Exporter le dossier du {jour_lisible(jour)}")
        c1, c2 = st.columns(2)
        df_final = pd.DataFrame(recs) if recs else None

        with c1:
            if vide or not OPENPYXL_OK:
                st.button("📗 Excel (.xlsx)", disabled=True, key="btn_excel_off", **LARGEUR,
                          help="Ajoute d'abord des colis" if vide else "Installe openpyxl : pip install openpyxl")
            else:
                try:
                    st.download_button(
                        "📗 Excel (.xlsx)", data=construire_excel(construire_layout(df_final, jour)),
                        file_name=f"shipping_{jour}.xlsx", key="dl_excel", type="primary",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **LARGEUR)
                except Exception as ex:
                    st.error(f"Export Excel impossible : {ex}")
        with c2:
            envoyer = st.button("📤 Google Sheets", type="primary", disabled=vide, key="btn_sheets",
                                help="Ajoute d'abord des colis" if vide else None, **LARGEUR)

        if envoyer:
            if client_gspread:
                with st.spinner("Mise en forme et envoi en cours..."):
                    result = envoyer_et_formater_sheet(client_gspread, df_final, nom_onglet, jour)
                if result is True:
                    st.success(f"Dossier du {jour_lisible(jour)} envoyé dans l'onglet « {nom_onglet} ».")
                else:
                    st.error(f"Erreur Google Sheets : {result}")
            else:
                st.error("Google Sheets n'est pas connecté : vérifie `gcp_service_account` dans "
                         "`.streamlit/secrets.toml` (voir l'onglet ⚙️ Réglages).")

        st.caption(f"Google Sheets : onglet « {nom_onglet} ». Les autres jours ne sont pas touchés.")
        if not vide:
            csv = df_final.drop(columns=["hash_photos"], errors="ignore").to_csv(index=False).encode("utf-8-sig")
            st.download_button("📄 Télécharger en CSV", data=csv, file_name=f"colis_{jour}.csv",
                               mime="text/csv", key="dl_csv", **LARGEUR)


# =====================================================
# 7. INTERFACE PRINCIPALE
# =====================================================
st.markdown(CSS, unsafe_allow_html=True)
st.markdown('<div class="entete"><div class="entete-titre">📦 Agent Colis</div>'
            '<div class="entete-sous">Scanner · Vérifier · Exporter</div></div>', unsafe_allow_html=True)

st.selectbox("📅 Dossier du jour", options_jours(), key="jour_actif", format_func=etiquette_jour)
if st.session_state.jour_actif != st.session_state.jour_charge:
    synchroniser_jour()                                   # on garde le dossier qu'on quitte
    st.session_state.jour_charge = st.session_state.jour_actif
    st.session_state.historique_colis = st.session_state.jours.get(st.session_state.jour_charge, [])
    for _k in ("colis_sel", "suppr_id", "recherche"):
        st.session_state.pop(_k, None)
if st.session_state.jour_charge != aujourdhui():
    st.warning(f"Tu consultes un ancien dossier : {jour_lisible(st.session_state.jour_charge)}. "
               "Les colis ajoutés iront dans CE dossier.")
    st.button("↩ Revenir à aujourd'hui", key="btn_aujourdhui", **LARGEUR,
              on_click=lambda: st.session_state.update(jour_actif=aujourdhui()))

tab_scan, tab_colis, tab_reglages = st.tabs(["📷 Scanner", "📋 Colis", "⚙️ Réglages"])

# ---------------- SCANNER ----------------
with tab_scan:
    nb_restaures = st.session_state.pop("restaure", 0)
    if nb_restaures:
        st.info(f"{nb_restaures} colis restaurés depuis la dernière session.")
    st.caption(f"📅 Les colis seront ajoutés au dossier du {jour_lisible(st.session_state.jour_charge)}.")

    etape(1, "Ajouter des photos")
    methode = st.radio("Comment ajouter ?", ["📁 Galerie / photo", "📷 Caméra"], horizontal=True,
                       label_visibility="collapsed")
    cle = st.session_state.cle_upload

    fichiers = []
    if methode.startswith("📷"):
        img = st.camera_input("Prenez la photo du colis", key=f"camera_{cle}")
        if img:
            fichiers = [img]
    else:
        fichiers = st.file_uploader(
            "Choisissez une ou plusieurs photos (sur téléphone : appareil photo ou galerie)",
            type=["jpg", "jpeg", "png"], accept_multiple_files=True, key=f"upload_{cle}") or []

    if fichiers:
        st.image([f.getvalue() for f in fichiers[:12]], width=80)

    with st.expander("Remplir automatiquement les infos manquantes (optionnel)"):
        st.caption("Utilisé seulement quand la photo ne contient pas l'information. "
                   "Laisse vide pour ne rien remplir (ex. nom de l'usine).")
        for k, label in CHAMPS_DEFAUT:
            st.text_input(label, key=f"def_{k}")

    if fichiers:
        etape(2, "Lancer l'analyse")
        if st.button(f"📷 ANALYSER {len(fichiers)} PHOTO(S)", type="primary", **LARGEUR):
            debut = time.time()
            barre = st.progress(0, text="Analyse en cours...")
            defauts = {k: st.session_state.get(f"def_{k}", "").strip() for k, _ in CHAMPS_DEFAUT}
            try:
                nouvel_historique, resume = traiter_photos(
                    [f.getvalue() for f in fichiers], client_ia, st.session_state.modeles_txt,
                    defauts, st.session_state.historique_colis, st.session_state.fusion_auto,
                    hashes_extra=tous_les_hashes(), progress=lambda p: barre.progress(p, text="Analyse en cours..."))
            except Exception as ex:
                st.error(f"Photo illisible ou erreur inattendue : {ex}")
                st.stop()
            st.session_state.historique_colis = nouvel_historique
            sauvegarder()
            resume["duree"] = time.time() - debut
            st.session_state.dernier_resume = resume
            if not resume["erreurs"]:
                st.session_state.cle_upload += 1  # vide la zone d'import
            st.rerun()

    r = st.session_state.pop("dernier_resume", None)
    if r:
        if r["total"]:
            msg = f"{r['ok']}/{r['total']} photo(s) analysée(s) en {r['duree']:.1f} s"
            (st.success if r["ok"] == r["total"] else st.warning)(msg)
        if r["regroupes"]:
            st.info(f"{r['regroupes']} colis regroupé(s) avec un colis du même client et de la même "
                    "destination : le nombre de colis a été mis à jour.")
        if r["ignorees"]:
            st.warning(f"{r['ignorees']} photo(s) déjà analysée(s) ignorée(s) (aucun quota utilisé).")
        for err in r["erreurs"]:
            st.error("Aucun modèle n'a pu analyser la photo :\n\n" + err.replace("\n", "\n\n"))
            if "quota épuisé" in err:
                st.info("Offre gratuite : 20 requêtes par jour et par modèle. Solutions : « Détecter mes "
                        "modèles disponibles » (onglet Réglages), attendre le renouvellement du quota, "
                        "ou activer la facturation sur ton projet Google Cloud.")
        if r["erreurs"]:
            st.caption("Les photos en échec sont conservées : relance l'analyse quand le problème est "
                       "réglé (les photos déjà réussies seront ignorées automatiquement).")
        recents = [x for x in st.session_state.historique_colis if x["ID"] in r["ids_recents"]]
        if recents:
            etape("✓", "Derniers colis ajoutés")
            st.markdown(cartes_html(recents), unsafe_allow_html=True)
            st.caption("Corrige ou complète dans l'onglet 📋 Colis.")

# ---------------- COLIS ----------------
with tab_colis:
    afficher_flash()
    recs = st.session_state.historique_colis
    if recs:
        st.markdown(barre_html(recs), unsafe_allow_html=True)
    bloc_export(recs)

    if not recs:
        vide_html("📦", "Aucun colis dans ce dossier", "Va dans l'onglet 📷 Scanner pour ajouter des photos.")
    else:
        vue = st.radio("Affichage", ["Cartes", "Tableau"], horizontal=True, key="vue",
                       label_visibility="collapsed")
        if vue == "Cartes":
            recherche = st.text_input("🔍 Rechercher (nom, téléphone, ville...)", key="recherche")
            trouves = filtrer(recs, recherche)
            if trouves:
                st.markdown(cartes_groupees_html(trouves), unsafe_allow_html=True)
            else:
                st.info("Aucun résultat.")
            if st.button("🔗 Regrouper les colis du même client", **LARGEUR):
                avant = len(recs)
                st.session_state.historique_colis = regrouper(recs)
                sauvegarder()
                diff = avant - len(st.session_state.historique_colis)
                flash("success" if diff else "info",
                      f"{diff} ligne(s) regroupée(s)." if diff else
                      "Rien à regrouper : aucun client n'apparaît deux fois pour la même destination.")
                st.rerun()
            section_edition(recs)
            section_ajout_manuel(recs)
        else:
            vue_tableau(recs)
        with st.expander("📊 Statistiques par destination"):
            bloc_stats(recs)
    bloc_historique()

# ---------------- RÉGLAGES ----------------
with tab_reglages:
    etape("🔌", "Connexions")
    st.write("Gemini : " + ("✅ connecté" if client_ia else "❌ clé manquante dans secrets.toml"))
    st.write("Google Sheets : " + ("✅ connecté" if client_gspread else
                                   "❌ non configuré (`gcp_service_account` dans secrets.toml)"))

    etape("📄", "Google Sheets")
    st.checkbox("Un onglet par jour (recommandé) : « Shipping jj-mm-aaaa »", key="onglet_auto")
    st.text_input("Nom d'onglet personnalisé", key="nom_onglet", disabled=st.session_state.onglet_auto)
    st.caption("Avec un onglet par jour, l'envoi d'aujourd'hui n'écrase jamais les jours précédents. "
               "Si tu renvoies le même jour, seul l'onglet de ce jour est remplacé.")

    etape("📦", "Colis")
    st.checkbox("Regrouper automatiquement les colis du même client et de la même destination",
                key="fusion_auto")
    st.checkbox("Sauvegarde automatique (reprise après rechargement de la page)", key="autosave")

    etape("🤖", "Modèles Gemini")
    if st.button("Détecter mes modèles disponibles", **LARGEUR):
        try:
            trouves = detecter_modeles(client_ia)
            if trouves:
                st.session_state.modeles_txt = ", ".join(trouves)
            else:
                st.warning("Aucun modèle flash trouvé pour cette clé.")
        except Exception as ex:
            st.error(f"Détection impossible : {ex}")
    st.text_input("Ordre d'essai (séparés par des virgules)", key="modeles_txt")
    st.caption("Si un modèle est épuisé (quota) ou introuvable, l'app passe automatiquement au suivant. "
               "Chaque modèle a son propre quota gratuit par jour.")

    etape("⚠️", "Zone de danger")
    with st.expander(f"🗑 Vider le dossier du {jour_lisible(st.session_state.jour_charge)}"):
        st.warning("Supprime les colis de CE jour uniquement. Les autres dossiers ne sont pas touchés. "
                   "Action irréversible.")
        confirmer = st.checkbox("Je confirme l'effacement de ce dossier")
        if st.button("Vider ce dossier", disabled=not confirmer, key="btn_effacer", **LARGEUR):
            st.session_state.historique_colis = []
            sauvegarder()
            flash("success", "Dossier vidé.")
            st.rerun()
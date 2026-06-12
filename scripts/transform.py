"""
transform.py — Transformation ERP → Reflex CSV
------------------------------------------------
1. Télécharge le fichier Excel depuis Google Drive (URL dans $GDRIVE_URL)
   Gère les gros fichiers (page de confirmation Google Drive)
2. Filtre les lignes où col. A = "5322A" ou "5322D"
3. Extrait col. B (Code Article) et col. CD (index 82, Prix)
4. Génère output/Export_Articles.csv (UTF-8 BOM, séparateur ;)
"""

import sys
import os
import re
import requests
import openpyxl
from pathlib import Path

# ── Constantes ───────────────────────────────────────────────────────────────
FILTER_VALUES = {"5322A", "5322D"}
COL_A  = 1
COL_B  = 2
COL_CD = 82
OUTPUT_FILE = "output/Export_Articles.csv"
LOCAL_EXCEL = "input/source.xlsx"

os.makedirs("input",  exist_ok=True)
os.makedirs("output", exist_ok=True)

# ── 1. Téléchargement depuis Google Drive ────────────────────────────────────
gdrive_url = os.environ.get("GDRIVE_URL", "").strip()
if not gdrive_url:
    print("[ERREUR] Variable GDRIVE_URL non définie.", file=sys.stderr)
    sys.exit(1)

match = re.search(r"/d/([a-zA-Z0-9_-]+)", gdrive_url)
if not match:
    print(f"[ERREUR] ID introuvable dans l'URL : {gdrive_url}", file=sys.stderr)
    sys.exit(1)

file_id = match.group(1)
print(f"[INFO] Téléchargement depuis Google Drive (ID: {file_id})…")

session = requests.Session()

def download_file(file_id):
    """Télécharge un fichier Google Drive, gère la confirmation pour les gros fichiers."""

    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    response = session.get(url, stream=True, timeout=300)

    # Détecter la page de confirmation HTML (gros fichiers)
    content_type = response.headers.get("Content-Type", "")
    if "text/html" in content_type:
        print("[INFO] Page de confirmation détectée (gros fichier), extraction du token…")
        html = response.text

        # Méthode 1 : chercher le token "confirm" dans le formulaire HTML
        token = None
        m = re.search(r'name="confirm"\s+value="([^"]+)"', html)
        if m:
            token = m.group(1)
        
        # Méthode 2 : chercher uuid dans l'action du formulaire
        if not token:
            m = re.search(r'confirm=([a-zA-Z0-9_-]+)', html)
            if m:
                token = m.group(1)

        # Méthode 3 : nouvelle API Google Drive (2024+)
        if not token:
            m = re.search(r'"downloadUrl":"([^"]+)"', html)
            if m:
                direct_url = m.group(1).replace('\\u003d', '=').replace('\\u0026', '&')
                print(f"[INFO] URL directe trouvée dans la page.")
                return session.get(direct_url, stream=True, timeout=600)

        if token:
            url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm={token}"
            print(f"[INFO] Token de confirmation : {token}")
            return session.get(url, stream=True, timeout=600)

        # Méthode 4 : utiliser l'API Drive v3 avec le lien direct
        print("[INFO] Tentative via API Drive v3…")
        url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm=t&uuid={file_id}"
        return session.get(url, stream=True, timeout=600)

    return response

response = download_file(file_id)
response.raise_for_status()

# Vérifier qu'on reçoit bien un fichier binaire et non du HTML
content_type = response.headers.get("Content-Type", "")
if "text/html" in content_type:
    print("[ERREUR] Google Drive renvoie toujours une page HTML. Vérifiez que le fichier est bien partagé en 'Lecteur - Tout le monde'.", file=sys.stderr)
    sys.exit(1)

total_bytes = 0
with open(LOCAL_EXCEL, "wb") as f:
    for chunk in response.iter_content(chunk_size=8 * 1024 * 1024):
        if chunk:
            f.write(chunk)
            total_bytes += len(chunk)
            print(f"[INFO]   … {total_bytes / 1_048_576:.1f} Mo téléchargés")

file_size = os.path.getsize(LOCAL_EXCEL)
print(f"[INFO] Fichier téléchargé : {file_size / 1_048_576:.1f} Mo → {LOCAL_EXCEL}")

if file_size < 10_000:
    print("[ERREUR] Le fichier téléchargé est trop petit — probablement une page HTML.", file=sys.stderr)
    with open(LOCAL_EXCEL, 'r', errors='replace') as f:
        print(f"[DEBUG] Début du contenu : {f.read(500)}", file=sys.stderr)
    sys.exit(1)

# ── 2. Lecture du fichier Excel ───────────────────────────────────────────────
print("[INFO] Ouverture du fichier Excel (read_only=True)…")
wb = openpyxl.load_workbook(LOCAL_EXCEL, read_only=True, data_only=True)
ws = wb.active

# ── 3. Transformation ─────────────────────────────────────────────────────────
def format_prix(val) -> str:
    if val is None or val == "":
        return ""
    try:
        return f"{float(str(val).replace(',', '.')):.3f}".replace(".", ",")
    except (ValueError, TypeError):
        return str(val)

def csv_escape(val) -> str:
    s = "" if val is None else str(val)
    if any(c in s for c in (";", '"', "\n")):
        return '"' + s.replace('"', '""') + '"'
    return s

total_read = total_filtered = total_exported = 0

print("[INFO] Transformation en cours…")

with open(OUTPUT_FILE, "w", encoding="utf-8-sig", newline="") as f:
    f.write("Activité;Code Article;Variante Logistique;Prix\r\n")
    for row in ws.iter_rows(min_row=2, values_only=True):
        val_a  = str(row[COL_A  - 1]).strip() if row[COL_A  - 1] is not None else ""
        val_b  = row[COL_B  - 1]
        val_cd = row[COL_CD - 1]
        if not val_a and val_b is None and val_cd is None:
            continue
        total_read += 1
        if val_a not in FILTER_VALUES:
            total_filtered += 1
            continue
        f.write(";".join([
            csv_escape("NE1"),
            csv_escape(val_b),
            csv_escape("10"),
            csv_escape(format_prix(val_cd)),
        ]) + "\r\n")
        total_exported += 1
        if total_exported % 10_000 == 0:
            print(f"[INFO]   … {total_exported} lignes exportées")

wb.close()

# ── 4. Résumé ─────────────────────────────────────────────────────────────────
print(f"\n[RÉSULTAT]")
print(f"  Lignes lues       : {total_read}")
print(f"  Lignes filtrées   : {total_filtered}  (col. A ≠ 5322A/5322D)")
print(f"  Lignes exportées  : {total_exported}")
print(f"  Fichier de sortie : {OUTPUT_FILE}")

if total_exported == 0:
    print("[ERREUR] Aucune ligne ne correspond au filtre.", file=sys.stderr)
    sys.exit(1)

print("[OK] Transformation terminée.")

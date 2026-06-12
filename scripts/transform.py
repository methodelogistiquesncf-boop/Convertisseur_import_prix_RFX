"""
transform.py — Transformation ERP → Reflex CSV
------------------------------------------------
1. Télécharge le fichier Excel depuis Google Drive (URL dans $GDRIVE_URL)
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
OUTPUT_FILE  = "output/Export_Articles.csv"
LOCAL_EXCEL  = "input/source.xlsx"

# ── 1. Téléchargement depuis Google Drive ────────────────────────────────────
gdrive_url = os.environ.get("GDRIVE_URL", "").strip()
if not gdrive_url:
    print("[ERREUR] Variable GDRIVE_URL non définie.", file=sys.stderr)
    sys.exit(1)

# Convertir l'URL de partage en URL de téléchargement direct
# Format attendu : https://drive.google.com/file/d/FILE_ID/view?usp=sharing
match = re.search(r"/d/([a-zA-Z0-9_-]+)", gdrive_url)
if not match:
    print("[ERREUR] Impossible d'extraire l'ID du fichier depuis l'URL Google Drive.", file=sys.stderr)
    print(f"  URL reçue : {gdrive_url}", file=sys.stderr)
    sys.exit(1)

file_id      = match.group(1)
download_url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm=t"

print(f"[INFO] Téléchargement depuis Google Drive (ID: {file_id})…")
os.makedirs("input",  exist_ok=True)
os.makedirs("output", exist_ok=True)

session = requests.Session()
response = session.get(download_url, stream=True, timeout=300)

# Google Drive redirige parfois vers une page de confirmation pour les gros fichiers
# On détecte ce cas et on suit la redirection avec le bon token
if "text/html" in response.headers.get("Content-Type", ""):
    token_match = re.search(r'name="confirm"\s+value="([^"]+)"', response.text)
    if not token_match:
        # Nouvelle méthode de confirmation Google Drive
        token_match = re.search(r'confirm=([0-9A-Za-z_-]+)', response.url)
    if token_match:
        confirm_token = token_match.group(1)
        download_url  = f"https://drive.google.com/uc?export=download&id={file_id}&confirm={confirm_token}"
        response      = session.get(download_url, stream=True, timeout=300)

response.raise_for_status()

total_bytes = 0
with open(LOCAL_EXCEL, "wb") as f:
    for chunk in response.iter_content(chunk_size=8 * 1024 * 1024):  # 8 Mo par chunk
        if chunk:
            f.write(chunk)
            total_bytes += len(chunk)
            print(f"[INFO]   … {total_bytes / 1_048_576:.1f} Mo téléchargés")

print(f"[INFO] Fichier téléchargé : {total_bytes / 1_048_576:.1f} Mo → {LOCAL_EXCEL}")

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

total_read     = 0
total_filtered = 0
total_exported = 0

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

        line = ";".join([
            csv_escape("NE1"),
            csv_escape(val_b),
            csv_escape("10"),
            csv_escape(format_prix(val_cd)),
        ])
        f.write(line + "\r\n")
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

import json
import chardet
from pathlib import Path

# Répertoire du script (dynamique)
base_dir = Path(__file__).resolve().parent

# Fichier source et fichier de sortie
filename = base_dir / 'DataModelSchema.txt'
output_filename = base_dir / 'fichier_converti.json'


def detect_encoding(raw_data):
    result = chardet.detect(raw_data)
    return result.get('encoding')

def try_load_with_encoding(path, encoding):
    try:
        text = path.read_text(encoding=encoding, errors='replace')
        return json.loads(text)
    except Exception as e:
        print(f"→ échec avec {encoding!r} : {e}")
        return None

# 1. Lire en binaire et détecter
raw = filename.read_bytes()
detected = detect_encoding(raw)
print(f"Encodage détecté (chardet) : {detected!r}")

# 2. Liste d'encodages à tester
encodings_to_try = []
if detected:
    encodings_to_try.append(detected)
encodings_to_try += ['utf-8-sig', 'utf-8', 'utf-16', 'cp1252', 'latin-1']

# 3. Tenter de charger
data = None
for enc in encodings_to_try:
    data = try_load_with_encoding(filename, enc)
    if data is not None:
        print(f"→ succès avec {enc!r}")
        break

# 4. Si échec total, afficher un extrait brut pour debug
if data is None:
    print("\nÉchec de la désérialisation JSON avec tous les encodages testés.")
    snippet = filename.read_text(encoding=encodings_to_try[-1], errors='replace')[:500]
    print("Extrait (500 premiers caractères) :")
    print(snippet)
    raise SystemExit("Impossible de parser le JSON.")

# 5. Sauvegarder en UTF-8 sans BOM
try:
    with open(output_filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    print(f"Fichier converti et sauvegardé sous : {output_filename}")
except Exception as e:
    print(f"Erreur lors de l’enregistrement : {e}")

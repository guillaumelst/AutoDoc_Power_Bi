import zipfile
import shutil
import json
import chardet
from pathlib import Path
import argparse

# 📁 Répertoire dynamique : dossier où se trouve ce script
DEFAULT_ROOT = Path(__file__).resolve().parent

def detect_encoding(raw_data):
    """Détecte l'encodage à partir d'octets bruts."""
    result = chardet.detect(raw_data)
    return result.get('encoding')

def try_load_with_encoding(path, encoding):
    """Tente de charger un JSON avec un encodage donné."""
    try:
        text = path.read_text(encoding=encoding, errors='replace')
        return json.loads(text)
    except Exception as e:
        print(f"→ Échec avec {encoding!r} : {e}")
        return None

def extract_and_clean_layout(root_dir: Path):
    """Extrait le Layout d'un .pbit, le décode proprement et le sauvegarde en JSON."""
    pbit_files = list(root_dir.glob("*.pbit"))
    if not pbit_files:
        print(f"Aucun fichier .pbit trouvé dans {root_dir}")
        return

    original_pbit = pbit_files[0]
    print(f"Fichier .pbit détecté : {original_pbit.name}")

    copy_pbit = original_pbit.with_name(original_pbit.stem + "_Originale.pbit")
    zip_file = copy_pbit.with_suffix(".zip")
    extract_folder = root_dir / original_pbit.stem

    # Copie et renommage
    shutil.copy(original_pbit, copy_pbit)
    if zip_file.exists():
        zip_file.unlink()
    copy_pbit.rename(zip_file)

    # Extraction partielle : uniquement Layout
    layout_file_found = None
    with zipfile.ZipFile(zip_file, "r") as zip_ref:
        for member in zip_ref.infolist():
            if "Report/Layout" in member.filename.replace("\\", "/"):
                layout_file_found = member.filename
                zip_ref.extract(member, extract_folder)
                break

    if not layout_file_found:
        print("Fichier Layout non trouvé dans le .pbit")
        return

    layout_path = extract_folder / layout_file_found
    print(f"Layout extrait : {layout_path.name}")

    # Détection d'encodage
    raw = layout_path.read_bytes()
    detected = detect_encoding(raw)
    print(f"Encodage détecté (chardet) : {detected!r}")

    encodings_to_try = [detected] if detected else []
    encodings_to_try += ['utf-8-sig', 'utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'cp1252', 'latin-1']

    data = None
    for enc in encodings_to_try:
        data = try_load_with_encoding(layout_path, enc)
        if data is not None:
            print(f"→ Succès avec {enc!r}")
            break

    if data is None:
        print("\n❌ Impossible de parser le JSON avec tous les encodages testés.")
        snippet = layout_path.read_text(encoding=encodings_to_try[-1], errors='replace')[:500]
        print("Extrait (500 premiers caractères) :")
        print(snippet)
        raise SystemExit("Échec de lecture JSON.")

    # Sauvegarde propre
    dest_json = root_dir / "Layout.json"
    with open(dest_json, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    print(f"✅ Fichier Layout converti et sauvegardé sous : {dest_json}")

    # Nettoyage
    try:
        if zip_file.exists(): zip_file.unlink()
        if copy_pbit.exists(): copy_pbit.unlink()
        if extract_folder.exists(): shutil.rmtree(extract_folder)
    except Exception as e:
        print(f"Avertissement suppression fichiers temporaires : {e}")

def main():
    parser = argparse.ArgumentParser(description="Extraction et conversion du Layout d’un fichier .pbit")
    parser.add_argument(
        'root_dir',
        nargs='?',
        default=None,
        help="Chemin vers le dossier contenant le .pbit (facultatif)"
    )
    args = parser.parse_args()
    root_path = Path(args.root_dir) if args.root_dir else DEFAULT_ROOT

    if not root_path.is_dir():
        print(f"Le chemin spécifié n’est pas un dossier valide : {root_path}")
    else:
        extract_and_clean_layout(root_path)

if __name__ == "__main__":
    main()

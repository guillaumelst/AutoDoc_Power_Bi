import zipfile
import shutil
from pathlib import Path
import argparse

# Chemin par défaut si aucun argument fourni
DEFAULT_ROOT = Path(r"C:\Users\GLT\eiffage.com\OneDrive - eiffageenergie.be\Documents\Guillaume\Programmation\Auto. Doc. PBI")

def extract_pbit_file(root_dir: Path):
    """
    Extrait le premier fichier .pbit trouvé dans root_dir, supprime les CustomVisuals,
    extrait le contenu, puis déplace et renomme DataModelSchema en .txt.
    Nettoie le .zip temporaire et le dossier d'extraction.
    """
    # Chercher automatiquement le premier fichier .pbit
    pbit_files = list(root_dir.glob('*.pbit'))
    if not pbit_files:
        print(f"Aucun fichier .pbit trouvé dans {root_dir}")
        return

    original_pbit = pbit_files[0]
    print(f"Fichier source détecté : {original_pbit.name}")

    # Préparer les chemins
    copy_pbit = original_pbit.with_name(original_pbit.stem + '_Originale.pbit')
    zip_file = copy_pbit.with_suffix('.zip')
    extract_folder = root_dir / original_pbit.stem

    # Copier puis renommer en .zip
    shutil.copy(original_pbit, copy_pbit)
    if zip_file.exists():
        zip_file.unlink()
    copy_pbit.rename(zip_file)

    # Extraction en sautant les CustomVisuals
    if extract_folder.exists():
        shutil.rmtree(extract_folder, onerror=lambda func, path, exc: print(f"Avertissement suppression : {path} — {exc}"))
    extract_folder.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        for member in zip_ref.infolist():
            fn = member.filename.replace('\\', '/')
            if fn.startswith("Report/CustomVisuals/"):
                continue

            target = extract_folder / fn
            if member.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                with zip_ref.open(member) as src_file, open(target, 'wb') as dest_file:
                    shutil.copyfileobj(src_file, dest_file)

    print(f"Extraction terminée dans {extract_folder}")

    # Recherche et déplacement de DataModelSchema*
    matches = list(extract_folder.rglob('DataModelSchema*'))
    if not matches:
        print("Fichier 'DataModelSchema' introuvable dans l'extraction.")
    else:
        schema_file = matches[0]
        print(f"Fichier source DataModelSchema détecté : {schema_file.name}")

        dest_txt = root_dir / (schema_file.name + '.txt')
        if dest_txt.exists():
            dest_txt.unlink()
        shutil.move(str(schema_file), str(dest_txt))
        print(f"{schema_file.name} déplacé et renommé en {dest_txt.name}")

    # Nettoyage : suppression du zip et du dossier d'extraction
    try:
        if zip_file.exists():
            zip_file.unlink()
        if extract_folder.exists():
            shutil.rmtree(extract_folder)
        # Supprimer la copie .pbit
        if copy_pbit.exists():
            copy_pbit.unlink()
    except Exception as e:
        print(f"Avertissement suppression fichiers temporaires : {e}")


def main():
    parser = argparse.ArgumentParser(
        description="Extrait un fichier .pbit et récupère DataModelSchema en .txt"
    )
    parser.add_argument(
        'root_dir',
        nargs='?',  # argument facultatif
        default=None,
        help="Chemin du dossier racine contenant le .pbit (optionnel)"
    )
    args = parser.parse_args()
    # Utilisation du chemin par défaut si non fourni
    root_path = Path(args.root_dir) if args.root_dir else DEFAULT_ROOT

    if not root_path.is_dir():
        print(f"Le chemin spécifié n'est pas un dossier valide : {root_path}")
    else:
        extract_pbit_file(root_path)


if __name__ == '__main__':
    main()

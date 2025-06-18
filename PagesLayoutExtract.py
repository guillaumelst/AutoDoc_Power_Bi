import json
import pandas as pd
from pathlib import Path
import argparse

# === 📁 Chemin par défaut ===
DEFAULT_ROOT = Path(r"C:\Users\GLT\eiffage.com\OneDrive - eiffageenergie.be\Documents\Guillaume\Programmation\Auto. Doc. PBI")

def analyser_layout(root_dir: Path):
    # Chercher Layout.json
    layout_file = next(root_dir.glob("Layout.json"), None)
    if not layout_file:
        print(f"❌ Aucun fichier 'Layout.json' trouvé dans : {root_dir}")
        return

    print(f"📂 Fichier détecté : {layout_file.name}")

    # Charger le fichier JSON
    try:
        with open(layout_file, encoding='utf-8') as f:
            layout_json = json.load(f)
    except Exception as e:
        print(f"❌ Erreur lors du chargement : {e}")
        return

    # Analyser les sections
    resultats = []
    for section in layout_json.get("sections", []):
        nom = section.get("displayName", "Sans nom")
        page_id = section.get("name", "Inconnu")
        config_str = section.get("config")

        statut = "Non défini (None)"
        if config_str:
            try:
                config_data = json.loads(config_str)
                visibility = config_data.get("visibility")
                if visibility == 0:
                    statut = "Visible"
                elif visibility == 1:
                    statut = "Masquée"
                else:
                    statut = f"Non défini ({visibility})"
            except json.JSONDecodeError:
                statut = "Erreur JSON dans config"

        resultats.append({
            "Nom de la page": nom,
            "ID de la page": page_id,
            "Visibilité": statut
        })

    # Affichage du tableau dans la console
    df = pd.DataFrame(resultats)
    print("\n✅ Résultat :")
    print(df)

def main():
    parser = argparse.ArgumentParser(description="Affiche la visibilité des pages d’un Layout.json dans la console")
    parser.add_argument('root_dir', nargs='?', default=None, help="Chemin du dossier contenant Layout.json")
    args = parser.parse_args()

    root = Path(args.root_dir) if args.root_dir else DEFAULT_ROOT
    if not root.exists() or not root.is_dir():
        print(f"❌ Dossier invalide : {root}")
        return

    analyser_layout(root)

if __name__ == "__main__":
    main()

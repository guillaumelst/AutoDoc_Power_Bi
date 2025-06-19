# ⚡ Documentation automatique Power BI 🚀

Ce projet fournit un ensemble de scripts Python 🐍 permettant d'automatiser l'extraction et la génération de la documentation d'un fichier Power BI Template (`.pbit`). Il :

- 📂 Extrait le fichier template et récupère le schéma du modèle de données (`DataModelSchema`).
- 🔄 Convertit ce schéma en JSON pour le traitement.
- 📄 Génère un document Word détaillant tables, mesures, colonnes calculées, partitions, hiérarchies et disposition graphique du rapport.

---

## 📚 Table des matières

1. [🚀 Prérequis](#-prérequis)
2. [🔧 Installation](#-installation)
3. [📂 Structure du projet](#-structure-du-projet)
4. [▶️ Utilisation](#-utilisation)
5. [✨ Détails des scripts](#-détails-des-scripts)
6. [🖼️ Exemple de sortie](#-exemple-de-sortie)
7. [📄 Licence](#-licence)

---

## 🚀 Prérequis

- Python 3.7 ou supérieur 🐍
- Modules Python :
  - `python-docx` 📁
  - `chardet` 🔍

Installez-les via :

```bash
pip install python-docx chardet
```

---

## 🔧 Installation

1. Clonez ce dépôt :

```bash
git clone https://github.com/votre-utilisateur/votre-projet.git
cd votre-projet
```

2. Installez les dépendances (voir [🚀 Prérequis](#-prérequis)).

---

## 📂 Structure du projet

```text
./
├── UnzipPBIP.py          # Extraction du .pbit et récupération de DataModelSchema
├── EncodeJSON.py         # Conversion du schéma texte en JSON
├── autodoc.py            # Génération du document Word à partir du JSON
├── LayoutFinder.py       # Extraction de la disposition des visuels dans le rapport
├── Table Doc .py         # Génération de la documentation tabulaire des visuels
├── Main_doc_PBI.py       # Script principal orchestrant les étapes
├── requirements.txt      # (optionnel) Liste des dépendances Python
├── images/               # Dossier pour les exemples d’images
└── README.md             # Ce fichier
```

---

## ▶️ Utilisation

Le script principal `Main_doc_PBI.py` exécute toutes les étapes de façon automatique :

```bash
python Main_doc_PBI.py
```

Cela exécute successivement :

1. 🔍 Extraction et nettoyage du `.pbit`, récupération de `DataModelSchema.txt`
2. 🔄 Conversion du `.txt` en JSON (`fichier_converti.json`)
3. ✍️ Génération du document `documentation.docx`
4. 📊 Analyse du layout graphique via `LayoutFinder.py`
5. 📄 Documentation des visuels via `Table Doc .py`

Chaque script est exécuté via `subprocess` avec une mesure du temps d'exécution.

Vous pouvez également lancer chaque étape individuellement :

```bash
python UnzipPBIP.py
python EncodeJSON.py
python autodoc.py
python LayoutFinder.py
python "Table Doc .py"
```

---

## ✨ Détails des scripts

### 1. `UnzipPBIP.py` 💚
- Recherche et copie le premier fichier `.pbit` du dossier (ou chemin fourni)
- Le renomme en `.zip` et extrait son contenu
- Récupère le fichier `DataModelSchema` et le renomme en `.txt`
- Nettoie les fichiers temporaires

### 2. `EncodeJSON.py` 📜
- Détecte l'encodage du fichier texte contenant le modèle
- Tente plusieurs encodages (`utf-8-sig`, `utf-16`, `cp1252`, etc.) pour parser le JSON proprement
- Sauvegarde le résultat dans `fichier_converti.json`

### 3. `autodoc.py` 📄
- Lit `fichier_converti.json`
- Extrait les éléments du modèle de données : tables, mesures, colonnes calculées, hiérarchies, partitions
- Génère un document Word stylisé

### 4. `LayoutFinder.py` 📈
- Analyse les fichiers de layout pour repérer les visuels présents par page
- Permet d'associer tables, champs et pages d'utilisation

### 5. `Table Doc .py` 📃
- Compile toutes les données extraites (modèle + layout)
- Produit une documentation tabulaire lisible, exportable si besoin

---

## 🖼️ Exemple de sortie

Voici un aperçu du `.docx` généré :  
![Exemple de sortie](images/output_example.png)

---

## 📄 Licence

Ce projet est distribué sous licence MIT.

La licence MIT est une licence libre et permissive. Elle permet :

- ✔️ d’utiliser, copier, modifier et distribuer le logiciel, à des fins privées, éducatives ou commerciales, sans restriction.
- ✔️ d’intégrer le logiciel dans des projets propriétaires.

Seule condition : inclure la notice de copyright et la licence MIT dans toutes les copies ou distributions du logiciel.

Pour plus de détails, consultez le fichier [LICENSE](LICENSE).

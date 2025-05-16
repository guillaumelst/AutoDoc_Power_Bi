# ⚡ Documentation automatique Power BI 🚀

Ce projet fournit un ensemble de scripts Python 🐍 permettant d'automatiser l'extraction et la génération de la documentation d'un fichier Power BI Template (`.pbit`). Il :

- 📂 Extrait le fichier template et récupère le schéma du modèle de données (`DataModelSchema`).
- 🔄 Convertit ce schéma en JSON pour le traitement.
- 📄 Génère un document Word détaillant tables, mesures, colonnes calculées, partitions et hiérarchies du modèle.

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
  - `python-docx` 📑
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
├── Main_doc_PBI.py       # Script principal orchestrant les étapes
├── requirements.txt      # (optionnel) Liste des dépendances Python
├── images/               # Dossier pour les exemples d’images
└── README.md             # Ce fichier
```

---

## ▶️ Utilisation

Le script principal `Main_doc_PBI.py` exécute les trois étapes dans l'ordre :

```bash
python Main_doc_PBI.py
```

Cela :

1. 🔍 Extrait et nettoie le `.pbit`, récupère `DataModelSchema` au format `.txt`.
2. 🔄 Convertit ce `.txt` en JSON (`fichier_converti.json`).
3. ✍️ Génère `documentation.docx` contenant la documentation de votre modèle.

Vous pouvez également lancer chaque étape séparément :

```bash
python UnzipPBIP.py [chemin/vers/dossier]
python EncodeJSON.py
python autodoc.py
```

---

## ✨ Détails des scripts

### 1. `UnzipPBIP.py` 🗜️

- Recherche et copie le premier fichier `.pbit` du dossier (ou chemin fourni).
- Renomme la copie en `.zip` et l'extrait en omettant les `CustomVisuals`.
- Récupère `DataModelSchema` et le renomme en `.txt` dans le dossier racine.
- Nettoie les fichiers temporaires.

### 2. `EncodeJSON.py` 📝

- Détecte l'encodage du fichier `DataModelSchema.txt`.
- Tente plusieurs encodages (`utf-8-sig`, `utf-16`, `cp1252`, etc.) pour parser le JSON.
- Sauvegarde le résultat propre en `fichier_converti.json`.

### 3. `autodoc.py` 📄

- Lit `fichier_converti.json` et extrait métadonnées :
  - Tables standards et calculées
  - Mesures, colonnes calculées, partitions, hiérarchies
- Construit un document Word (`documentation.docx`) stylé, avec tableaux et en-têtes colorés.

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

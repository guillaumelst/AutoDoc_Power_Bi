import subprocess
import sys
import time
from pathlib import Path

def run_python_script(script_path: Path):
    print(f"\n🚀 Lancement du script : {script_path.name}")
    start = time.time()
    try:
        subprocess.run([sys.executable, str(script_path)], check=True)
        duration = time.time() - start
        print(f"✅ Script {script_path.name} terminé en {duration:.2f} secondes.")
        return duration
    except subprocess.CalledProcessError as e:
        print(f"❌ Erreur lors de l'exécution du script {script_path.name}: {e}")
        return None

# Répertoire racine dynamique (dossier du script)
root_dir = Path(__file__).resolve().parent

# Liste des scripts à exécuter, chemins relatifs au répertoire racine
scripts = [
    root_dir / "UnzipPBIP.py",
    root_dir / "EncodeJSON.py",
    root_dir / "autodoc.py",
]

total_start = time.time()
durations = {}

# Exécution de chaque script
for script in scripts:
    durations[script.name] = run_python_script(script) or 0.0

total_duration = time.time() - total_start

# Résumé
print("\n📊 Résumé des temps d'exécution :")
for name, dur in durations.items():
    print(f"  ⏱️ {name} : {dur:.2f} s")
print(f"\n🧾 Temps total d'exécution : {total_duration:.2f} s")

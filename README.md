# Générateur et Fusionneur d'Horaires en Python 

Ce projet permet de :
- Générer un fichier `.docx` contenant les horaires d'un programme donné à partir d'un fichier `.csv`
- Fusionner automatiquement plusieurs documents Word contenant ces horaires
- Ouvrir automatiquement les fichiers générés
- Utiliser une interface terminale simple et interactive

---

## Prérequis

- Python 3.7+
- pip (gestionnaire de paquets)

---

## Installation

1. Clone le dépôt ou télécharge les fichiers source.
2. Installe les dépendances via le terminal :

```bash
pip install -r requirements.txt
```

---

## Organisation des fichiers

```
/votre_dossier/
├── horaire.csv              # Fichier source contenant les données d'horaires
├── main.py                  # Fichier principal à exécuter
├── requirements.txt         # Dépendances du projet
├── README.md                # Ce fichier
```

---

## Utilisation

Lancer le programme :

```bash
python main.py
```

Ensuite, tu peux :
1. Générer un document contenant les cours d’un génie spécifique
2. Fusionner tous les documents `.docx` existants dans le dossier
3. Quitter le programme

---

## Format du fichier CSV attendu

Le fichier `horaire.csv` doit contenir des lignes structurées ainsi :

```csv
Sigle,Type,Groupe,Jour,Debut,Duree,Fin,Frequence,Examen
LOG1015,Cours,A,Lundi,08:30,3h,11:30,Chaque semaine,2025-06-10
...
```

---

## À faire / améliorations possibles

- Interface graphique (Tkinter, PyQt, etc.)
- Export au format PDF

---

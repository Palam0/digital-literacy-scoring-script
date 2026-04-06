# Script de cotation automatisée du questionnaire de littératie numérique

Ce dépôt contient un script Python permettant d’automatiser la cotation d’un questionnaire de littératie numérique à partir d’un fichier Excel de réponses.

L’objectif de cette mise à disposition est de favoriser la **transparence**, la **reproductibilité** et la **réutilisation** de l’outil par d’autres chercheurs ou professionnels.

## Fonction du script

Le script :

* lit un fichier Excel contenant les réponses des participants ;
* applique une grille de cotation prédéfinie ;
* calcule les scores pour chacune des six dimensions du questionnaire ;
* calcule un score total ;
* génère automatiquement un nouveau fichier Excel contenant les scores de chaque participant.

## Dimensions évaluées

Le questionnaire est structuré autour de six dimensions :

* **Résilience**
* **Esprit critique**
* **Comportement social dans les environnements numériques**
* **Compétences techniques**
* **Traitement de l’information**
* **Création**

## Principe de cotation

Chaque question est associée :

1. à une **dimension** ;
2. à un ensemble de **modalités de réponse** ;
3. à un **coefficient** attribué à chaque modalité.

Le script compare les réponses de chaque participant à cette grille de cotation et additionne les scores obtenus pour chaque dimension.

### Particularité du calcul

Pour chaque dimension, si le score final est négatif, il est automatiquement ramené à **0**.

Le **score total** correspond à la somme des six scores dimensionnels.

## Prérequis

Ce script nécessite :

* **Python 3**
* le package **openpyxl**

Installation du package :

```bash
pip install openpyxl
```

## Structure des fichiers

Le script attend un fichier d’entrée au format Excel contenant les réponses des participants.

Exemple d’organisation du dossier :

```text
.
├── script_cotation_questionnaire.py
├── README.md
├── donnees/
│   └── questionnaire_reponses.xlsx
└── resultats/
```

## Utilisation

1. Placer le fichier de réponses dans le dossier `donnees/` sous le nom :

```text
questionnaire_reponses.xlsx
```

2. Exécuter le script :

```bash
python script_cotation_questionnaire.py
```

3. Le script génère automatiquement un fichier de sortie dans le dossier `resultats/` :

```text
scores_questionnaire.xlsx
```

## Sortie produite

Le fichier généré contient, pour chaque participant :

* un identifiant ;
* un score de **Résilience** ;
* un score d’**Esprit critique** ;
* un score de **Comportement social** ;
* un score de **Compétences techniques** ;
* un score de **Traitement de l’information** ;
* un score de **Création** ;
* un **score total**.

## Précautions

* Le script suppose que l’ordre des réponses dans le fichier Excel correspond à l’ordre des questions défini dans la grille de cotation.
* Avant toute diffusion publique, il est recommandé de **ne pas inclure de données réelles identifiantes** dans le dépôt.
* Si nécessaire, utiliser un fichier d’exemple anonymisé.

## Objectif de la mise à disposition

La mise à disposition publique de ce script vise à faciliter l’usage de l’outil, à limiter les erreurs de cotation manuelle, et à permettre sa réutilisation dans d’autres contextes de recherche ou de pratique.

## Auteur

* Kévin Berenger

Script développé dans le cadre d’un travail de recherche sur la littératie numérique.



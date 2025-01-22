
# Old Office File Manager Manual

_For: Mayr-Melnhof Group_  
_By: Heriaud Thomas_  
_Date: 22/01/2025_

---

## Table des Matières

- [Avant d’ouvrir l’application](#avant-douvrir-lapplication)
- [Deux façons d’ouvrir l’application](#deux-façons-douvrir-lapplication)
  - [En mode « utilisateur »](#en-mode-utilisateur)
  - [En mode « admin »](#en-mode-admin)
- [Fonctionnement](#fonctionnement)
  - [Choix des dossiers](#choix-des-dossiers)
  - [Les Logs](#les-logs)
  - [Options supplémentaires](#opt-supp)
- [Le rapport (.csv)](#le-rapport-csv)

---

## Avant d’ouvrir l’application

Quelques points à vérifier :

- Être sous Windows et posséder **PowerShell 5.0 minimum** (pré-installé avec Windows 10).
- En cas de conversion des fichiers, avoir la version moderne de **Word**, **PowerPoint**, et **Excel**.

---

## Deux façons d’ouvrir l’application

### En mode « utilisateur »
Ce mode permet uniquement la **lecture de fichiers** avec les mêmes droits que l’utilisateur et la **génération d’un rapport**.

### En mode « admin »
Ce mode permet la **génération de rapports** ET la **conversion de fichiers** en fonction des applications Office installées.

---

## Fonctionnement

### Choix des dossiers

- Utiliser le bouton **« Parcourir »** pour choisir le dossier à scanner. Tous les dossiers enfants seront également scannés à la recherche de fichiers `.doc`, `.xls`, et `.ppt`.
- Utiliser le bouton **« Parcourir »** pour choisir le dossier dans lequel un rapport au format CSV sera généré.  
  **Important :** Les droits d’écriture dans le dossier sont nécessaires, car aucun dossier par défaut n’est défini.



### Les Logs

Dans cette zone défilante s’afficheront les éléments suivants :
- Erreur de sélection de dossier.
- Chemins complets sélectionnés.
- Fichiers avec les extensions recherchées.
- Avancée de la conversion.
- Noms et chemins du rapport.



### Options supplémentaires :
+ Une case à cocher, si sélectionnée, lance la conversion des fichiers (nécessite que l’application soit lancée en mode administrateur).
+ Un bouton démarre le processus de recherche des fichiers et conversion (si activée), tout en affichant les étapes dans la zone de logs.

---

## Le rapport (.csv)

Le rapport généré est un document CSV. Si le délimiteur utilisé par l’utilisateur n’est pas correct, une manipulation est nécessaire pour rendre les données lisibles.

### Procédure de correction :

1. Aller dans l’onglet **Données**, puis cliquer sur **À partir d’un tableau ou d’une plage** dans la section _Récupérer et transformer des données_.
2. Cocher **« Mon tableau comporte des en-têtes »** pour pouvoir trier de grandes quantités de données, puis valider. Ne pas changer la plage : elle est automatiquement correcte.
3. Dans l’onglet **Accueil**, au niveau de la section _Transformer_, sélectionner **Fractionner la colonne** puis l’option **Par délimiteur**.
4. Vérifier que **« Virgule »** est bien sélectionnée, puis valider avec **Ok**.
5. Valider avec **Fermer et charger** pour le faire apparaître dans le document original.
6. Enregistrer le document.

---

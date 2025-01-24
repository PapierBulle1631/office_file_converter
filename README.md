
# Old Office File Manager Manual (office_converter.exe)

_For: Mayr-Melnhof Group_  
_By: Heriaud Thomas_  
_Date: 22/01/2025_

---

## Table des Matières

- [Avant d’ouvrir l’application](#avant-douvrir-lapplication)
- [Deux façons d’ouvrir l’application](#deux-façons-douvrir-lapplication)
  - [Lecture de fichiers](#lecture-de-fichiers-)
  - [Conversion de fichiers](#conversion-de-fichiers-)
  - [Génération d'un rapport](#génération-dun-rapport-)
  - [Affichage de logs :](#affichage-de-logs-)
- [Fonctionnement](#fonctionnement)
  - [Choix des dossiers](#choix-des-dossiers)
  - [Les Logs](#les-logs)
  - [Options supplémentaires](#options-supplémentaires-)
- [Le rapport (.csv)](#le-rapport-csv)
  - [Procédure de correction](#procédure-de-correction-)

---

## Avant d’ouvrir l’application

Quelques points à vérifier :

- Être sous Windows et posséder **PowerShell 5.0 minimum** (pré-installé avec Windows 10).
- En cas de conversion des fichiers, avoir la version moderne de **Word**, **PowerPoint**, et **Excel**.

---

## Ce que fait le logiciel 
- ### Lecture de fichiers :
Tout les dossiers et sous dossiers sont scannés à la recherche de documents au format .doc .ppt et .xls (soit les anciennes versions d'Office).
- ### Conversion de fichiers :
Tout les fichiers subissent une tentative de conversion dans le format moderne .docx .pptx et .xls en fonction de leur extension.
- ### Génération d'un rapport :
Un rapport au format .csv est généré content le chemin complet et le nom des fichiers convertis ainsi que leur taille en octet (la dernière ligne contient la taille totale en Go).
- ### Affichage de logs :
Une zone réservée à l'affichage est présente. Y seront affichées toute les actions et leurs résultats.

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
+ Un bouton de suppression des fichiers convertis se débloque après le listing des documents.

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

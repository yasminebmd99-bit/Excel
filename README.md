Cahier des Charges pour Application Web de Gestion et Réorganisation de Fichiers Excel
1. Objectif du projet
Développer une application web capable d'importer plusieurs fichiers Excel, de les fusionner en un fichier unique, puis de permettre l'import et la correction d'un fichier Excel mal structuré en réassociant correctement chaque colonne à son titre spécifique.


2. Fonctionnalités principales

Importation de plusieurs fichiers Excel (.xlsx, .xls, .csv) via interface utilisateur.

Fusion automatique des fichiers importés en un seul fichier consolidé.

Importation d’un fichier Excel erroné (colonnes mélangées).

Outil de réorganisation permettant de réattribuer manuellement ou automatiquement chaque colonne au bon en-tête/titre.

Export du fichier corrigé en Excel.

Gestion des erreurs d'importation (formats incompatibles, données corrompues).

Interface utilisateur simple et intuitive pour visualiser les fichiers importés et les colonnes à arranger.

3. Technologies suggérées
Frontend : HTML, CSS, JavaScript (Framework possible : React ou Vue.js)

Backend : Node.js (ou équivalent) pour traitement des fichiers Excel

Bibliothèque traitement Excel : XLSX.js ou équivalent

Hébergement : serveur web adapté avec gestion des fichiers temporaires

4. Contraintes techniques
Support des formats Excel courants (.xlsx, .xls, .csv).

Capacité à gérer des fichiers volumineux (taille à préciser par le client).

Sécurité des données importées, aucune donnée ne doit être stockée de manière permanente sans consentement.

Performance optimisée pour une fusion et réorganisation rapides.
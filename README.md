Fusion de notes
===============

Problématique
-------------
Les différents outils générant ou gérant des notes (Moodle, QCM automatiques, Intracursus) récupèrent les notes depuis des bases de données ou des listings potentiellement différents.
En exportant vers le tableur, pour fusionner ces notes, un étudiant donné n'apparait pas toujours sous le même nom (casse, prénom-nom vs nom-prénom, accents, tirets dans les noms, etc.)
De plus, des étudiants sont présents dans certains et pas dans d'autres, selon le moment où le listing a été mis à jour (étudiants démissionnaires, ou à l'inverse inscrits tardivement).

Principe
--------
Le script suivant essaye de **fusionner ces données non homogènes**, copiées dans un fichier XLSX, en associant autant que possible à un étudiant toutes les notes qui lui correspondent.


Entrée
------
Un **fichier XLSX**, avec **une feuille pour chaque source de donnée**.

Attention, les données doivent être disposées en colonne, *sans entêtes* ni lignes blanches en haut de la feuille.

*   La **première feuille** correspond à la **liste des étudiants** (nom et prénom), telle qu'on veut la récupérer. 
    Typiquement, c'est la liste apparaissant dans Intracursus, pour pouvoir importer ensuite les notes dans Intracursus.

    Les *noms et prénoms* doivent être dans la *colonne A*, ou éventuellement les *colonnes A et B*.

    L'ordre (prénom-nom ou nom-prénom) n'a pas d'importance.

    Les colonnes suivantes (C, D, E...) sont simplement ignorées.

*   Les **autres feuilles** correspondent à des listes de **notes**, précédées du prénom et nom des étudiants.

    Là encore, deux possibilités :
    - *colonne A pour les noms et prénoms*, *colonne B, C, D... pour les notes*.
    - *colonne A et B pour les noms et prénoms*, *colonne C, D, E... pour les notes*.

    L'ordre (prénom-nom ou nom-prénom) n'a non plus d'importance ici.

    Il peut y avoir plusieurs colonnes *successives* de notes.
    Si une colonne est vide, toutes les colonnes suivantes sont ignorées.


Sortie
------

Un nouveau fichier XLSX est généré dans le même répertoire.
Il reprend le fichier original en y ajoutant une feuille 'Fusion'.
Les données qui ont été fusionnées avec des heuristiques peu fiables correspondent aux cellules en rouge, à vérifier manuellement.


Installation des dépendances
----------------------------
En supposant python (3.6+) et pip déjà installés :
$ python -m pip --user fire openpyxl

Attention, il est fréquent que sous Linux, l'exécutable pour Python 3+ 
s'appelle python3, et non python :
$ python3 -m pip install --user fire openpyxl


Exemple d'utilisation
---------------------

> $ python merge-scores.py notes-a-fusionner.xlsx

Un fichier *notes-a-fusionner_output.xlsx* est généré, avec les données fusionnées dans l'onglet fusion.

Ou encore (Linux) :
> $ chmod u+x merge-scores.py
> $ ./merge-scores.py notes-a-fusionner.xlsx


FAQ
---
*   *Pourquoi ne pas utiliser le format OpenDocument (.ods) ?*
    Je n'ai pas trouvé de librairie python permettant un usage avancé de ce format (formatage des cellules notamment).
    Je me suis dès lors rabattu sur le format OpenXML (.xlsx)...



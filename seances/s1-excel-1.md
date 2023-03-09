<p style="text-align:left;">
    [Retour au sommaire](../README.md)
    <span style="float:right;">
        [Séance 2 - Excel: Fonction de bases - suite](s2-excel-2.md)
    </span>
</p>

# Excel - Fonction de base

## Rappels

### Le tableur

Un tableur est un logiciel qui permet de manipuler des données sous forme de tableaux. Il permet de créer des feuilles de calculs, des graphiques, des tableaux croisés dynamiques, des bases de données, des présentations, etc.

### La feuille de calcul

Une feuille de calcul est un tableau composé de lignes et de colonnes. Chaque cellule est identifiée par une lettre de colonne et un numéro de ligne. Par exemple, la cellule A1 est la cellule de la première ligne et de la première colonne.

### La formule

Une formule est une expression qui renvoie une valeur. Elle commence par un signe égal (=) et peut contenir des références, des fonctions, des opérateurs, des constantes et des noms.

### Les références

Une référence est une adresse de cellule ou de plage de cellules. Elle peut être absolue, relative ou mixte. Le signe de dollar ($) est utilisé pour indiquer si une référence est absolue ou relative.

#### Référence absolue

Une référence absolue est une référence qui ne change pas lorsqu'elle est copiée ou déplacée. Elle est composée d'une lettre de colonne et d'un numéro de ligne précédés d'un signe de dollar ($). 

Par exemple, la référence absolue `$A$1` renvoie la valeur de la cellule A1.

#### Référence relative

Une référence relative est une référence qui change lorsqu'elle est copiée ou déplacée. Elle est composée d'une lettre de colonne et d'un numéro de ligne. Dans ce cas là, ni la ligne ni la colonne n'est fixée. 

Par exemple, la référence relative `A1` renvoie la valeur de la cellule A1.

#### Référence mixte

Une référence mixte est une référence où l'un des deux éléments (la lettre de colonne ou le numéro de ligne) est fixé et l'autre est relatif. Elle est composée d'une lettre de colonne et d'un numéro de ligne précédés d'un signe de dollar ($).

Par exemple, la référence mixte `$A1` renvoie la valeur de la cellule A1. La colonne `A` est fixée mais pas la ligne.

### Les liaisons

Une liaison est une référence qui permet de faire référence à une plage de cellules d'une autre feuille de calcul. Elle est composée du nom de la feuille de calcul, d'un point et d'une référence.

#### Liaison entre deux cellules d'une même feuille de calcul

C'est la référence définie précédemment, par exemple `=A1`.

#### Liaison entre deux cellules de deux feuilles de calcul

Il faut cette fois-ci préciser le nom de la feuille de calcul.

La syntaxe est la suivante : `=NomFeuille!CL` où `NomFeuille` est le nom de la feuille de calcul et `CL` est la référence de la cellule.

Par exemple, la liaison `=Feuille1!A1` renvoie la valeur de la cellule A1 de la feuille de calcul `Feuille1`.

#### Liaison entre deux cellules de deux classeurs différents

Il faut cette fois-ci préciser le nom du classeur et le nom de la feuille de calcul.

La syntaxe est la suivante : `=NomClasseur!NomFeuille!CL` où `NomClasseur` est le nom du classeur, `NomFeuille` est le nom de la feuille de calcul et `CL` est la référence de la cellule.

Par exemple, la liaison `=Classeur1!Feuille1!A1` renvoie la valeur de la cellule A1 de la feuille de calcul `Feuille1` du classeur `Classeur1`.

### Les messages d'erreurs

Les messages d'erreurs sont des messages qui apparaissent dans une cellule lorsque la formule est incorrecte. Ils sont composés d'un code d'erreur et d'un message d'erreur.

Les principaux codes d'erreur sont résumés dans le tableau suivant :

| Code d'erreur | Message d'erreur |
|---------------|------------------|
| #DIV/0!       | Division par zéro |
| #N/A          | Valeur non disponible |
| #NAME?        | Nom de fonction incorrect |
| #NULL!        | Valeur nulle |
| #NUM!         | Valeur numérique incorrecte |
| #REF!         | Référence incorrecte |
| #VALUE!       | Valeur incorrecte |


## Les fonctions

Une fonction est une expression qui renvoie une valeur. Elle est composée d'un nom de fonction, d'un ou plusieurs arguments et d'un signe de parenthèse ouvrante et fermante.

### Quelques fonctions de base

#### Somme

La fonction `SOMME` permet de calculer la somme des valeurs d'une plage de cellules.

La syntaxe est la suivante : `=SOMME(CL)` où `CL` est la référence de la plage de cellules.

Par exemple, la formule `=SOMME(A1:A5)` renvoie la somme des valeurs des cellules A1 à A5.

#### Moyenne

La fonction `MOYENNE` permet de calculer la moyenne des valeurs d'une plage de cellules.

La syntaxe est la suivante : `=MOYENNE(CL)` où `CL` est la référence de la plage de cellules.

Par exemple, la formule `=MOYENNE(A1:A5)` renvoie la moyenne des valeurs des cellules A1 à A5.

### Les conditions

Une condition est une expression qui permet de vérifier si une condition est vraie ou fausse. Elle est composée d'une expression logique et d'un opérateur de comparaison.

#### La fonction SI

La fonction `SI` permet de renvoyer une valeur si une condition est vraie et une autre valeur si elle est fausse.

La syntaxe est la suivante : `=SI(Condition;ValeurSiVrai;ValeurSiFaux)` où `Condition` est l'expression logique, `ValeurSiVrai` est la valeur renvoyée si la condition est vraie et `ValeurSiFaux` est la valeur renvoyée si la condition est fausse.

Par exemple, la formule `=SI(A1>10;A1;0)` renvoie la valeur de la cellule A1 si elle est supérieure à 10 et renvoie 0 sinon.

#### Les opérateurs de comparaison

Les opérateurs de comparaison sont des opérateurs qui permettent de comparer deux valeurs. Ils sont composés d'un opérateur de comparaison et de deux valeurs.

Les principaux opérateurs de comparaison sont résumés dans le tableau suivant :

| Opérateur de comparaison | Description |
|--------------------------|-------------|
| =                        | Égal à      |
| <>                       | Différent de |
| >                        | Supérieur à |
| <                        | Inférieur à |
| >=                       | Supérieur ou égal à |
| <=                       | Inférieur ou égal à |

#### La fonction NB.SI

La fonction `NB.SI` permet de compter le nombre de cellules qui vérifient une condition.

La syntaxe est la suivante : `=NB.SI(CL;Condition)` où `CL` est la référence de la plage de cellules et `Condition` est l'expression logique.

Par exemple, considérons le tableau suivant, représentant un ensemble de ventes:



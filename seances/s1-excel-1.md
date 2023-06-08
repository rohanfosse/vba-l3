<p style="text-align:left;">
    [Retour au sommaire](../README.md)
    <span style="float:right;">
        [Séance 2 - Excel: Fonction de bases - suite](s2-excel-2.md)
    </span>
</p>

# Excel - Fonction de base

## Introduction

### Objectifs

Pour ce premier cours, nous avons plusieurs objectifs :

- Découvrir l'interface d'Excel
- Découvrir les formules de base
- Découvrir les références
- Découvrir les fonctions de base

### Prérequis

Pour ce cours, nous allons utiliser Excel. Il est donc nécessaire d'avoir un ordinateur avec Excel installé. Si vous n'avez pas Excel, vous pouvez utiliser [Excel Online](https://office.live.com/start/Excel.aspx) qui est gratuit.

Si vous êtes étudiant, vous pouvez également bénéficier de la suite Office gratuitement. Pour cela, il vous suffit de vous rendre sur le site [https://www.office.com/](https://www.office.com/) et de vous connecter avec votre adresse mail universitaire.

## Pour commencer

Avant de commencer, nous allons faire un petit rappel sur les notions de base d'Excel.

### Qu'est ce qu'un tableur ?

Un tableur est un logiciel qui permet de manipuler des données sous forme de tableaux. Il permet de créer des feuilles de calculs, des graphiques, des tableaux croisés dynamiques, des bases de données, des présentations, etc.

### La feuille de calcul

Une feuille de calcul est un tableau composé de lignes et de colonnes. Chaque cellule est identifiée par une lettre de colonne et un numéro de ligne. Par exemple, la cellule `A1` est la cellule de la première ligne et de la première colonne.

### La formule

Une formule est une expression qui renvoie une valeur. Elle commence par un signe égal (=) et peut contenir des références, des fonctions, des opérateurs, des constantes et des noms.

Dans Excel, il existe deux types de formules :

- Les formules de calcul
- Les formules de texte

#### Les formules de calcul

Les formules de calcul permettent de faire des calculs sur des nombres. Elles sont composées d'opérateurs, de références, de fonctions et de constantes.

Par exemple, la formule `=A1+B1` renvoie la somme des cellules A1 et B1.

#### Les formules de texte

Les formules de texte permettent de manipuler du texte. Elles sont composées d'opérateurs, de références, de fonctions et de constantes.

Par exemple, la formule `=CONCATENER(A1;B1)` renvoie la concaténation des cellules A1 et B1.

#### La barre de formule

La barre de formule est la barre située en haut de la feuille de calcul. Elle permet d'afficher la formule de la cellule sélectionnée. Elle permet également de saisir une formule.

### Les références

Une référence est une adresse de cellule ou de plage de cellules. Elle peut être absolue, relative ou mixte. Le signe de dollar ($) est utilisé pour indiquer si une référence est absolue ou relative.

#### Référence absolue

Une référence absolue est une référence qui ne change pas lorsqu'elle est copiée ou déplacée. Elle est composée d'une lettre de colonne et d'un numéro de ligne précédés d'un signe de dollar ($).

Par exemple, la référence absolue `$A$1` renvoie la valeur de la cellule A1. Si on copie cette référence dans une autre cellule, elle renverra toujours la valeur de la cellule A1. De même si on "étire" la référence vers le bas ou vers la droite.

#### Référence relative

Une référence relative est une référence qui change lorsqu'elle est copiée ou déplacée. Elle est composée d'une lettre de colonne et d'un numéro de ligne. Dans ce cas là, ni la ligne ni la colonne n'est fixée. 

Par exemple, la référence relative `A1` renvoie la valeur de la cellule A1.

#### Référence mixte

Une référence mixte est une référence où l'un des deux éléments (la lettre de colonne ou le numéro de ligne) est fixé et l'autre est relatif. Elle est composée d'une lettre de colonne et d'un numéro de ligne précédés d'un signe de dollar ($).

Par exemple, la référence mixte `$A1` renvoie la valeur de la cellule A1. La colonne `A` est fixée mais pas la ligne.
Cela veut dire que si on "étire" la référence vers le bas, la colonne restera fixée mais la ligne changera. Mais si on "étire" la référence vers la droite, ni la colonne ni la ligne ne changeront.

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


---

## Les fonctions

Maintenant que nous avons repris certaines notions de base d'Excel, nous allons attaquer les fonctions.

Une fonction peut être vu comme une boîte noire. On lui donne des arguments et elle nous renvoie une valeur.
Si l'on devait faire une analogie avec la vie réelle, on pourrait dire que la fonction est un robot. On lui donne des instructions et il nous renvoie un résultat.

Dans notre cas, une fonction est une expression qui renvoie une valeur.

Elle est composée d'un nom de fonction (le nom du robot qui va effectuer les instructions dans notre analogie), d'un ou plusieurs arguments (les instructions) et d'un signe de parenthèse ouvrante et fermante.

La syntaxe d'une fonction est la suivante :

`=NomFonction(Argument1;Argument2;...)` où `NomFonction` est le nom de la fonction, `Argument1`, `Argument2`, ... sont les arguments de la fonction.

### Quelques fonctions de base

Pour essayer de mieux comprendre le fonctionnement des fonctions, nous allons voir quelques fonctions de base.

#### Somme

La fonction `SOMME` permet de calculer la somme des valeurs d'une plage de cellules.

La syntaxe est la suivante : `=SOMME(CL)` où `CL` est la référence de la plage de cellules.

En reprenant l'analogie avec la vie réelle, on pourrait dire que la fonction `SOMME` est un robot qui va additionner les valeurs des cellules. On lui donne les cellules à additionner et il nous renvoie la somme.

Par exemple, la formule `=SOMME(A1:A5)` renvoie la somme des valeurs des cellules A1 à A5.

#### Moyenne

La fonction `MOYENNE` permet de calculer la moyenne des valeurs d'une plage de cellules.

La syntaxe est la suivante : `=MOYENNE(CL)` où `CL` est la référence de la plage de cellules.

Par exemple, la formule `=MOYENNE(A1:A5)` renvoie la moyenne des valeurs des cellules A1 à A5.

#### Minimum

La fonction `MIN` permet de renvoyer la valeur minimale d'une plage de cellules.

La syntaxe est la suivante : `=MIN(CL)` où `CL` est la référence de la plage de cellules.

Par exemple, la formule `=MIN(A1:A5)` renvoie la valeur minimale des valeurs des cellules A1 à A5.

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
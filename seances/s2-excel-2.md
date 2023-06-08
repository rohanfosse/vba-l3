<p style="text-align:left;">
    [Retour au sommaire](../README.md)
    <span style="float:right;">
        [Séance 3 - VBA: Fonctions et procédures, variables, tests et boucles](s3-vba-1.md)
    </span>
</p>

# Séance 2 - Excel: Fonction de bases - suite

Maintenant que nous avons revu le fonctionnement de base d'Excel, nous allons voir quelques fonctions qui permettent de manipuler les données.


## La fonction INDEX

La fonction **INDEX** permet de récupérer une valeur d'une cellule à partir de sa position dans un tableau.

La syntaxe de la fonction **INDEX** est la suivante:

```vb
INDEX(plage; ligne; colonne)
```

* **plage**: la plage de cellules dans laquelle on souhaite récupérer une valeur;
* **ligne**: la ligne de la cellule dont on souhaite récupérer la valeur;
* **colonne**: la colonne de la cellule dont on souhaite récupérer la valeur.

<div class="exemple">

La formule `=INDEX(A1:B3; 2; 2)` renvoie la valeur de la cellule B2.

</div>

La fonction **INDEX** peut aussi être utilisée pour récupérer une valeur à partir de sa position dans un tableau à une dimension. Dans ce cas, il faut utiliser la syntaxe suivante:

```vb
INDEX(plage; position)
```

* **plage**: la plage de cellules dans laquelle on souhaite récupérer une valeur;
* **position**: la position de la cellule dont on souhaite récupérer la valeur.

<div class="exemple">

La formule `=INDEX(A1:A6; 2)` renvoie la valeur de la cellule A2.

</div>


## La fonction EQUIV

La fonction **EQUIV** recherche un élément dans une plage de cellules et renvoie la position **relative** de cette cellule dans la plage.

La syntaxe de la fonction **EQUIV** est la suivante:

```vb
EQUIV(élément; plage;[type])
```

* **élément**: l'élément à rechercher dans la plage;
* **plage**: la plage de cellules dans laquelle on souhaite rechercher l'élément;
* **type**: cet argument est **optionnel**. Il prend la valeur -1, 0 ou 1:
    * -1: EQUIV recherche la valeur la plus petite qui est supérieure ou égale à celle de l’argument **élément**. Les valeurs de l’argument **plage** doivent être placées dans l’ordre **décroissant**;
    * 0: EQUIV recherche la première valeur qui est égale à celle de l’argument **élément**. Les valeurs de l’argument **plage** peuvent être placées dans n’importe quel ordre;
    * 1: EQUIV recherche la valeur la plus grande qui est inférieure ou égale à celle de l’argument **élément**. Les valeurs de l’argument **plage** doivent être placées dans l’ordre **croissant**.

<div class="exemple">

Prenons le tableau suivant:

| A | B | C |
|---|---|---|
| 1 | 2 | 3 |
| 4 | 5 | 6 |
| 7 | 8 | 9 |

La formule `=EQUIV(5;A1:C3;0)` renvoie la valeur 2 car la cellule B2 contient la valeur 5.

</div>

## Les fonctions RECHERCHEV et RECHERCHEH

Ces deux fonctions fonctionnent de la même manière. La seule différence est que **RECHERCHEV** recherche une valeur dans une colonne et **RECHERCHEH** recherche une valeur dans une ligne.

### La fonction RECHERCHEV

La fonction **RECHERCHEV** permet de rechercher une valeur dans une colonne et de renvoyer la valeur d'une cellule dans la même ligne mais dans une autre colonne. Cette fonction est très utile pour faire des tableaux de correspondance.

La syntaxe de la fonction **RECHERCHEV** est la suivante:

```vb
RECHERCHEV(élément; plage; colonne; [approximation])
```

* **élément**: l'élément à rechercher dans la plage;
* **plage**: la plage de cellules dans laquelle on souhaite rechercher l'élément;
* **colonne**: le numéro de la colonne dans laquelle on souhaite récupérer la valeur;
* **approximation**: cet argument est **optionnel**. Il prend la valeur VRAI ou FAUX:
    * VRAI: la fonction **RECHERCHEV** recherche une valeur approximative. Si la valeur recherchée n'est pas trouvée, la fonction renvoie la valeur la plus proche inférieure à la valeur recherchée;
    * FAUX: la fonction **RECHERCHEV** recherche une valeur exacte. Si la valeur recherchée n'est pas trouvée, la fonction renvoie la valeur d'erreur **#N/A**.

<div class="exemple">

Considérons le tableau suivant:

| A | B  | C  |
|---|----|----|
| 1 | 10 | 100|
| 3 | 20 | 200|
| 5 | 30 | 300|

**Exemple 1:**

Formule: `=RECHERCHEV(3;A1:C3;2;FAUX)`

Dans cet exemple, la formule cherche la valeur 3 dans la première colonne (A1:A3). Lorsqu'elle trouve le 3 (en ligne 2), elle renvoie la valeur correspondante dans la deuxième colonne du tableau (B1:B3), donc le résultat est 20.

**Exemple 2:**

Formule: `=RECHERCHEV(4;A1:C3;3;FAUX)`

Ici, la formule recherche la valeur 4 dans la première colonne (A1:A3). Comme il n'y a pas de 4 dans cette colonne et que l'option d'approximation est définie sur FAUX, la formule retourne #N/A, qui signifie qu'aucune valeur correspondante n'a été trouvée.

**Exemple 3:**

Formule: `=RECHERCHEV(4;A1:C3;3;VRAI)`

Dans cet exemple, la formule cherche la valeur 4 dans la première colonne (A1:A3). Comme l'option d'approximation est définie sur VRAI et qu'il n'y a pas de 4 exact dans la première colonne, Excel cherche la plus grande valeur qui est inférieure à 4. Dans ce cas, c'est 3. En conséquence, la formule renvoie la valeur de la troisième colonne (C1:C3) qui est sur la même ligne que 3, donc le résultat est 200.

</div>

### La fonction RECHERCHEH

De la même façon, la fonction **RECHERCHEH** permet de rechercher une valeur dans une ligne et de renvoyer la valeur d'une cellule dans la même colonne mais dans une autre ligne.

La syntaxe de la fonction **RECHERCHEH** est la suivante:

```vb
RECHERCHEH(élément; plage; ligne; [approximation])
```

* **élément**: l'élément à rechercher dans la plage;
* **plage**: la plage de cellules dans laquelle on souhaite rechercher l'élément;
* **ligne**: le numéro de la ligne dans laquelle on souhaite récupérer la valeur;
* **approximation**: cet argument est **optionnel**. Il prend la valeur VRAI ou FAUX:
    * VRAI: la fonction **RECHERCHEH** recherche une valeur approximative. Si la valeur recherchée n'est pas trouvée, la fonction renvoie la valeur la plus proche inférieure à la valeur recherchée;
    * FAUX: la fonction **RECHERCHEH** recherche une valeur exacte. Si la valeur recherchée n'est pas trouvée, la fonction renvoie la valeur d'erreur **#N/A**.

## Mélangeons tout ça!

### La fonction INDEX et EQUIV

La fonction **INDEX** permet de renvoyer la valeur d'une cellule dans une plage de cellules. La fonction **EQUIV** permet de renvoyer la position d'une valeur dans une plage de cellules. En combinant ces deux fonctions, on peut renvoyer la valeur d'une cellule en fonction de sa position dans une plage de cellules.

#### Quel est l'intérêt de combiner ces deux fonctions?

- Recherche bidirectionnelle : Contrairement à **RECHERCHEV** ou **RECHERCHEH**, qui recherchent uniquement verticalement ou horizontalement, la combinaison **INDEX-EQUIV** permet une recherche bidirectionnelle. C'est-à-dire qu'elle peut rechercher à la fois verticalement et horizontalement dans une table.

- Flexibilité : Avec **INDEX-EQUIV**, vous pouvez référencer des colonnes à gauche de la colonne de recherche, ce qui n'est pas possible avec **RECHERCHEV** ou **RECHERCHEH**.

- Précision : **INDEX-EQUIV** peut retourner une valeur exacte même si les données ne sont pas triées. Les fonctions **RECHERCHEV** et **RECHERCHEH**, quant à elles, doivent travailler avec des données triées pour renvoyer des résultats précis lors de l'utilisation de la recherche approximative.

#### Comment combiner ces deux fonctions?

On peut appeler les fonctions **INDEX** et **EQUIV** de deux manières différentes:

* **Méthode 1**: en utilisant la fonction **INDEX** comme argument de la fonction **EQUIV**:

```vb
EQUIV(élément; INDEX(plage; 0; 1); [approximation])
```

Avec cette méthode, la fonction **INDEX** renvoie la première colonne de la plage de cellules. La fonction **EQUIV** recherche ensuite l'élément dans cette première colonne et renvoie sa position.

L'objectif ici est donc de renvoyer la position d'une valeur dans une plage de cellules.

<div class="exemple">

**Exemple 1:**

Tableau:

| A | B  | C  |
|---|----|----|
| 1 | 10 | 100|
| 2 | 20 | 200|
| 3 | 30 | 300|

Formule : `=EQUIV(20; INDEX(B1:C3; 0; 1); 0)`

Ici, la fonction INDEX renvoie toute la première colonne de la plage B1:C3 (c'est-à-dire B1:B3). Ensuite, la fonction EQUIV cherche la valeur 20 dans cette plage et trouve qu'elle est à la deuxième position.

**Exemple 2:**

Tableau:

| A | B  | C  |
|---|----|----|
| 4 | 40 | 400|
| 5 | 50 | 500|
| 6 | 60 | 600|

Formule : `=EQUIV(6; INDEX(A1:C3; 0; 1); 0)`

Ici, la fonction INDEX renvoie toute la première colonne de la plage A1:C3 (c'est-à-dire A1:A3). Ensuite, la fonction EQUIV cherche la valeur 6 dans cette plage et trouve qu'elle est à la troisième position.

</div>

* **Méthode 2**: en utilisant la fonction **EQUIV** comme argument de la fonction **INDEX**:

```vb
INDEX(plage; EQUIV(élément; plage; [approximation]); 1)
```

Avec cette méthode, la fonction **EQUIV** renvoie la position de l'élément dans la plage de cellules. La fonction **INDEX** renvoie ensuite la valeur de la cellule qui se trouve à cette position.

L'objectif ici est donc de renvoyer la valeur d'une cellule en fonction de sa position dans une plage de cellules.

<div class="exemple">

**Exemple 3:**

Tableau:

| A | B  | C  |
|---|----|----|
| 1 | 10 | 100|
| 3 | 20 | 200|
| 5 | 30 | 300|

Formule : `=INDEX(B1:C3; EQUIV(3; A1:A3; 0); 2)`

Ici, la fonction EQUIV cherche la valeur 3 dans la plage A1:A3 et trouve qu'elle est à la deuxième position. Ensuite, la fonction INDEX renvoie la valeur de la deuxième ligne et deuxième colonne de la plage B1:C3, donc le résultat est 200.

**Exemple 4:**

Tableau:

| A | B  | C  |
|---|----|----|
| 7 | 40 | 400|
| 9 | 50 | 500|
| 11 | 60 | 600|

Formule : `=INDEX(B1:C3; EQUIV(11; A1:A3; 0); 2)`

Dans cet exemple, la fonction EQUIV cherche la valeur 11 dans la plage A1:A3 et trouve qu'elle est à la troisième position. Ensuite, la fonction INDEX renvoie la valeur de la troisième ligne et deuxième colonne de la plage B1:C3, donc le résultat est 600.

</div>

### La fonction INDEX et RECHERCHEV

La fonction **INDEX** permet de renvoyer la valeur d'une cellule dans une plage de cellules. La fonction **RECHERCHEV** permet de rechercher une valeur dans une colonne et de renvoyer la valeur d'une cellule dans la même ligne mais dans une autre colonne. En combinant ces deux fonctions, on peut renvoyer la valeur d'une cellule en fonction de sa position dans une plage de cellules.

#### Quel est l'intérêt de combiner ces deux fonctions?

La fonction **RECHERCHEV** permet de rechercher une valeur dans une colonne et de renvoyer la valeur d'une cellule dans la même ligne mais dans une autre colonne. Cependant, la fonction **RECHERCHEV** ne permet pas de choisir la colonne dans laquelle on veut rechercher la valeur. Elle recherche toujours la valeur dans la première colonne de la plage de cellules.

La fonction **INDEX** permet de renvoyer la valeur d'une cellule dans une plage de cellules. En utilisant la fonction **INDEX** comme argument de la fonction **RECHERCHEV**, on peut donc choisir la colonne dans laquelle on veut rechercher la valeur.

#### Comment combiner ces deux fonctions?

On peut appeler les fonctions **INDEX** et **RECHERCHEV** de deux manières différentes:

* **Méthode 1**: en utilisant la fonction **INDEX** comme argument de la fonction **RECHERCHEV**:

```vb
RECHERCHEV(élément; INDEX(plage; 0; 1); colonne; [approximation])
```

Avec cette méthode, la fonction **INDEX** renvoie la première colonne de la plage de cellules. La fonction **RECHERCHEV** recherche ensuite l'élément dans cette première colonne et renvoie la valeur de la cellule qui se trouve dans la même ligne mais dans la colonne indiquée.


<div class="exemple">

**Exemple 1:**

Tableau:

| A | B  | C  |
|---|----|----|
| 1 | 10 | 100|
| 2 | 20 | 200|
| 3 | 30 | 300|

Formule : `=RECHERCHEV(2; INDEX(A1:C3; 0; 1); 2; 0)`

Ici, la fonction INDEX renvoie toute la première colonne de la plage A1:C3 (c'est-à-dire A1:A3). Ensuite, la fonction RECHERCHEV cherche la valeur 2 dans cette plage et trouve qu'elle est à la deuxième position. La fonction RECHERCHEV renvoie donc la valeur de la deuxième ligne et deuxième colonne de la plage A1:C3, donc le résultat est 20.

</div>

* **Méthode 2**: en utilisant la fonction **RECHERCHEV** comme argument de la fonction **INDEX**:

```vb
INDEX(plage; RECHERCHEV(élément; plage; [approximation]); colonne)
```

Avec cette méthode, la fonction **RECHERCHEV** renvoie la position de l'élément dans la plage de cellules. La fonction **INDEX** renvoie ensuite la valeur de la cellule qui se trouve à cette position dans la colonne indiquée.

L'objectif ici est donc de renvoyer la valeur d'une cellule en fonction de sa position dans une plage de cellules.

<div class="exemple">

**Exemple 2:**

Tableau:

| A | B  | C  |
|---|----|----|
| 1 | 10 | 100|
| 3 | 20 | 200|
| 5 | 30 | 300|

Formule : `=INDEX(B1:C3; RECHERCHEV(3; A1:A3; 0); 2)`

Ici, la fonction RECHERCHEV cherche la valeur 3 dans la plage A1:A3 et trouve qu'elle est à la deuxième position. Ensuite, la fonction INDEX renvoie la valeur de la deuxième ligne et deuxième colonne de la plage B1:C3, donc le résultat est 200.

</div>

---

## Ce qu'il faut retenir

* La fonction **INDEX** permet de renvoyer la valeur d'une cellule dans une plage de cellules.
* La fonction **EQUIV** permet de rechercher une valeur dans une plage de cellules et de renvoyer sa position.
* La fonction **RECHERCHEV** permet de rechercher une valeur dans une colonne et de renvoyer la valeur d'une cellule dans la même ligne mais dans une autre colonne.
* La combinaison des fonctions **INDEX** et **EQUIV** permettent de renvoyer la valeur d'une cellule en fonction de sa position dans une plage de cellules.
* La combinaison des fonctions **INDEX** et **RECHERCHEV** permettent de renvoyer la valeur d'une cellule en fonction de sa position dans une plage de cellules.

---

<p style="text-align:left;">
    [Retour au sommaire](../README.md)
    <span style="float:right;">
        [Séance 3 - VBA: Fonctions et procédures, variables, tests et boucles](s3-vba-1.md)
    </span>
</p>
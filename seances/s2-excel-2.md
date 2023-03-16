<p style="text-align:left;">
    [Retour au sommaire](../README.md)
    <span style="float:right;">
        [Séance 3 - VBA: Fonctions et procédures, variables, tests et boucles](s3-vba-1.md)
    </span>
</p>

# Séance 2 - Excel: Fonction de bases - suite

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

La fonction **EQUIV** recherche un élément dans une plage de cellules et renvoie la position relative de cette cellule dans la plage.

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

To be continued
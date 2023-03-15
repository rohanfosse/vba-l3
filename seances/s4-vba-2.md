<p style="text-align:left;">
    [Retour au sommaire](../README.md)
    <span style="float:right;">
        [Séance 5 - VBA: Tableaux et Enregistrements](s5-vba-3.md)
    </span>
</p>

<div style="text-align:center;">

# Séance 4 - Structures Iteratives et communications avec Excel

</div>

---
## Avant de commencer <a name="avant-de-commencer"></a>

#### Une itération

Lorsque l'on répète plusieurs fois les mêmes actions, on parle **d'itération**.

Une **structure d'itération** permet de rejouer les mêmes actions, avec d'éventuelles petites différences.

Les **structures itératives** fournissent un moyen d'effectuer des boucles sur des instructions : **la boucle** permet d'exécuter des itérations.

Il existe plusieurs types de structures itératives, mais elles sont généralement **communes** entre les différents langages.

<div class="line"></div>

#### Condition d'arrêt

Une boucle s'exécute un certain nombre de fois avant de s'interrompre et que la suite du programme poursuive son exécution. Si une boucle ne s’interrompt jamais, c'est une **boucle infinie** : le programme reste bloqué car la boucle se répète indéfiniment.

Les structures itératives nécessitent donc **une condition d'arrêt**, c'est-à-dire une condition qui interrompt les itérations dès qu'elle est remplie.

<div class="line"></div>

#### Compteur

Un **compteur** est souvent utilisé à l'intérieur de la boucle : une variable entière, généralement initialisée à 0, est incrémentée à chaque nouvelle itération. Le compteur permet ainsi simplement de compter le nombre d'itérations déjà effectué.

La valeur du compteur est très souvent utilisée dans la condition de sortie, pour interrompre la boucle au bout d'un certain nombre d'itération.

---

## Les boucles <a name="les-boucles"></a>

### Boucle **While** <a name="boucle-while"></a>

La boucle **While** permet d'exécuter une instruction tant qu'une condition est vraie.

La syntaxe est la suivante :

```vb
While condition
    'instruction'
Wend
```



<div class="exemple">

Si l'on souhaite afficher les nombres de 1 à 10, le code sera le suivant :

```vb
Dim i As Integer
i = 1

While i <= 10
    MsgBox i
    i = i + 1
Wend
```

</div>

<div class="line"></div>

Un autre exemple serait de demander à l'utilisateur de saisir un nombre. Tant que ce nombre n'est pas compris entre 1 et 10, on lui redemande.

<div class ="exemple">

```vb
Sub boucle_while()
    Dim nombre As Integer

    nombre = InputBox("Entrez un nombre", "Nombre")

    While nombre < 1 Or nombre > 10
        nombre = InputBox("Entrez un nombre", "Nombre")
    Wend
End Sub
```

Dans cet exemple, on peut voir que l'on doit écrire deux fois le code

`nombre = InputBox("Entrez un nombre", "Nombre")`

En effet, il faut d'abord que l'utilisateur saisisse un nombre, puis que le programme vérifie si ce nombre est compris entre 1 et 10.

Si ce n'est pas le cas, on redemande à l'utilisateur de saisir un nombre.

</div>

<div class="line"></div>

### Boucle **Do While** <a name="boucle-do-while"></a>

La boucle **Do While** permet d'exécuter une instruction tant qu'une condition est vraie.

Elle est similaire à la boucle **While** mais la condition est testée **à la fin** de l'exécution de l'instruction. La syntaxe est la suivante :

```vb
Do
    'instruction'
Loop While condition
```

A la différence de la boucle **While** , les instructions sont exécutées **au moins une fois**.

<div class="exemple">

Pour afficher les nombres de 1 à 10, le code sera le suivant :

```vb
Dim i As Integer
i = 1

Do
    MsgBox i
    i = i + 1
Loop While i <= 10
```

</div>

Si nous reprenons le même exemple que la boucle **While** , on obtient le code suivant :

<div class="exemple">

```vb
Dim nombre As Integer

Do
    nombre = InputBox("Entrez un nombre", "Nombre")
Loop While nombre < 1 Or nombre > 10
```

</div>

<div class="line"></div>

### Boucle **Do Loop Until** <a name="boucle-do-until"></a>

La boucle **Do Loop Until** permet d'exécuter une instruction tant qu'une condition est **fausse**. Elle est similaire à la boucle **Do While** mais la condition est testée à la fin de l'exécution de l'instruction. La syntaxe est la suivante :

```vb
Do
    instruction
Loop Until condition
```

<div class="exemple">

Pour afficher les nombres de 1 à 10, le code sera le suivant :

```vb
Dim i As Integer
i = 1

Do
    MsgBox i
    i = i + 1
Loop Until i > 10
```

</div>

Si nous reprenons le même exemple que la boucle **While**, on obtient le code suivant :

<div class="exemple">

```vb
Sub do_while()
    Dim nombre As Integer

    Do
        nombre = InputBox("Entrez un nombre", "Nombre")
    Loop Until nombre >= 1 And nombre <= 10
End Sub
```

</div>

<div class="line"></div>

### Boucle **For** <a name="boucle-for"></a>

La boucle **For** permet d'exécuter une instruction un nombre défini de fois.

La syntaxe est la suivante :

```vb
For i = valeur_de_depart To valeur_de_fin Step pas
    'instruction'
Next i
```

*Step* sert à définir le pas d'itération. Il est toutefois optionnel. Si on ne le précise pas, le pas vaut 1.

***Next i*** permet de terminer la boucle. Il est important de préciser ***i*** car il s'agit du nom de la variable de boucle.


<div class="exemple">

Pour afficher les nombres de 1 à 10, le code sera le suivant :

```vb
For i = 1 To 10 
    MsgBox i
Next i
```

</div>

On peut aussi utiliser la boucle **For** pour parcourir un tableau.

<div class="exemple">

Pour afficher les valeurs d'un tableau **tableau** de taille 3, le code sera le suivant :

```vb
Dim tableau(2) As Integer
tableau(0) = 1
tableau(1) = 2
tableau(2) = 3

For i = 0 To 2
    MsgBox tableau(i)
Next i
```

</div>

Une autre façon de définir le tableau est la suivante :

<div class="exemple">

```vb
Dim tableau() As Variant
tableau = Array(1, 2, 3)

For i = 0 To 2
    MsgBox tableau(i)
    Next i
```

</div>

<div class="line"></div>

### Boucle For Each In Next <a name="boucle-for-each-in-next"></a>

La boucle **For Each In Next** permet d'exécuter une instruction pour chaque élément d'un tableau.

La syntaxe est la suivante :

```vb
For Each element In tableau
    'instruction'
Next element
```

Pour afficher les valeurs d'un tableau **tableau** de taille 3, le code sera le suivant:

<div class="exemple">

```vb
Dim tableau(2) As Integer
tableau(0) = 1
tableau(1) = 2
tableau(2) = 3

For Each element In tableau
    MsgBox element
Next element
```

</div>

---

## Communication avec Excel <a name="communication-avec-excel"></a>

#### Range <a name="range"></a>

Le type **Range** permet de manipuler des cellules ou des plages de cellules.

Pour créer un objet **Range** , on utilise la syntaxe suivante :

<div class="exemple">

```vb
Dim range As Range
Set range = Range("A1")
```

</div>

On peut aussi créer un objet **Range** à partir d'une plage de cellules :

<div class="exemple">

```vb
Dim range As Range
Set range = Range("A1:B2")
```

</div>

On peut aussi créer un objet **Range** à partir d'une plage de cellules en utilisant les coordonnées :

<div class="exemple">

```vb
Dim range As Range
Set range = Range(Cells(1, 1), Cells(2, 2))
```

</div>

La méthode **Clear** permet de supprimer le contenu d'une cellule ou d'une plage de cellules :

<div class="exemple">

```vb
Dim range As Range
Set range = Range("A1:B2")
range.Clear
```

</div>

La méthode **Value** permet de récupérer la valeur d'une cellule ou d'une plage de cellules :

<div class="exemple">

```vb
Dim range As Range
Set range = Range("A1:B2")
MsgBox range.Value
```

</div>

La méthode **Cells** permet de spécifier une cellule à partir d'une plage de cellules :

<div class="exemple">

```vb
Dim range As Range
Set range = Range("A1:B2")
range.Cells(1, 1).Value = 1
range.Cells(1, 2).Value = 2
range.Cells(2, 1).Value = 3
range.Cells(2, 2).Value = 4
```

</div>

#### Application <a name="application"></a>

L'objet **Application** permet de manipuler Excel.

Pour créer un objet **Application** , on utilise la syntaxe suivante :

```vb
Dim app As Application
Set app = Application
```

La méthode **Run** permet d'exécuter une macro :

<div class="exemple">

```vb
Dim app As Application
Set app = Application
app.Run "NomDeLaMacro"
```

</div>

La méthode **Run** permet aussi d'exécuter une macro avec des paramètres :

<div class="exemple">

```vb
Dim app As Application
Set app = Application
app.Run "NomDeLaMacro", "param1", "param2"
```

</div>

---

## Exercices Corrigés <a name="-exercices-corriges-4"></a>

#### Exercice 1 <a name="exercice-1-4"></a>

<div class="exercice">
Ecrire une fonction qui demande à l'utilisateur un entier **n** et fait la somme des entiers de **1 à n**.

Puis, ecrire une procédure qui affiche le résultat de la fonction.

##### Correction


<details>

Tout d'abord, comme on connait le nombre d'itérations, on sait que l'on peut utiliser une boucle **For**.

Ensuite, on crée une variable **somme** qui contiendra la somme des entiers de 1 à **n**.

Enfin, on fait la somme des entiers de 1 à **n** en ajoutant la valeur de **i** à la variable **somme** à chaque itération.

A la fin de la boucle, on renvoie la valeur de **somme** en oubliant pas de l'affecter à **sommeEntiers** (le nom de la fonction).

```vb
Function sommeEntiers() As Integer
    Dim n As Integer
    Dim somme As Integer
    Dim i As Integer

    n = InputBox("Entrez un entier")
    somme = 0

    For i = 1 To n
        somme = somme + i
    Next i

    sommeEntiers = somme
End Function

Sub afficherSomme()
    MsgBox sommeEntiers()
End Sub
```

</details>
</div>

<div class="line"></div>

#### Exercice 2 <a name="exercice-2-4"></a>

<div class="exercice">

Ecrire une fonction **double_tableau** qui prend en paramètre un tableau d'entiers **tab** et sa taille **n**. La fonction renvoie un booléen indiquant si le tableau ne contient que des entiers pairs.

Par exemple, si **tab** contient les valeurs **1, 2, 3, 4**, la fonction renvoie **False** mais renvoie **True** si **tab** contient les valeurs **2, 4, 6, 8**.

Vous pouvez utiliser l'opérateur **Mod** pour calculer le modulo. Par exemple, **5 Mod 2** renvoie **1**.

##### Correction

<details>

```vb
Function double_tableau(tableau() As Integer, n As Integer) As Boolean
    Dim i As Integer
    Dim resultat As Boolean

    resultat = True
    i = 0

    Do While i < n And resultat = True
        If tableau(i) Mod 2 <> 0 Then
            resultat = False
        End If
        i = i + 1
    Loop

    double_tableau = resultat
End Function
```

</details>
</div>

<div class="line"></div>

#### Exercice 3 <a name="exercice-3-4"></a>

<div class="exercice">

Ecrire une fonction **fibonacci** qui prend en paramètre un entier **n** et retourne un tableau contenant les **n** premiers termes de la suite de Fibonacci. Le type de retour de la fonction est **Variant** car on ne connait pas à l'avance la taille du tableau.

Pour rappel, la suite de Fibonacci est définie par la relation suivante :

```vb
f(0) = 0
f(1) = 1
f(n) = f(n-1) + f(n-2)
```

##### Correction

<details>

```vb
Function fibonacci(n As Integer) As Variant
    Dim tableau(n) As Integer
    Dim i As Integer

    tableau(0) = 0
    tableau(1) = 1

    For i = 2 To n
        tableau(i) = tableau(i - 1) + tableau(i - 2)
    Next i

    fibonacci = tab
End Function
```

</details>
</div>

---

## Correction TD4 <a name="correction-td4-4"></a>

Une correction du TD4 sera disponible prochainement.

<p style="text-align:left;">
    [Retour au sommaire](../README.md)
    <span style="float:right;">
        [Séance 5 - VBA: Tableaux et Enregistrements](s4-vba-3.md)
    </span>
</p>

<div style="text-align:center;">

# Séance 4 - Structures Iteratives et communications avec Excel

</div>

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

Un autre exemple serait de demander à l'utilisateur de saisir un nombre tant que ce nombre n'est pas compris entre 1 et 10.

<div class ="exemple">

```vb
Dim nombre As Integer

nombre = InputBox("Entrez un nombre", "Nombre")

While nombre < 1 Or nombre > 10
    nombre = InputBox("Entrez un nombre", "Nombre")
Wend
```

Dans cet exemple, on doit écrire deux fois le code `nombre = InputBox("Entrez un nombre", "Nombre")`. En effet, il faut d'abord que l'utilisateur saisisse un nombre, puis que le programme vérifie si ce nombre est compris entre 1 et 10.
Si ce n'est pas le cas, on redemande à l'utilisateur de saisir un nombre.

</div>

### Boucle **Do While** <a name="boucle-do-while"></a>

La boucle **Do While** permet d'exécuter une instruction tant qu'une condition est vraie. Elle est similaire à la boucle **While** mais la condition est testée **à la fin** de l'exécution de l'instruction.

La syntaxe est la suivante :

```vb
Do
    'instruction'
Loop While condition
```

La différence avec la boucle **While** , c'est que les instructions sont exécutées au moins une fois.

<div class="exemple">

Par exemple, pour afficher les nombres de 1 à 10, on utilise la syntaxe suivante :

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

```vb
Dim nombre As Integer

Do
    nombre = InputBox("Entrez un nombre", "Nombre")
Loop While nombre < 1 Or nombre > 10
```

### Boucle **Do Loop Until** <a name="boucle-do-until"></a>

La boucle **Do Loop Until** permet d'exécuter une instruction tant qu'une condition est **fausse**. Elle est similaire à la boucle **Do While** mais la condition est testée à la fin de l'exécution de l'instruction.

La syntaxe est la suivante :

```vb
Do
    instruction
Loop Until condition
```

Par exemple, pour afficher les nombres de 1 à 10, on utilise la syntaxe suivante :

```vb
Dim i As Integer
i = 1

Do
    MsgBox i
    i = i + 1
Loop Until i > 10
```

Si nous reprenons le même exemple que la boucle **While`, on obtient le code suivant :

```vb
Dim nombre As Integer

nombre = InputBox("Entrez un nombre", "Nombre")

Do
    nombre = InputBox("Entrez un nombre", "Nombre")
Loop Until nombre >= 1 And nombre <= 10
```

### Boucle **For** <a name="boucle-for"></a>

La boucle **For** permet d'exécuter une instruction un certain nombre de fois.

La syntaxe est la suivante :

```vb
For i = valeur_de_depart To valeur_de_fin Step pas
    'instruction'
Next i
```

*Step* sert à définir le pas d'itération. Il est toutefois optionnel. Si on ne le précise pas, le pas vaut 1.

Par exemple, pour afficher les nombres de 1 à 10, on utilise la syntaxe suivante :

```vb
For i = 1 To 10 
    MsgBox i
Next i
```

On peut aussi utiliser la boucle **For** pour parcourir un tableau. Par exemple, pour afficher les valeurs d'un tableau **tab** de taille 3, on utilise la syntaxe suivante :

```vb
Dim tab(2) As Integer
tab(0) = 1
tab(1) = 2
tab(2) = 3

For i = 0 To 2
    MsgBox tab(i)
Next i
```

Une autre façon de définir le tableau est la suivante :

```vb
Dim tab() As Variant
tab = Array(1, 2, 3)

For i = 0 To 2
    MsgBox tab(i)
    Next i
```

#### Boucle **For Each In Next** <a name="boucle-for-each-in-next"></a>

La boucle **For Each In Next** permet d'exécuter une instruction pour chaque élément d'un tableau.

La syntaxe est la suivante :

```vb
For Each element In tableau
    'instruction'
Next element
```

Par exemple, pour afficher les valeurs d'un tableau **tab** de taille 3, on utilise la syntaxe suivante :

```vb
Dim tab(2) As Integer
tab(0) = 1
tab(1) = 2
tab(2) = 3

For Each element In tab
    MsgBox element
Next element
```

---

## Communication avec Excel <a name="communication-avec-excel"></a>

### Range <a name="range"></a>

Le type **Range** permet de manipuler des cellules ou des plages de cellules.

Pour créer un objet **Range** , on utilise la syntaxe suivante :

```vb
Dim range As Range
Set range = Range("A1")
```

On peut aussi créer un objet **Range** à partir d'une plage de cellules :

```vb
Dim range As Range
Set range = Range("A1:B2")
```

On peut aussi créer un objet **Range** à partir d'une plage de cellules en utilisant les coordonnées :

```vb
Dim range As Range
Set range = Range(Cells(1, 1), Cells(2, 2))
```

La méthode **Clear** permet de supprimer le contenu d'une cellule ou d'une plage de cellules :

```vb
Dim range As Range
Set range = Range("A1:B2")
range.Clear
```

La méthode **Value** permet de récupérer la valeur d'une cellule ou d'une plage de cellules :

```vb
Dim range As Range
Set range = Range("A1:B2")
MsgBox range.Value
```

La méthode **Cells** permet de spécifier une cellule à partir d'une plage de cellules :

```vb
Dim range As Range
Set range = Range("A1:B2")
range.Cells(1, 1).Value = 1
range.Cells(1, 2).Value = 2
range.Cells(2, 1).Value = 3
range.Cells(2, 2).Value = 4
```

### Application <a name="application"></a>

L'objet **Application** permet de manipuler Excel.

Pour créer un objet **Application** , on utilise la syntaxe suivante :

```vb
Dim app As Application
Set app = Application
```

La méthode **Run** permet d'exécuter une macro :

```vb
Dim app As Application
Set app = Application
app.Run "NomDeLaMacro"
```

La méthode **Run** permet aussi d'exécuter une macro avec des paramètres :

```vb
Dim app As Application
Set app = Application
app.Run "NomDeLaMacro", "param1", "param2"
```

---

## Exercices Corrigés <a name="-exercices-corriges-4"></a>

#### Exercice 1 <a name="exercice-1-4"></a>

<div class="exercice">
Ecrire une fonction qui demande à l'utilisateur un entier **n** et fait la somme des entiers de 1 à **n** .
Ecrire une procédure qui affiche le résultat de la fonction.

##### Correction

<details>

```vb
Function sommeEntiers() As Integer
    Dim nombre As Integer
    Dim somme As Integer
    Dim i As Integer

    n = InputBox("Entrez un nombre", "Nombre")
    somme = 0
    i = 1

    Do While i <= n
        somme = somme + i
        i = i + 1
    Loop

    sommeEntiers = somme
End Function

Sub afficherSomme()
    MsgBox sommeEntiers()
End Sub
```

</details>
</div>

#### Exercice 2 <a name="exercice-2-4"></a>

<div class="exercice">

Ecrire une fonction **`double_tableau`** qui prend en paramètre un tableau d'entiers **tab** et sa taille **n**. La fonction renvoie un booléen indiquant si le tableau ne contient que des entiers pairs.

Par exemple, si **tab** contient les valeurs **1, 2, 3, 4**, la fonction renvoie **False** mais renvoie **True** si **tab** contient les valeurs **2, 4, 6, 8**.

Vous pouvez utiliser l'opérateur **Mod** pour calculer le modulo. Par exemple, **5 Mod 2** renvoie **1**.

##### Correction

<details>

```vb
Function double_tableau(tab() As Integer, n As Integer) As Boolean
    Dim i As Integer
    Dim resultat As Boolean

    resultat = True
    i = 0

    Do While i < n And resultat = True
        If tab(i) Mod 2 <> 0 Then
            resultat = False
        End If
        i = i + 1
    Loop

    double_tableau = resultat
End Function
```

</details>
</div>

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
    Dim tab(n) As Integer
    Dim i As Integer

    tab(0) = 0
    tab(1) = 1

    For i = 2 To n
        tab(i) = tab(i - 1) + tab(i - 2)
    Next i

    fibonacci = tab
End Function
```

</details>
</div>

---

## Correction TD4 <a name="correction-td4-4"></a>

Une correction du TD4 sera disponible prochainement.
